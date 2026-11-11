import os, re, sys, time, smtplib, datetime, logging
from collections import OrderedDict
from email.message import EmailMessage
import csv, io, urllib.request, urllib.parse
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

os.chdir(os.path.dirname(os.path.abspath(__file__)))
load_dotenv()
logging.basicConfig(filename="timesheet.log", level=logging.INFO,
                    format="%(asctime)s %(message)s")

USER = os.environ.get("TS_USER", "")
PASS = os.environ.get("TS_PASS", "")
TASK_SOURCE_URL = os.environ.get("TASK_SOURCE_URL", "").strip()
TASK_SHEET_NAME = os.environ.get("TASK_SHEET_NAME", "Daily Tasks")
TS_DISPLAY_NAME = os.environ.get("TS_DISPLAY_NAME", "").strip()
LOGGED_IN_NAME = TS_DISPLAY_NAME or USER
MAIL_USER = os.environ.get("MAIL_USER", "")
MAIL_PASS = os.environ.get("MAIL_PASS", "")
MAIL_TO = os.environ.get("MAIL_TO", "")

BASE = "https://bsoftagile.boscosofttech.com/"
TODAY_D = datetime.date.today()
TODAY = TODAY_D.strftime("%d/%m/%Y")

# ---- Working-day rules -------------------------------------------------
WEEKDAY_HOURS = 8            # Monday to Friday
SATURDAY_HOURS = 6           # working Saturdays only
WORKING_SATURDAYS = (2, 4)   # which Saturdays of the month are working days
# ------------------------------------------------------------------------

# Modes:
#   python timesheet.py          visible browser, SAVES, checks tracking, emails
#   python timesheet.py dry      visible browser, fills only (no save, no email)
#   python timesheet.py hide     no window (for the scheduler), SAVES, emails
#   python timesheet.py track    only opens the tracking page and saves what it sees
#   python timesheet.py mailtest only sends a test email
MODE = sys.argv[1] if len(sys.argv) > 1 else "run"
SLOW_MO = 0
END_PAUSE_MS = 2000
BLOCK_HEAVY = True
T0 = time.time()

REPORT = {"tasks": [], "shots": [], "tracking": None, "error": None}


def log(msg):
    line = f"[{time.time() - T0:5.1f}s] {msg}"
    print(line)
    logging.info(msg)


# ---------------------------------------------------------------- day rules
def day_plan():
    """Returns (entries_file, expected_hours, label) or None if not a working day."""
    wd = TODAY_D.weekday()
    if wd <= 4:
        return "daily.txt", WEEKDAY_HOURS, "Weekday"
    if wd == 5:
        nth = (TODAY_D.day - 1) // 7 + 1
        if nth in WORKING_SATURDAYS:
            return "saturday.txt", SATURDAY_HOURS, f"Saturday #{nth}"
        return None
    return None




def normalize_task_row(row):
    def get(*names, default=""):
        for name in names:
            value = row.get(name)
            if value is not None and str(value).strip() and str(value).strip().lower() != "nan":
                return str(value).strip()
        return default

    project_code = get("project code", "project id", "code", "project")
    project_name = get("project name", "project title", default=project_code)
    discipline = get("discipline", "category")
    desc = get("activity description", "description", "task description", "activity")
    hours = get("actual hours", "hours", "hrs", "time")
    status = get("status", default="Completed")
    date_value = get("date", "work date")

    if not any((date_value, project_code, discipline, desc, hours)):
        return None
    if not date_value:
        raise ValueError("Date is required for every task row. Use DD-MM-YYYY.")

    date_text = date_value.strip()
    if " " in date_text and "T" not in date_text:
        date_text = date_text.split(" ")[0]
    row_date = None
    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%b-%Y", "%d %b %Y"):
        try:
            row_date = datetime.datetime.strptime(date_text, fmt).date()
            break
        except ValueError:
            continue
    if row_date is None:
        raise ValueError(f"Invalid date '{date_value}'. Use DD-MM-YYYY.")
    if row_date != TODAY_D:
        return None

    if not all((project_code, discipline, desc, hours)):
        raise ValueError(
            "Today's task row needs Project Code, Discipline, Activity Description and Actual Hours."
        )
    try:
        number = float(hours)
    except ValueError:
        raise ValueError(f"Hours must be numeric: {hours}")
    if number <= 0 or number > 24:
        raise ValueError(f"Hours must be greater than 0 and at most 24: {hours}")

    return {
        "project": project_code,          # stable unique key
        "project_code": project_code,
        "project_name": project_name,     # display/dropdown name
        "discipline": discipline,
        "desc": desc,
        "hours": f"{number:g}",
        "status": status
    }


def read_csv_bytes(raw):
    reader = csv.DictReader(io.StringIO(raw.decode("utf-8-sig")))
    return [normalize_task_row({str(k).strip().lower(): v for k, v in row.items()})
            for row in reader]


def read_xlsx_bytes(raw):
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(raw), data_only=True, read_only=True)
    sheet = TASK_SHEET_NAME if TASK_SHEET_NAME in wb.sheetnames else wb.sheetnames[0]
    ws = wb[sheet]
    iterator = ws.iter_rows(values_only=True)
    headers = next(iterator, None)
    if not headers:
        return []
    keys = [str(h or "").strip().lower() for h in headers]
    entries = []
    for values in iterator:
        row = {keys[i]: values[i] for i in range(min(len(keys), len(values))) if keys[i]}
        item = normalize_task_row(row)
        if item:
            entries.append(item)
    wb.close()
    return entries


def read_entries(source):
    if source.startswith(("http://", "https://")):
        url = source
        if "docs.google.com/spreadsheets/d/" in url:
            match = re.search(r"/spreadsheets/d/([A-Za-z0-9_-]+)", url)
            gid = re.search(r"[?&]gid=(\\d+)", url)
            if match:
                url = f"https://docs.google.com/spreadsheets/d/{match.group(1)}/export?format=csv"
                if gid:
                    url += f"&gid={gid.group(1)}"
        else:
            parsed = urllib.parse.urlparse(url)
            query = urllib.parse.parse_qs(parsed.query)
            query["download"] = ["1"]
            url = urllib.parse.urlunparse(parsed._replace(
                query=urllib.parse.urlencode(query, doseq=True)))
        request = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(request, timeout=30) as response:
            raw = response.read()
            content_type = response.headers.get("Content-Type", "").lower()
        if not raw or b"<html" in raw[:500].lower() or b"<!doctype html" in raw[:500].lower():
            raise ValueError(
                "The URL returned a login/web page, not spreadsheet data. "
                "Share a public direct-download Excel link or Google Sheets viewer link."
            )
        entries = read_xlsx_bytes(raw) if raw[:2] == b"PK" else read_csv_bytes(raw)
    else:
        file_path = Path(source)
        if not file_path.exists():
            raise FileNotFoundError(f"Task spreadsheet not found: {source}")
        raw = file_path.read_bytes()
        entries = read_xlsx_bytes(raw) if file_path.suffix.lower() != ".csv" else read_csv_bytes(raw)

    entries = [entry for entry in entries if entry]
    # Empty means there is no row for today's date; main() emails a reminder.
    return entries


# ---------------------------------------------------------------- JS snippets
PICK_JS = """
([sel, text]) => {
  const els = document.querySelectorAll(sel);
  if (!els.length) throw new Error('Field not found: ' + sel);
  const el = els[els.length - 1];
  const opt = Array.from(el.options).find(
      o => o.text.toLowerCase().includes(text.toLowerCase()));
  if (!opt) throw new Error('Option not found: ' + text);
  el.value = opt.value;
  el.dispatchEvent(new Event('change', {bubbles: true}));
  if (window.jQuery) window.jQuery(el).trigger('change');
}
"""

CURRENT_PROJECT_JS = """
() => {
  const el = document.querySelector('#select2basic');
  return el ? el.options[el.selectedIndex].text : 'unknown';
}
"""

CARDS_JS = """
() => {
  const rx = /(\\d{2}\\/\\d{2}\\/\\d{4})\\s*-\\s*(\\d{2}\\/\\d{2}\\/\\d{4})/;
  const sel = "a[href*='go-to-tabs']";
  const out = [];
  document.querySelectorAll(sel).forEach(a => {
    let el = a;
    while (el.parentElement && el.parentElement.querySelectorAll(sel).length === 1)
      el = el.parentElement;
    const txt = el.innerText;
    const m = txt.match(rx);
    const t = txt.match(/(\\d{5,7})\\s*:\\s*([^\\n]+)/);
    out.push({href: a.getAttribute("href"),
              start: m ? m[1] : null, end: m ? m[2] : null,
              title: t ? t[0] : "(unknown)"});
  });
  return out;
}
"""

CLICK_HREF_JS = """
(h) => {
  const a = Array.from(document.querySelectorAll("a")).find(x => x.getAttribute("href") === h);
  if (!a) throw new Error("card link not found");
  a.click();
}
"""

ROWS_JS = "() => Array.from(document.querySelectorAll('tr')).map(r => r.innerText)"

FORM_JS = """
() => {
  const out = ["URL: " + location.href];
  document.querySelectorAll("input, select, textarea, button").forEach(e => {
    const r = e.getBoundingClientRect();
    out.push(e.tagName + " | type: " + e.type + " | id: " + e.id +
             " | value: " + (e.value || "").slice(0, 40) +
             " | visible: " + (r.width > 0 && r.height > 0));
  });
  return out.join("\\n");
}
"""


# ---------------------------------------------------------------- helpers
def settle(page):
    try:
        page.wait_for_load_state("domcontentloaded", timeout=10000)
    except Exception:
        pass


def safe_eval(page, js, arg=None, tries=12):
    last = None
    for _ in range(tries):
        try:
            if arg is not None:
                return page.evaluate(js, arg)
            return page.evaluate(js)
        except Exception as ex:
            last = ex
            page.wait_for_timeout(250)
    raise last


def body_text(page):
    return safe_eval(page, "() => document.body.innerText").lower()


def dump_debug(page, name="debug_page"):
    try:
        text = page.evaluate("() => location.href + '\\n\\n' + document.body.innerText")
    except Exception:
        text = "could not read page"
    with open(name + ".txt", "w", encoding="utf-8") as f:
        f.write(text)
    try:
        page.screenshot(path=name + ".png", full_page=True)
    except Exception:
        pass


def dump_form(page):
    try:
        text = page.evaluate(FORM_JS)
        with open("form_debug.txt", "w", encoding="utf-8") as f:
            f.write(text)
        page.screenshot(path="form_debug.png", full_page=True)
    except Exception:
        pass


def block_heavy(route):
    if route.request.resource_type in ("image", "font", "media"):
        route.abort()
    else:
        route.continue_()



# ----------------------------------------------------------- screenshots
DESKTOP_WIDTH = 1920
DESKTOP_HEIGHT = 1200


def prepare_desktop_capture(page):
    """
    Normalize browser zoom and use a real desktop-sized CSS viewport.
    Do not force every application div to 100%: that can break the site's layout.
    """
    page.evaluate("""() => {
        document.documentElement.style.setProperty('zoom', '1', 'important');
        document.body.style.setProperty('zoom', '1', 'important');
        document.documentElement.style.setProperty('transform', 'none', 'important');
        document.documentElement.style.setProperty('width', '100%', 'important');
        window.scrollTo(0, 0);
    }""")
    page.wait_for_timeout(500)
    metrics = page.evaluate("""() => ({
        innerWidth: window.innerWidth,
        innerHeight: window.innerHeight,
        devicePixelRatio: window.devicePixelRatio,
        documentWidth: document.documentElement.scrollWidth,
        bodyWidth: document.body.scrollWidth
    })""")
    log("Screenshot viewport: " + str(metrics))


def scroll_cell_into_view(page, cell):
    """Scroll both vertically and horizontally so a table cell is visible."""
    page.evaluate("""el => {
        const ancestors = [];
        let p = el.parentElement;
        while (p && p !== document.body) {
            const cs = getComputedStyle(p);
            if (p.scrollWidth > p.clientWidth + 5 &&
                /(auto|scroll)/.test(cs.overflowX)) {
                ancestors.push(p);
            }
            p = p.parentElement;
        }
        // Scroll the innermost horizontal scroller first.
        ancestors.forEach(scroller => {
            const er = el.getBoundingClientRect();
            const sr = scroller.getBoundingClientRect();
            if (er.left < sr.left) scroller.scrollLeft -= (sr.left - er.left) + 30;
            else if (er.right > sr.right) scroller.scrollLeft += (er.right - sr.right) + 30;
        });
        el.scrollIntoView({block: 'center', inline: 'center', behavior: 'instant'});
    }""", cell)
    page.wait_for_timeout(350)


def take_clear_screenshot(page, path, target_selector=None):
    """
    Capture at desktop resolution. For ordinary task pages preserve the page's
    layout and full vertical content. Avoid blindly expanding every div.
    """
    try:
        prepare_desktop_capture(page)
        if target_selector:
            target = page.locator(target_selector).first
            if target.count():
                target.scroll_into_view_if_needed(timeout=5000)
        page.evaluate("() => window.scrollTo(0, 0)")
        page.wait_for_timeout(250)
        page.screenshot(
            path=path,
            full_page=True,
            animations="disabled",
            type="png"
        )
        log(f"Screenshot saved: {path}")
    except Exception as ex:
        log(f"Full screenshot fallback for {path}: {ex}")
        page.screenshot(path=path, full_page=False, animations="disabled")


def scroll_tracking_to_today(page):
    """
    Find today's date in tracking table headers/cells and horizontally scroll
    the nested table container until that date is visible.
    """
    variants = date_variants(TODAY_D)
    candidates = page.locator("th, td")
    for i in range(candidates.count()):
        cell = candidates.nth(i)
        try:
            text_value = " ".join(cell.inner_text(timeout=700).split())
            normalized = text_value.lower()
            if any(v.lower() in normalized for v in variants):
                scroll_cell_into_view(page, cell)
                log(f"Tracking date column brought into view: {text_value}")
                return True
        except Exception:
            continue

    # Fallback: search all elements in case the date is rendered outside a table.
    for value in variants:
        locator = page.get_by_text(value, exact=False)
        try:
            if locator.count():
                locator.first.scroll_into_view_if_needed(timeout=3000)
                page.wait_for_timeout(300)
                return True
        except Exception:
            pass
    log(f"Today's date {TODAY} was not found in visible tracking table cells")
    return False


def save_tracking_screenshots(page):
    """
    Save exactly one tracking screenshot.
    The current-date column is brought into view before capture; use a viewport
    screenshot (not full_page) so the table remains readable and is not shrunk.
    """
    prepare_desktop_capture(page)
    found = scroll_tracking_to_today(page)

    # Keep the tracking table/date column in the visible desktop viewport.
    page.wait_for_timeout(300)
    page.screenshot(
        path="tracking.png",
        full_page=False,
        animations="disabled",
        type="png"
    )

    # The tracking page is common to all projects, so attach it only once.
    REPORT["shots"] = [("tracking.png", "tracking.png")]
    log("Saved the single tracking screenshot: tracking.png")
    return found


def build_report_html():
    tracking = REPORT["tracking"] or {"status": "NOT CHECKED", "row": ""}
    total = sum(float(e["hours"]) for tk in REPORT["tasks"] for e in tk["entries"])
    project_totals = OrderedDict()
    for tk in REPORT["tasks"]:
        project_totals[tk["project"]] = project_totals.get(tk["project"], 0.0) + sum(
            float(e["hours"]) for e in tk["entries"]
        )

    esc = lambda value: (str(value).replace("&", "&amp;").replace("<", "&lt;")
                         .replace(">", "&gt;").replace('"', "&quot;"))
    rows = []
    for tk in REPORT["tasks"]:
        for entry in tk["entries"]:
            rows.append(
                "<tr>"
                f"<td>{esc(tk['project'])}</td>"
                f"<td>{esc(tk['title'])}</td>"
                f"<td>{esc(entry['discipline'])}</td>"
                f"<td>{esc(entry['desc'])}</td>"
                f"<td style='text-align:right;white-space:nowrap'>{esc(entry['hours'])} hrs</td>"
                f"<td>{esc(entry['status'])}</td>"
                "</tr>"
            )
    if not rows:
        rows.append("<tr><td colspan='6'>No task entries were recorded.</td></tr>")

    project_rows = "".join(
        f"<tr><td>{esc(project)}</td><td style='text-align:right'>{hours:g} hrs</td></tr>"
        for project, hours in project_totals.items()
    ) or "<tr><td colspan='2'>No project entries recorded.</td></tr>"

    return f"""<!doctype html>
<html><body style="font-family:Arial,sans-serif;color:#202938;font-size:13px">
<h2 style="margin-bottom:4px">Daily Timesheet Report</h2>
<p style="margin-top:0;color:#667085">{esc(TODAY)} · {esc(TODAY_D.strftime('%A'))}</p>
<table cellpadding="8" cellspacing="0" style="border-collapse:collapse;width:100%;max-width:900px;margin-bottom:18px">
<tr style="background:#eef2f7"><th align="left">Project</th><th align="right">Hours</th></tr>
{project_rows}
<tr style="background:#f4f6f8;font-weight:bold"><td>Total logged</td><td align="right">{total:g} hrs</td></tr>
</table>
<h3>Task-wise details</h3>
<table cellpadding="7" cellspacing="0" border="1" style="border-collapse:collapse;width:100%;border-color:#d0d5dd">
<thead style="background:#eef2f7"><tr>
<th align="left">Project</th><th align="left">Task</th><th align="left">Discipline</th>
<th align="left">Description</th><th align="right">Hours</th><th align="left">Status</th>
</tr></thead><tbody>{''.join(rows)}</tbody></table>
<h3>Tracking</h3>
<p><b>Status:</b> {esc(tracking.get('status','UNKNOWN'))}</p>
<p><b>Today's tracking row:</b> {esc(tracking.get('row','Not available'))}</p>
<p style="color:#667085">One tracking screenshot is attached for all projects.</p>
</body></html>"""


# ---------------------------------------------------------------- email
def send_email(subject, body, attachments):
    if not MAIL_USER or not MAIL_PASS or not MAIL_TO:
        log("Email NOT sent: MAIL_USER / MAIL_PASS / MAIL_TO missing")
        return False
    try:
        msg = EmailMessage()
        msg["From"] = MAIL_USER
        msg["To"] = MAIL_TO
        msg["Subject"] = subject
        msg.set_content(body)
        msg.add_alternative(build_report_html(), subtype="html")
        for path, name in attachments:
            if os.path.exists(path):
                with open(path, "rb") as f:
                    msg.add_attachment(f.read(), maintype="image",
                                       subtype="png", filename=name)
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=30) as s:
            s.login(MAIL_USER, MAIL_PASS)
            s.send_message(msg)
        log(f"Email sent to {MAIL_TO}")
        return True
    except Exception as ex:
        log(f"Email FAILED: {ex}")
        return False


def build_report_email(expected, label):
    t = REPORT["tracking"] or {"status": "NOT CHECKED", "row": ""}
    ok = t["status"] == "OK" and not REPORT["error"]
    attention_prefix = "" if ok else "[ATTENTION] "
    subject = f"{attention_prefix}Timesheet - {TODAY_D.strftime('%d-%m-%Y')} - {LOGGED_IN_NAME}"
    lines = [
        f"Timesheet report for {TODAY} ({TODAY_D.strftime('%A')}, {label})",
        f"Expected hours: {expected}",
        f"Logged-in user: {LOGGED_IN_NAME}",
        f"Tracking page status: {t['status']}",
    ]
    if t.get("row"):
        lines.append(f"Tracking row text: {t['row']}")
    if t.get("note"):
        lines.append(f"Note: {t['note']}")
    if REPORT["error"]:
        lines.append(f"ERROR: {REPORT['error']}")
    lines.append("")
    lines.append("Tasks:")
    if not REPORT["tasks"]:
        lines.append("  (no task was entered)")
    for tk in REPORT["tasks"]:
        lines.append(f"- {tk['project']} | {tk['title']} | {tk['result']}")
        for e in tk["entries"]:
            lines.append(f"    {e['discipline']} | {e['desc']} | {e['hours']}h | {e['status']}")
    project_totals = OrderedDict()
    for tk in REPORT["tasks"]:
        project_totals[tk["project"]] = project_totals.get(tk["project"], 0.0) + sum(
            float(e["hours"]) for e in tk["entries"]
        )
    total = sum(project_totals.values())
    lines.append("")
    lines.append("Project-wise hours:")
    for project, hours in project_totals.items():
        lines.append(f"  {project}: {hours:g} hrs")
    lines.append(f"Total hours in this run: {total:g} hrs")
    lines.append("Screenshot attached: one current-date tracking page for all projects.")
    return subject, "\n".join(lines)


# ---------------------------------------------------------------- steps
def login(page):
    page.goto(BASE + "employee", wait_until="domcontentloaded")
    page.wait_for_selector("#email", state="visible")
    page.fill("#email", USER)
    page.fill("#password", PASS)
    page.click("input[type='button']")
    try:
        page.wait_for_function(
            "() => !document.title.toLowerCase().includes('employee sign in')",
            timeout=20000)
    except Exception:
        raise Exception("Login failed - check TS_USER / TS_PASS in .env")
    global LOGGED_IN_NAME
    if not TS_DISPLAY_NAME:
        try:
            detected = page.evaluate("""() => {
                const selectors = [
                    '.navbar .dropdown-toggle', '.navbar .user-name',
                    '.user-menu .dropdown-toggle', '.topbar .user-name',
                    '[data-user-name]', '.profile-name'
                ];
                for (const selector of selectors) {
                    const el = document.querySelector(selector);
                    if (el && el.innerText.trim()) return el.innerText.trim();
                }
                return '';
            }""")
            if detected:
                LOGGED_IN_NAME = " ".join(detected.split())[:80]
        except Exception:
            pass
    log(f"Logged in as: {LOGGED_IN_NAME}")


def board_link_ready(page):
    link = page.locator("a[href='board']")
    return link.count() > 0 and link.first.is_visible()



def open_board(page, project_code, project_name=""):
    page.goto(BASE + "project-details", wait_until="domcontentloaded")
    page.wait_for_timeout(1500)

    log(f"Selecting project by permanent code: {project_code}")

    try:
        # Find the visible project code text
        code = page.get_by_text(
            re.compile(rf"^{re.escape(project_code)}\s*-", re.I)
        ).first

        code.wait_for(state="visible", timeout=15000)

        # Find the project card containing the code
        card = code
        for _ in range(8):
            if card.locator("a").count() > 0:
                break
            card = card.locator("xpath=..")

        # Find the board link inside this project card
        board_link = card.locator("a[href*='board']").first

        if board_link.count() == 0:
            # If the code is inside the card but the link is outside,
            # click the project title within the nearest card.
            title = page.get_by_text(
                project_name, exact=False
            ).first
            title.click(timeout=10000)
        else:
            board_link.click(timeout=10000)

        page.wait_for_timeout(1500)

        if page.locator("a[href*='go-to-tabs']").count() == 0:
            dump_debug(page)
            raise Exception("Board link clicked, but task board did not load.")

        log(f"Board ready: {project_name} [{project_code}]")

    except Exception as ex:
        dump_debug(page)
        raise Exception(
            f"Could not open project {project_code}: {ex}"
        )

def in_range(card):
    if not card["start"] or not card["end"]:
        return False
    s = datetime.datetime.strptime(card["start"], "%d/%m/%Y").date()
    e = datetime.datetime.strptime(card["end"], "%d/%m/%Y").date()
    return s <= TODAY_D <= e


def todays_hours(page):
    rows = safe_eval(page, ROWS_JS)
    total = 0.0
    for r in rows:
        if r.strip().startswith(TODAY):
            for tok in r.split():
                if len(tok) == 5 and tok[2] == ":" and tok.replace(":", "").isdigit():
                    h, m = tok.split(":")
                    total += int(h) + int(m) / 60
                    break
    return total


def add_row(page):
    before = page.locator("#discipline").count()
    page.click("#Addrows")
    page.wait_for_function(
        "n => document.querySelectorAll('#discipline').length > n", arg=before,
        timeout=8000)


def fill_current_task(page, entries):
    """Returns: saved | done | no_add | dry"""
    try:
        page.wait_for_selector("#SaveTimelog", timeout=15000)
    except Exception:
        if "access denied" in body_text(page):
            log("Access denied on task page - skipped")
        else:
            log("Not a time log page - skipped")
        return "no_add"

    add = page.locator("#Addrows")
    if add.count() == 0 or not add.first.is_visible():
        log("No 'Add New Item' button on this task - trying next task")
        return "no_add"

    # Let the history table load so we never save a duplicate entry
    try:
        page.wait_for_load_state("networkidle", timeout=2500)
    except Exception:
        pass

    need = sum(float(e["hours"]) for e in entries)
    done = todays_hours(page)
    if done >= need:
        log(f"Already {done}h logged for {TODAY} on this task - nothing to do")
        return "done"

    date_box = page.locator("#currentdate")
    if date_box.count() > 0:
        try:
            cur = date_box.first.input_value(timeout=3000)
            if cur != TODAY:
                log(f"Form date is {cur}, setting {TODAY}")
                date_box.first.fill(TODAY, timeout=3000)
        except Exception:
            pass

    try:
        for i, e in enumerate(entries):
            if i > 0 or page.locator("#discipline").count() == 0:
                add_row(page)
            page.evaluate(PICK_JS, ["#discipline", e["discipline"]])
            page.locator("#description").last.fill(e["desc"])
            page.locator("#scduledHours").last.fill(e["hours"])
            page.locator("#actualHours").last.fill(e["hours"])
            page.evaluate(PICK_JS, ["#status", e["status"]])
            log(f"Filled: {e['discipline']} | {e['hours']}h | {e['status']}")
    except Exception:
        dump_form(page)
        raise

    if MODE == "dry":
        log("DRY mode - Save NOT clicked")
        return "dry"

    page.click("#SaveTimelog")
    after = 0.0
    for _ in range(24):
        page.wait_for_timeout(250)
        try:
            after = todays_hours(page)
        except Exception:
            continue
        if after >= need:
            break
    if after >= need:
        log(f"Saved OK - {after}h now logged for {TODAY}")
    else:
        message = f"Save could not be verified: expected {need:g}h, but only {after:g}h appears for {TODAY}"
        log("WARNING: " + message)
        REPORT["error"] = (REPORT["error"] + "; " if REPORT["error"] else "") + message
        return "save_failed"
    return "saved"


def process_project(page, project_code, project_name, ents):
    open_board(page, project_code, project_name)
    cards = safe_eval(page, CARDS_JS)
    log(f"{len(cards)} task card(s) found")
    for c in cards:
        log(f"  {c['title']} | {c['start']} - {c['end']}")

    todays = [c for c in cards if in_range(c)]
    if not todays:
        message = f"No task card covers {TODAY} for {project_name} [{project_code}]"
        log(message)
        REPORT["error"] = (REPORT["error"] + "; " if REPORT["error"] else "") + message
        REPORT["tasks"].append(dict(project=f"{project_name} [{project_code}]",
                                    title="No matching task card",
                                    result="NOT ENTERED", entries=ents))
        return

    for idx, c in enumerate(todays):
        if idx > 0:
            open_board(page, project_code, project_name)
        log(f"Opening task: {c['title']}")
        page.evaluate(CLICK_HREF_JS, c["href"])
        result = fill_current_task(page, ents)
        if result == "dry":
            input("Check the browser, then press Enter to continue... ")
            return
        if result in ("saved", "done"):
            REPORT["tasks"].append(dict(project=f"{project_name} [{project_code}]", title=c["title"],
                                        result=result, entries=ents))
            return
        if result == "save_failed":
            REPORT["tasks"].append(dict(project=f"{project_name} [{project_code}]", title=c["title"],
                                        result="SAVE NOT VERIFIED", entries=ents))
            return
    message = f"No matching task for {project_name} [{project_code}] had an Add New Item button"
    log(message)
    REPORT["error"] = (REPORT["error"] + "; " if REPORT["error"] else "") + message
    REPORT["tasks"].append(dict(project=f"{project_name} [{project_code}]",
                                title="Task form unavailable",
                                result="NOT ENTERED", entries=ents))


# ---------------------------------------------------------------- tracking page
def date_variants(d):
    return [
        d.strftime("%d/%m/%Y"), f"{d.day}/{d.month}/{d.year}",
        d.strftime("%d-%m-%Y"), d.strftime("%Y-%m-%d"),
        d.strftime("%d-%b-%Y"), d.strftime("%d %b %Y"), d.strftime("%d %B %Y"),
        d.strftime("%b %d, %Y"), d.strftime("%B %d, %Y"),
    ]


def check_tracking(page, expected):
    """Open timesheet-tracking, find today's row, report OK / LOP / CHECK."""
    info = {"status": "UNKNOWN", "row": "", "note": ""}
    try:
        link = page.locator("a[href*='timesheet-tracking']")
        if link.count() > 0:
            link.first.click()
        else:
            page.goto(BASE + "timesheet-tracking", wait_until="domcontentloaded")
        settle(page)
        try:
            page.wait_for_selector("tr", timeout=8000)
        except Exception:
            pass
        try:
            page.wait_for_load_state("networkidle", timeout=3000)
        except Exception:
            pass

        found_today = save_tracking_screenshots(page)
        if not found_today:
            info["note"] = "Today's date column was not found; tracking screenshots were still saved."

        if "access denied" in body_text(page):
            info["status"] = "UNKNOWN"
            info["note"] = "Access denied on timesheet-tracking page"
            return info

        rows = safe_eval(page, ROWS_JS)
        variants = [v.lower() for v in date_variants(TODAY_D)]
        match = None
        for r in rows:
            flat = " ".join(r.split())
            if len(flat) < 600 and any(v in flat.lower() for v in variants):
                match = flat
                break
        if not match:
            info["status"] = "CHECK"
            info["note"] = "Today tracking date could not be verified; please review the attached tracking screenshot."
            return info

        info["row"] = match
        if re.search(r"\bLOP\b", match, re.I):
            info["status"] = "LOP"
            return info

        hours = []
        for h, m in re.findall(r"(\d{1,2}):(\d{2})", match):
            hours.append(int(h) + int(m) / 60)
        for v in re.findall(r"(\d+(?:\.\d+)?)\s*(?:hrs?|hours)\b", match, re.I):
            hours.append(float(v))
        if not hours:
            info["status"] = "CHECK"
            info["note"] = "Today's row found but no hours could be read from it"
        elif max(hours) >= expected:
            info["status"] = "OK"
        else:
            info["status"] = "CHECK"
            info["note"] = f"Row shows less than the expected {expected} hours"
        return info
    except Exception as ex:
        info["note"] = f"Tracking check error: {ex}"
        return info


def track_only(page):
    """Discovery: open the tracking page and save what it contains."""
    link = page.locator("a[href*='timesheet-tracking']")
    if link.count() > 0:
        log("Found a tracking link in the menu, clicking it")
        link.first.click()
    else:
        log("No menu link, opening the URL directly")
        page.goto(BASE + "timesheet-tracking", wait_until="domcontentloaded")
    settle(page)
    try:
        page.wait_for_load_state("networkidle", timeout=4000)
    except Exception:
        pass
    page.wait_for_timeout(1500)
    rows = safe_eval(page, ROWS_JS)
    text = "URL: " + page.url + "\n\n---- ROWS ----\n"
    text += "\n".join(" ".join(r.split())[:300] for r in rows)
    text += "\n\n---- PAGE TEXT ----\n" + safe_eval(page, "() => document.body.innerText")[:5000]
    with open("tracking_page.txt", "w", encoding="utf-8") as f:
        f.write(text)
    take_clear_screenshot(page, "tracking_page.png", "tr")
    print(text[:3500])
    log("Saved tracking_page.txt and tracking_page.png")


# ---------------------------------------------------------------- main
def main():
    if MODE == "mailtest":
        if not MAIL_USER or not MAIL_PASS or not MAIL_TO:
            raise Exception("Set MAIL_USER, MAIL_PASS and MAIL_TO in GitHub Secrets / .env")
        if not send_email("Timesheet bot test email",
                          "If you can read this, email sending works.", []):
            raise Exception("Test email could not be sent. Check SMTP credentials and recipient.")
        return

    if MODE == "track":
        expected, label, groups = WEEKDAY_HOURS, "Tracking", {}
    else:
        plan = day_plan()
        if plan is None:
            log(f"{TODAY_D.strftime('%A')} {TODAY} is not a working day - skipped")
            return
        path, expected, label = plan
        source = TASK_SOURCE_URL or path

        try:
            if not TASK_SOURCE_URL and not os.path.exists(path):
                raise FileNotFoundError(
                    "Task spreadsheet source is missing. Configure TASK_SOURCE_URL "
                    f"or provide {path}."
                )
            entries = read_entries(source)

            if not entries:
                REPORT["error"] = (
                    f"No task rows found in the spreadsheet for current date "
                    f"{TODAY_D}. Please add today's tasks to the '{TASK_SHEET_NAME}' sheet."
                )
                log(REPORT["error"])
                subject = f"[ATTENTION] Timesheet task reminder - {TODAY_D:%d-%m-%Y} - No tasks in Excel"
                body = (
                    f"TIMESHEET REMINDER - ACTION REQUIRED\\n\\n"
                    f"No task rows were found for today: {TODAY_D:%d-%m-%Y} ({TODAY_D:%A}).\\n"
                    f"Expected hours: {expected} ({label}).\\n"
                    f"Source sheet: {TASK_SHEET_NAME}.\\n\\n"
                    "Please add today's Project Code, Project Name, Discipline, "
                    "Activity Description, Actual Hours and Status to the task sheet.\\n"
                    "No timesheet entries were submitted."
                )
                if not send_email(subject, body, []):
                    raise Exception(REPORT["error"] + " Reminder email failed; check MAIL_USER, MAIL_PASS and MAIL_TO.")
                log("No-task reminder email sent.")
                return

            total = sum(float(e["hours"]) for e in entries)
            if abs(total - expected) > 0.01:
                warning = f"Spreadsheet hours total {total:g}h; expected {expected}h for {label}."
                log("WARNING: " + warning)
                REPORT["error"] = warning

            groups = OrderedDict()
            for e in entries:
                groups.setdefault(e["project_code"], {
                    "name": e["project_name"], "entries": []
                })["entries"].append(e)

        except Exception as ex:
            REPORT["error"] = f"Could not read today's task spreadsheet: {ex}"
            log(REPORT["error"])
            subject = f"[ATTENTION] Timesheet spreadsheet error - {TODAY_D:%d-%m-%Y}"
            body = (
                f"TIMESHEET ALERT - SPREADSHEET ISSUE\\n\\n"
                f"Date: {TODAY_D:%d-%m-%Y}\\n"
                f"Sheet: {TASK_SHEET_NAME}\\n"
                f"Error: {REPORT['error']}\\n\\n"
                "No timesheet entries were submitted. Check the spreadsheet URL, "
                "sharing permission, sheet tab and today's rows."
            )
            if not send_email(subject, body, []):
                raise Exception(REPORT["error"] + "; alert email also failed.")
            return

    if MODE != "track" and (not USER or not PASS):
        REPORT["error"] = "TS_USER or TS_PASS is missing."
        if not send_email(f"[ATTENTION] Timesheet configuration error - {TODAY_D:%d-%m-%Y}",
                          REPORT["error"], []):
            raise Exception(REPORT["error"] + "; alert email also failed.")
        raise Exception(REPORT["error"])

    browser = context = None
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(
                headless=(MODE == "hide"),
                slow_mo=SLOW_MO,
                args=["--start-maximized"]
            )
            context = browser.new_context(
                no_viewport=True,
                is_mobile=False,
                has_touch=False
            )
            page = context.new_page()

            if MODE != "hide":
                try:
                    cdp = context.new_cdp_session(page)
                    window_id = cdp.send("Browser.getWindowForTarget")["windowId"]
                    cdp.send("Browser.setWindowBounds",
                             {"windowId": window_id, "bounds": {"windowState": "fullscreen"}})
                    log("Browser fullscreen enabled")
                except Exception as ex:
                    log(f"Browser fullscreen warning: {ex}")

            page.set_default_timeout(15000)
            if BLOCK_HEAVY:
                page.route("**/*", block_heavy)

            login(page)
            if MODE == "track":
                track_only(page)
                return

            for project_code, project_data in groups.items():
                process_project(page, project_code, project_data["name"], project_data["entries"])

            if MODE != "dry":
                REPORT["tracking"] = check_tracking(page, expected)
                t = REPORT["tracking"]
                log(f"Tracking page: {t['status']} | {t['row']} {t['note']}")
                if t["status"] != "OK":
                    problem = f"Tracking verification status: {t['status']}. {t.get('note', '')}"
                    REPORT["error"] = (REPORT["error"] + "; " if REPORT["error"] else "") + problem

                subject, body = build_report_email(expected, label)
                if not send_email(subject, body, REPORT["shots"]):
                    raise Exception("Timesheet processing finished, but report email could not be sent.")
            log(f"FINISHED in {time.time() - T0:.1f} seconds")
            if MODE not in ("hide", "dry"):
                page.wait_for_timeout(END_PAUSE_MS)

    except Exception as ex:
        logging.error(f"FAILED: {ex}")
        print("FAILED:", ex)
        REPORT["error"] = (REPORT["error"] + "; " if REPORT["error"] else "") + str(ex)
        if MODE not in ("dry", "track"):
            try:
                subject, body = build_report_email(expected, label)
                send_email("[ATTENTION] Timesheet FAILED " + TODAY, body, REPORT["shots"])
            except Exception as mail_ex:
                log(f"Failure email could not be prepared/sent: {mail_ex}")
        raise
    finally:
        if context is not None:
            context.close()
        if browser is not None:
            browser.close()

main()
