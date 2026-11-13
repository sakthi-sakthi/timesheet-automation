import os, re, sys, time, smtplib, datetime, logging
from collections import OrderedDict
from pathlib import Path
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

REPORT = {"tasks": [], "shots": [], "tracking": None, "error": None, "attempted": [], "verified_hours": 0.0}


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


# Maximum actual hours accepted by the portal for ONE time-log row.
# Larger source tasks are split into balanced rows; total hours are preserved.
MAX_HOURS_PER_ENTRY = 5.0


def hours_to_portal_time(hours):
    """Convert decimal hours from the sheet to the portal's HH:MM format.

    Examples: 0.5 -> 00:30, 1 -> 01:00, 1.5 -> 01:30, 5 -> 05:00.
    The portal time inputs must never receive decimal strings such as '0.5'.
    """
    from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

    try:
        value = Decimal(str(hours))
    except (InvalidOperation, ValueError):
        raise ValueError(f"Invalid hours value for portal time field: {hours}")

    if value <= 0:
        raise ValueError(f"Hours must be greater than zero: {hours}")

    total_minutes = int((value * Decimal("60")).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
    whole_hours, minutes = divmod(total_minutes, 60)
    return f"{whole_hours:02d}:{minutes:02d}"


def split_long_task_entries(entries):
    """
    Split any task above the per-row portal limit into balanced entries.

    Examples:
      6h  -> 3h + 3h
      7h  -> 3.5h + 3.5h
      11h -> 3.67h + 3.67h + 3.66h

    The first row keeps the original description. Additional rows retain the
    same project and discipline and get a clear continuation suffix.
    """
    split_entries = []

    for entry in entries:
        total_hours = float(entry["hours"])

        if total_hours <= MAX_HOURS_PER_ENTRY:
            split_entries.append(dict(entry))
            continue

        parts = int((total_hours + MAX_HOURS_PER_ENTRY - 1e-9) // MAX_HOURS_PER_ENTRY)
        parts = max(parts, 2)

        # Balance the hours across the minimum number of rows needed.
        # Round all but the final part to 2 decimals; put the remainder in
        # the final part so the original total is preserved exactly to cents.
        base = round(total_hours / parts, 2)
        remaining = round(total_hours, 2)

        for index in range(parts):
            if index == parts - 1:
                part_hours = round(remaining, 2)
            else:
                part_hours = base
                remaining = round(remaining - part_hours, 2)

            if part_hours <= 0 or part_hours > MAX_HOURS_PER_ENTRY:
                raise ValueError(
                    f"Could not safely split {entry['desc']} ({total_hours:g}h) "
                    f"into rows of at most {MAX_HOURS_PER_ENTRY:g}h."
                )

            part = dict(entry)
            part["hours"] = f"{part_hours:g}"

            if index > 0:
                part["desc"] = (
                    f"{entry['desc']} – Additional work (Part {index + 1})"
                )

            split_entries.append(part)

        log(
            f"Split task: {entry['desc']} | {total_hours:g}h into "
            + " + ".join(
                f"{split_entries[-parts + i]['hours']}h"
                for i in range(parts)
            )
        )

    original_total = round(sum(float(e["hours"]) for e in entries), 2)
    split_total = round(sum(float(e["hours"]) for e in split_entries), 2)

    if original_total != split_total:
        raise ValueError(
            f"Task splitting changed total hours: "
            f"original={original_total:g}h, split={split_total:g}h"
        )

    if any(float(e["hours"]) > MAX_HOURS_PER_ENTRY for e in split_entries):
        raise ValueError("A split entry still exceeds the 5-hour portal row limit.")

    return split_entries



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
    if not entries:
        raise ValueError(
            "No valid rows found. Fill Project, Discipline, Activity Description, "
            "Actual Hours and Status in the Daily Tasks sheet."
        )
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


def current_week_dates():
    """Monday through Sunday for the current local date."""
    monday = TODAY_D - datetime.timedelta(days=TODAY_D.weekday())
    return [monday + datetime.timedelta(days=i) for i in range(7)]


def save_tracking_screenshots(page):
    """Capture only the current week's date columns, with vertical scroll tiles."""
    prepare_desktop_capture(page)
    week_dates = current_week_dates()
    week_labels = []
    for d in week_dates:
        week_labels.extend([
            d.strftime("%d %b").upper(),
            f"{d.day} {d.strftime('%b').upper()}",
            d.strftime("%d %B").upper(),
            f"{d.day} {d.strftime('%B').upper()}"
        ])
    week_labels = list(dict.fromkeys(week_labels))

    # Find the main scrollable table container and the first/last current-week
    # date headers. Keep a little space before Monday for employee columns.
    scroll_info = page.evaluate("""(labels) => {
        const norm = s => String(s || '').replace(/\\s+/g, ' ').trim().toUpperCase();
        const candidates = Array.from(document.querySelectorAll('body *')).filter(el => {
          const r = el.getBoundingClientRect(), cs = getComputedStyle(el);
          return r.width > 80 && r.height > 80 &&
            (el.scrollHeight > el.clientHeight + 30 || el.scrollWidth > el.clientWidth + 30) &&
            /(auto|scroll)/.test(cs.overflowY + ' ' + cs.overflowX);
        });
        candidates.sort((a,b) => {
          const score = el => (el.scrollHeight-el.clientHeight) +
            (el.scrollWidth-el.clientWidth) + Math.min(el.clientWidth*el.clientHeight/1000,1000);
          return score(b)-score(a);
        });
        const scroller = candidates[0];
        if (!scroller) return null;
        if (!scroller.dataset.tsCaptureId) scroller.dataset.tsCaptureId = 'ts-week-capture';
        const headers = Array.from(document.querySelectorAll('th')).map(el => ({
          text: norm(el.innerText),
          x: el.getBoundingClientRect().left - scroller.getBoundingClientRect().left + scroller.scrollLeft,
          right: el.getBoundingClientRect().right - scroller.getBoundingClientRect().left + scroller.scrollLeft
        }));
        const matched = headers.filter(h => labels.some(label => h.text === norm(label) || h.text.startsWith(norm(label) + ' ') || h.text.includes(norm(label))));
        return {
          selector: '[data-ts-capture-id="ts-week-capture"]',
          sh: scroller.scrollHeight, ch: scroller.clientHeight,
          sw: scroller.scrollWidth, cw: scroller.clientWidth,
          firstX: matched.length ? Math.min(...matched.map(h => h.x)) : null,
          lastRight: matched.length ? Math.max(...matched.map(h => h.right)) : null,
          matched: matched.map(h => h.text)
        };
    }""", week_labels)

    shots = []
    try:
        if not scroll_info:
            page.screenshot(path="tracking_week_01.png", full_page=True,
                            animations="disabled", type="png")
            shots.append(("tracking_week_01.png", "tracking_week_01.png"))
        else:
            target = page.locator(scroll_info["selector"]).first
            dims = target.evaluate("el => ({sh:el.scrollHeight,ch:el.clientHeight,sw:el.scrollWidth,cw:el.clientWidth})")
            max_y = max(0, dims["sh"] - dims["ch"])
            step_y = max(250, int(dims["ch"] * 0.78))
            y_positions = list(range(0, max_y + 1, step_y)) or [0]
            if y_positions[-1] != max_y:
                y_positions.append(max_y)

            # Focus horizontal capture on Monday-Sunday only. If the week spans
            # more than one viewport, take the minimum number of tiles needed.
            if scroll_info["firstX"] is not None:
                start_x = max(0, int(scroll_info["firstX"] - 300))
                end_x = max(start_x, int(scroll_info["lastRight"] - dims["cw"] + 20))
                step_x = max(300, int(dims["cw"] * 0.78))
                x_positions = list(range(start_x, end_x + 1, step_x)) or [start_x]
                if x_positions[-1] != end_x:
                    x_positions.append(end_x)
                log("Current week headers found: " + ", ".join(scroll_info["matched"]))
            else:
                # Header labels were not recognized; use one centered view rather
                # than emailing every historical date column.
                x_positions = [max(0, min(dims["sw"] - dims["cw"], int((dims["sw"] - dims["cw"]) / 2)))]
                log("Current-week headers not recognized; captured one centered tracking view")

            original = target.evaluate("el => ({x:el.scrollLeft,y:el.scrollTop})")
            index = 0
            for y in y_positions:
                for x in x_positions:
                    target.evaluate("(el,p) => {el.scrollTop=p.y; el.scrollLeft=p.x;}", {"x": x, "y": y})
                    page.wait_for_timeout(250)
                    index += 1
                    name = f"tracking_week_{index:02d}.png"
                    page.screenshot(path=name, full_page=False,
                                    animations="disabled", type="png")
                    shots.append((name, name))
                    log(f"Current-week screenshot {index}: scrollX={x}, scrollY={y}")
            target.evaluate("(el,p) => {el.scrollTop=p.y; el.scrollLeft=p.x;}", original)
    except Exception as ex:
        log(f"Current-week screenshot warning: {ex}")
        if not shots:
            page.screenshot(path="tracking_week_01.png", full_page=False,
                            animations="disabled", type="png")
            shots.append(("tracking_week_01.png", "tracking_week_01.png"))

    REPORT["shots"] = shots
    log(f"Saved {len(shots)} current-week tracking screenshot(s)")
    return scroll_tracking_to_today(page)


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
<tr style="background:#f4f6f8;font-weight:bold"><td>Confirmed task-entry hours</td><td align="right">{total:g} hrs</td></tr>
<tr><td>Newly verified saved hours</td><td align="right">{REPORT.get("verified_hours", 0.0):g} hrs</td></tr>
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
<p style="color:#667085">Only the current Monday-Sunday tracking dates are attached, across vertical scroll sections.</p>
</body></html>"""


# ---------------------------------------------------------------- email
def send_email(subject, body, attachments):
    if not MAIL_USER or not MAIL_PASS:
        log("Email NOT sent: MAIL_USER / MAIL_PASS missing in .env")
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
    lines.append("Current-week tracking screenshots attached (Monday-Sunday), split only where vertical scrolling is required.")
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
    """Read the total hours shown in today's time-log history on this task page."""
    rows = safe_eval(page, ROWS_JS)
    total = 0.0
    variants = {TODAY, TODAY_D.strftime("%-d/%-m/%Y") if os.name != "nt" else f"{TODAY_D.day}/{TODAY_D.month}/{TODAY_D.year}",
                TODAY_D.strftime("%d-%m-%Y"), TODAY_D.strftime("%Y-%m-%d")}
    for raw in rows:
        row = " ".join(str(raw).split())
        # Match date anywhere in the row; some portal tables prepend a serial number.
        if not any(v.lower() in row.lower() for v in variants):
            continue
        # The history table can show scheduled and actual time in the same row.
        # Count only the first time value per dated history row to avoid doubling it.
        match = re.search(r"(?<!\d)(\d{1,2}):(\d{2})(?!\d)", row)
        if match:
            h, m = match.groups()
            minutes = int(m)
            if minutes < 60:
                total += int(h) + minutes / 60
    return round(total, 2)


def _timesheet_field_counts(page):
    """Return counts for the repeated controls in the time-log form."""
    return page.evaluate("""() => ({
        discipline: document.querySelectorAll('#discipline').length,
        description: document.querySelectorAll('#description').length,
        scheduled: document.querySelectorAll('#scduledHours').length,
        actual: document.querySelectorAll('#actualHours').length,
        status: document.querySelectorAll('#status').length
    })""")


def add_row(page):
    """Add one row and wait until every required control gains one instance.

    These selectors are not guaranteed to have equal baseline counts: for example,
    #description may also match static/search fields outside the repeated row.
    Therefore compare each selector with its own pre-click count rather than
    requiring all selector counts to be identical.
    """
    button = page.locator("#Addrows").first
    if button.count() == 0 or not button.is_visible():
        raise RuntimeError("Add New Item button (#Addrows) was not found or is not visible.")

    before = _timesheet_field_counts(page)
    expected = {key: value + 1 for key, value in before.items()}

    def wait_for_complete_row(timeout_ms):
        page.wait_for_function(
            """expected => {
                const now = {
                    discipline: document.querySelectorAll('#discipline').length,
                    description: document.querySelectorAll('#description').length,
                    scheduled: document.querySelectorAll('#scduledHours').length,
                    actual: document.querySelectorAll('#actualHours').length,
                    status: document.querySelectorAll('#status').length
                };
                return Object.keys(expected).every(key => now[key] >= expected[key]);
            }""",
            arg=expected, timeout=timeout_ms
        )

    button.scroll_into_view_if_needed(timeout=5000)
    button.click(timeout=10000)
    try:
        wait_for_complete_row(15000)
    except Exception:
        current = _timesheet_field_counts(page)
        # Retry only when no field count changed at all; partial DOM cloning is unsafe.
        if any(current[k] != before[k] for k in before):
            dump_form(page)
            raise RuntimeError(f"Incomplete new row controls. Before={before}; after={current}.")
        page.locator("#Addrows").first.evaluate("el => el.click()")
        try:
            wait_for_complete_row(15000)
        except Exception as ex:
            after = _timesheet_field_counts(page)
            dump_form(page)
            raise RuntimeError(f"Could not add a complete time-log row. Before={before}; after={after}.") from ex

    after = _timesheet_field_counts(page)
    missing = [key for key in before if after[key] < before[key] + 1]
    if missing:
        dump_form(page)
        raise RuntimeError(
            f"New row is missing required controls {missing}. "
            f"Before={before}; after={after}."
        )
    log(f"Complete new time-log row added. Field counts: {after}")


def fill_current_task(page, entries):
    """Fill every entry for one project task, save, and wait for verified persistence.

    Returns: saved | done | no_add | partial | dry. A failed save raises and
    stops the run so the next project is never opened prematurely.
    """
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

    try:
        page.wait_for_load_state("networkidle", timeout=5000)
    except Exception:
        pass

    need = round(sum(float(e["hours"]) for e in entries), 2)
    baseline = todays_hours(page)
    if baseline >= need:
        log(f"Already {baseline:g}h logged for {TODAY} on this task; expected {need:g}h. No duplicate entry added.")
        return "done"
    if baseline > 0:
        raise RuntimeError(
            f"Partial existing entries detected on task: {baseline:g}h present, {need:g}h expected. "
            "Stopped to prevent duplicate hours. Reconcile the portal entries manually."
        )

    date_box = page.locator("#currentdate")
    if date_box.count() > 0:
        cur = date_box.first.input_value(timeout=3000)
        if cur != TODAY:
            date_box.first.fill(TODAY, timeout=5000)

    try:
        for i, e in enumerate(entries):
            if i > 0 or page.locator("#discipline").count() == 0:
                add_row(page)

            counts = _timesheet_field_counts(page)
            if len(set(counts.values())) != 1 or counts["discipline"] <= i:
                raise RuntimeError(f"Form row controls are not aligned before filling entry {i + 1}: {counts}")

            page.evaluate(PICK_JS, ["#discipline", e["discipline"]])
            page.locator("#description").nth(i).fill(e["desc"])
            portal_time = hours_to_portal_time(e["hours"])
            page.locator("#scduledHours").nth(i).fill(portal_time)
            page.locator("#actualHours").nth(i).fill(portal_time)
            page.evaluate(PICK_JS, ["#status", e["status"]])
            log(f"Filled entry {i + 1}/{len(entries)}: {e['desc']} | {e['hours']}h => {portal_time}")

        # Final pre-save validation: all rows exist and each required field is populated.
        form_check = page.evaluate("""() => {
            const ids = ['discipline','description','scduledHours','actualHours','status'];
            const counts = Object.fromEntries(ids.map(id => [id, document.querySelectorAll('#'+id).length]));
            const n = counts.discipline;
            const values = ids.map(id => Array.from(document.querySelectorAll('#'+id)).slice(0,n).map(el =>
                (el.value || el.innerText || '').trim()));
            return {counts, complete: Object.values(counts).every(v => v === n) &&
                values.every(list => list.length === n && list.every(v => v !== ''))};
        }""")
        if not form_check.get("complete") or form_check["counts"]["discipline"] < len(entries):
            dump_form(page)
            raise RuntimeError(f"Pre-save form validation failed: {form_check}")
    except Exception:
        dump_form(page)
        raise

    if MODE == "dry":
        log("DRY mode - all project entries filled; Save NOT clicked")
        return "dry"

    log(f"Saving project task with {len(entries)} entries ({need:g}h total). Waiting for portal confirmation...")
    page.locator("#SaveTimelog").click(timeout=15000)

    # Wait up to 90 seconds for the server-side save/history table to reflect all hours.
    # Require two consecutive matching reads to avoid accepting a transient UI update.
    deadline = time.time() + 90
    consecutive = 0
    last_seen = baseline
    while time.time() < deadline:
        page.wait_for_timeout(1000)
        try:
            last_seen = todays_hours(page)
            if last_seen >= round(baseline + need, 2):
                consecutive += 1
                if consecutive >= 3:
                    log(f"PROJECT SAVE VERIFIED: {last_seen:g}h visible for {TODAY}; next project may start.")
                    REPORT["verified_hours"] += need
                    return "saved"
            else:
                consecutive = 0
        except Exception as ex:
            log(f"Waiting for saved history to load: {ex}")
        # A server-side validation message should fail promptly, not wait the full timeout.
        try:
            alert_text = page.locator(".alert-danger, .text-danger, .validation-summary-errors").all_inner_texts()
            errors = [x.strip() for x in alert_text if x.strip()]
            if errors:
                raise RuntimeError("Portal validation error after save: " + " | ".join(errors)[:1000])
        except RuntimeError:
            raise
        except Exception:
            pass

    dump_form(page)
    raise RuntimeError(
        f"SAVE NOT VERIFIED for this project task. Expected {need:g}h additional; "
        f"history shows {last_seen:g}h (baseline {baseline:g}h). "
        "Stopped before opening the next project to prevent misleading success or duplicate entries."
    )


def process_project(page, project_code, project_name, ents):
    """Process one project as a unit; do not return until its save is verified."""
    open_board(page, project_code, project_name)
    cards = safe_eval(page, CARDS_JS)
    log(f"Project {project_name} [{project_code}]: {len(cards)} task card(s) found")
    todays = [c for c in cards if in_range(c)]
    if not todays:
        raise RuntimeError(f"No task card covers {TODAY} for {project_name} [{project_code}]. Project not saved.")

    for idx, c in enumerate(todays):
        if idx > 0:
            open_board(page, project_code, project_name)
        log(f"Opening task for project {project_code}: {c['title']}")
        page.evaluate(CLICK_HREF_JS, c["href"])
        result = fill_current_task(page, ents)
        if result == "dry":
            REPORT["attempted"].append({"project": f"{project_name} [{project_code}]", "hours": sum(float(e["hours"]) for e in ents), "result": "DRY - not saved"})
            input("Check the browser, then press Enter to continue... ")
            return "dry"
        if result in ("saved", "done"):
            REPORT["tasks"].append(dict(project=f"{project_name} [{project_code}]", title=c["title"],
                                        result=("verified saved" if result == "saved" else "already present"), entries=ents))
            log(f"PROJECT COMPLETE: {project_name} [{project_code}] - {result}. Proceeding to next project.")
            return result
        if result == "no_add":
            continue
        raise RuntimeError(f"Project {project_code} stopped with status: {result}")

    raise RuntimeError(f"No eligible task page with Add New Item was found for {project_name} [{project_code}].")


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
            raise Exception("Set MAIL_USER, MAIL_PASS and MAIL_TO in .env before testing email")
        send_email("Timesheet bot test email",
                   "If you can read this, email sending works.", [])
        return

    if MODE not in ("track",) and (not USER or not PASS):
        raise Exception("Set TS_USER and TS_PASS in .env")
    plan = day_plan()
    if plan is None and MODE != "track":
        log(f"{TODAY_D.strftime('%A')} {TODAY} is not a working day - skipped")
        return

    if MODE != "track":
        path, expected, label = plan
        source = TASK_SOURCE_URL or path
        if not TASK_SOURCE_URL and not os.path.exists(path):
            raise Exception(
                "Set TASK_SOURCE_URL in .env to your public Excel/Google Sheets URL, "
                f"or provide the local {path} file."
            )
        entries = read_entries(source)

        # Keep every source task as a separate portal task entry.
        # Only split an individual task when it exceeds the portal's 5-hour
        # per-row limit. Decimal hours are converted to HH:MM at form fill time.
        source_total = sum(float(e["hours"]) for e in entries)
        entries = split_long_task_entries(entries)
        total = sum(float(e["hours"]) for e in entries)

        if round(source_total, 2) != round(total, 2):
            raise ValueError(
                f"Hours changed during task splitting: "
                f"source={source_total:g}h, prepared={total:g}h"
            )

        # Convert decimal hours to HH:MM before sending values to the portal.
        for entry in entries:
            log(
                f"Prepared task entry: {entry['project_code']} | "
                f"{entry['desc']} | {entry['hours']}h "
                f"=> {hours_to_portal_time(entry['hours'])}"
            )

        # Expected hours are a MINIMUM, never a maximum. Extra hours are valid.
        if total + 0.01 < expected:
            log(f"WARNING: {path} adds up to {total:g}h; {label} minimum is {expected}h")
        else:
            log(f"Hours validation passed: {total:g}h (minimum {expected}h; no upper cap)")
        groups = OrderedDict()
        for e in entries:
            groups.setdefault(e["project_code"], {"name": e["project_name"], "entries": []})["entries"].append(e)
    else:
        path, expected, label = "", WEEKDAY_HOURS, "track"
        groups = {}

    with sync_playwright() as p:
        # Open a real desktop browser window (not a fixed emulated viewport).
        browser = p.chromium.launch(
            headless=(MODE == "hide"),
            slow_mo=SLOW_MO,
            args=["--start-maximized"]
        )

        # Use the actual Windows display dimensions.
        context = browser.new_context(
            no_viewport=True,
            is_mobile=False,
            has_touch=False
        )
        page = context.new_page()

        # Force the visible Chromium window into fullscreen.
        if MODE != "hide":
            try:
                cdp = context.new_cdp_session(page)
                window_id = cdp.send("Browser.getWindowForTarget")["windowId"]
                cdp.send(
                    "Browser.setWindowBounds",
                    {
                        "windowId": window_id,
                        "bounds": {"windowState": "fullscreen"}
                    }
                )
                log("Browser fullscreen enabled")
            except Exception as ex:
                log(f"Browser fullscreen warning: {ex}")

        page.set_default_timeout(15000)
        if BLOCK_HEAVY:
            page.route("**/*", block_heavy)
        try:
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
                subject, body = build_report_email(expected, label)
                send_email(subject, body, REPORT["shots"])

            log(f"FINISHED in {time.time() - T0:.1f} seconds")
            if MODE not in ("hide", "dry"):
                page.wait_for_timeout(END_PAUSE_MS)
        except Exception as ex:
            logging.error(f"FAILED: {ex}")
            print("FAILED:", ex)
            REPORT["error"] = str(ex)
            # Do not attach task/error screenshots; only the single tracking
            # screenshot is included when the tracking page was reached.
            if MODE not in ("dry", "track"):
                subject, body = build_report_email(expected, label)
                send_email("[ATTENTION] Timesheet FAILED " + TODAY, body, REPORT["shots"])
            raise
        finally:
            # The sync_playwright context manager stops Playwright automatically.
            # Close resources once, in order, and do not let cleanup mask the real error.
            try:
                if context and not context.pages[0].is_closed():
                    context.close()
            except Exception as cleanup_error:
                log(f"Context cleanup warning (ignored): {cleanup_error}")
            try:
                if browser and browser.is_connected():
                    browser.close()
            except Exception as cleanup_error:
                log(f"Browser cleanup warning (ignored): {cleanup_error}")


main()
