"""Security_Settings.xlsx (next to the program): everything the security run needs, editable in Excel.

    Sites   Company | No | Site name | Client No | Address | Mon open | Mon close | ... | Sun close | Active | 24/7
            one row per site, grouped by company; a new site = one new row under its company.
            Times like  5:00 AM / 12:00 AM  (12:00 AM as a CLOSE time = midnight). Both empty for a day = closed that day.
    Run     the rules (10-minute allowance, 4-hour arming window, 3.5-hour disarm window), the fortnight start, the schedule.
"""
import datetime
import os
import re

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from . import paths

FILE_NAME = "Security_Settings.xlsx"
DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
SITE_HEADERS = (["Company", "No", "Site name", "Client No", "Address"]
                + [f"{d} {w}" for d in DAYS for w in ("open", "close")] + ["Active", "24/7"])
RUN_DEFAULTS = [
    ("Late opening allowance (minutes)", "10", "OPEN is highlighted when the first disarm is MORE than this many minutes after the opening time."),
    ("Late arming window (hours)", "4", "An arm up to this many hours after the closing time still counts for that day. Later = Out is MISSING."),
    ("Early disarm window (hours)", "3.5", "A disarm up to this many hours before the opening time counts as that day's OPEN. Earlier = it belongs to the previous day."),
    ("Fortnight starts on", "2026-10-05", "A Monday. Every Excel holds two weeks counted from this date: 05-10 TO 18-10, 19-10 TO 01-11, ... "
     "The file is named by the month of its SECOND Monday (a fifth week therefore starts the next month's file)."),
    ("Keep downloaded reports", "Yes", "Yes = the report PDFs are kept in <month folder>\\Downloaded Reports."),
    ("Schedule day", "Monday", "The automatic security run (panel Setup step 4)."),
    ("Schedule time", "09:00", "24-hour time, e.g. 09:00."),
]
ABOUT = [
    "SECURITY SETTINGS - the bot reads this file at every security run.",
    "",
    "Sites: one row per site, grouped by company. Client No must be exactly as printed in the alarm reports (e.g. 0989-25, 5652-SA).",
    "   Opening / closing times per weekday, written like 5:00 AM and 10:00 PM. A closing time of 12:00 AM means midnight.",
    "   Both times empty for a day = the site is closed that day (its OPEN and CLOSE are highlighted).",
    "   Active = No: the site is left out of the Excel. 24/7 = Yes: the site is listed but its times are not checked.",
    "   A new site: add a row under its company. A report with a Client No that is not here is skipped and named in the run summary.",
    "Run: the rules and the schedule (see the notes in column C).",
    "Save and close this file before a run.",
]

TIME = re.compile(r"(\d{1,2})(?:\s*[:.]\s*(\d{2}))?\s*:?\s*([AaPp])\.?\s*[Mm]")


def path():
    return paths.p(FILE_NAME)


class SecurityError(Exception):
    pass


def parse_time(text, closing=False):
    """'5:00 AM' -> 300 (minutes after midnight). As a CLOSING time, 12:00 AM = 1440 (midnight at the END of the day)."""
    if text is None or str(text).strip() == "":
        return None
    if isinstance(text, datetime.time):
        minutes = text.hour * 60 + text.minute
    else:
        m = TIME.search(str(text))
        if not m:
            return None
        h, mi, ap = int(m.group(1)), int(m.group(2) or 0), m.group(3).upper()
        if not (1 <= h <= 12 and 0 <= mi < 60):
            return None
        minutes = (h % 12 + (12 if ap == "P" else 0)) * 60 + mi
    if closing and minutes == 0:
        return 1440
    return minutes


def show_time(minutes):
    """300 -> '5:00 AM'; 1440 -> '12:00 AM'."""
    if minutes is None:
        return ""
    m = minutes % 1440
    h, mi = divmod(m, 60)
    return f"{(h % 12) or 12}:{mi:02d} {'AM' if h < 12 else 'PM'}"


def timing_text(open_m, close_m):
    """'04:00 AM - 12:00 AM' (the agreed single format)."""
    def two(m):
        t = show_time(m)
        return t if len(t.split(":")[0]) == 2 else "0" + t
    return f"{two(open_m)} - {two(close_m)}"


def split_timing(text):
    """'05:30AM - 09:30PM' -> (330, 1290); '5 AM - 10 PM' -> (300, 1320); anything unreadable -> (None, None)."""
    found = TIME.findall(str(text or ""))
    if len(found) < 2:
        return None, None
    a = "%s:%s %sM" % (found[0][0], found[0][1] or "00", found[0][2])
    b = "%s:%s %sM" % (found[1][0], found[1][1] or "00", found[1][2])
    return parse_time(a), parse_time(b, closing=True)


def create_default_file(target=None, sites=()):
    target = target or path()
    wb = Workbook()
    about = wb.active
    about.title = "About"
    for line in ABOUT:
        about.append([line])
    about.column_dimensions["A"].width = 130
    about["A1"].font = Font(bold=True, size=13)
    ws = wb.create_sheet("Sites")
    ws.append(SITE_HEADERS)
    for row in sites:
        ws.append(row)
    widths = [26, 6, 34, 11, 40] + [10] * 14 + [8, 7]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
        c = ws.cell(1, i)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = "F2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(SITE_HEADERS))}{max(ws.max_row, 2)}"
    run = wb.create_sheet("Run")
    run.append(["Setting", "Value", "Notes"])
    for row in RUN_DEFAULTS:
        run.append(list(row))
    for i, w in enumerate([34, 16, 120], start=1):
        run.column_dimensions[get_column_letter(i)].width = w
        run.cell(1, i).font = Font(bold=True, color="FFFFFF")
        run.cell(1, i).fill = PatternFill("solid", fgColor="1F4E78")
    wb.save(target)
    return target


def _text(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, datetime.time):
        return show_time(v.hour * 60 + v.minute)
    return str(v).strip()


def load(target=None):
    """{'sites': [site dicts in file order], 'run': {setting: value}}. Raises SecurityError when the file is missing or open."""
    target = target or path()
    if not os.path.exists(target):
        create_default_file(target)
    try:
        wb = load_workbook(target, data_only=True, read_only=True)
    except PermissionError:
        raise SecurityError(f"{FILE_NAME} is open in Excel. Save and close it, then run again.")
    except Exception as e:
        raise SecurityError(f"{FILE_NAME} cannot be read ({e}).")
    try:
        run = {k: d for k, d, _n in RUN_DEFAULTS}
        if "Run" in wb.sheetnames:
            for r in wb["Run"].iter_rows(min_row=2, values_only=True):
                if r and r[0]:
                    run[str(r[0]).strip()] = _text(r[1]) or run.get(str(r[0]).strip(), "")
        sites = []
        if "Sites" not in wb.sheetnames:
            raise SecurityError(f"{FILE_NAME} has no 'Sites' sheet.")
        rows = list(wb["Sites"].iter_rows(values_only=True))
        head = [str(h or "").strip() for h in rows[0]] if rows else []
        col = {h: i for i, h in enumerate(head)}
        for r in rows[1:]:
            get = lambda h: _text(r[col[h]]) if h in col and col[h] < len(r) else ""
            client = get("Client No").upper()
            if not client or not get("Site name"):
                continue
            hours = {}
            for i, d in enumerate(DAYS):
                o, c = parse_time(get(f"{d} open")), parse_time(get(f"{d} close"), closing=True)
                hours[i] = (o, c) if o is not None and c is not None else None
            sites.append({"company": get("Company"), "no": get("No"), "name": get("Site name"), "client": client,
                          "address": get("Address"), "hours": hours,
                          "active": get("Active").lower() not in ("no", "n", "false", "0"),
                          "always_open": get("24/7").lower() in ("yes", "y", "true", "1")})
        return {"sites": sites, "run": run}
    finally:
        wb.close()


def number(run, key, default):
    try:
        return float(str(run.get(key, default)).strip())
    except ValueError:
        raise SecurityError(f"'{key}' in {FILE_NAME} (sheet Run) must be a number.")


def import_from_workbook(source, target=None):
    """Builds Security_Settings.xlsx from a filled STORE OPEN & CLOSE workbook (the sheet with Client No and ADDRESS):
    company groups, site number, name, Client No, address, and each site's LATEST opening hours per weekday."""
    wb = load_workbook(source, data_only=True)
    ws = None
    for s in wb.worksheets:
        if any(str(s.cell(2, c).value or "").strip().lower() == "client no" for c in range(1, 8)):
            ws = s
            break
    if ws is None:
        raise SecurityError("No sheet with a 'Client No' column was found in that workbook.")
    label_col = next(c for c in range(1, 10) if str(ws.cell(2, c).value or "").strip().lower() == "date")
    day_cols = [(c, str(ws.cell(3, c).value or "").strip().lower()[:3]) for c in range(label_col + 1, ws.max_column + 1)
                if str(ws.cell(3, c).value or "").strip()]
    order = [d.lower() for d in DAYS]
    company, out = "", []
    for r in range(5, ws.max_row + 1):
        if ws.cell(r, 1).value:
            company = str(ws.cell(r, 1).value).strip()
        if str(ws.cell(r, label_col).value or "").strip() != "Timing":
            continue
        latest = {}
        for c, day in day_cols:                              # the right-most (latest) timing of each weekday wins
            o, cl = split_timing(ws.cell(r, c).value)
            if day in order and o is not None:
                latest[order.index(day)] = (o, cl)
        row = [company, _text(ws.cell(r, label_col - 4).value), _text(ws.cell(r, label_col - 3).value),
               _text(ws.cell(r, label_col - 2).value), _text(ws.cell(r, label_col - 1).value)]
        for i in range(7):
            o, cl = latest.get(i, (None, None))
            row += [show_time(o), show_time(cl)]
        row += ["Yes", "No"]
        if row[2]:
            out.append(row)
    return create_default_file(target, out), len(out)
