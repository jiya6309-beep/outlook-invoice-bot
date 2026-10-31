"""Excel files for people (the database is the real record; these are views of it).

    Reports/Record_YYYY-MM.xlsx        every document of a month (Record) and the ones that need a person (Needs attention)
    Reports/To_Fix.xlsx                what needs a person - staff TYPE the fixes in the yellow columns (drop-down lists)
    Reports/Report_<from>_<to>.xlsx    the same report for any From/To period, on request
    Reports/Shadow_Report_<time>.xlsx  what a shadow run would have done

Big sheets are written with openpyxl's streaming writer (a month of 55,000 rows takes about 5 seconds and 45 MB).
If a file is open in Excel the new version is saved next to it with '_new' in the name - the data is safe in the database."""
import datetime
import os
import stat

from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.cell.cell import Cell
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.worksheet.datavalidation import DataValidation

from . import db, paths, settings, sites

NAVY, BLUE, LIGHT, PALE = "1F3864", "2F5597", "DDEBF7", "F2F6FC"
GREEN, AMBER, RED, GREY = "C6EFCE", "FFEB9C", "FFC7CE", "F2F2F2"
# a row's colour shows where its FILE went when the invoice date and the arrival month differ (see the Remarks column)
CASE_FILL = {"past": "FFF2A8", "far": "E1D5F0", "future": "F8CBAD", "moved": "F8CBAD", "early": "BDD7EE"}
STATUS_FILL = {"Filed": GREEN, "Parked": LIGHT, "Unassigned": AMBER, "Unrecognised": RED}
MONEY = "#,##0.00;[Red]-#,##0.00"


def reports_dir():
    folder = settings.reports_root()
    os.makedirs(folder, exist_ok=True)
    return folder


def _date(value):
    try:
        return datetime.date.fromisoformat(str(value)[:10])
    except ValueError:
        return None


def _save(wb, target):
    """Save; if the file is open in Excel save beside it. Returns the path actually written."""
    try:
        wb.save(target)
        return target
    except PermissionError:
        base, ext = os.path.splitext(target)
        alt = f"{base}_new{ext}"
        wb.save(alt)
        return alt


def _print_setup(ws, landscape=True):
    """Prints on one page wide (as many pages tall as needed)."""
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.paperSize = 9                                   # A4
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)


def _mode_label():
    try:
        return settings.destination_mode().upper()
    except settings.SettingsError:
        return "?"


# ----------------------------------------------------------------------------------------- styling helpers

def S(ws, value=None, *, bold=False, size=11, color="000000", fill=None, fmt=None, h=None, wrap=False, italic=False, line=False):
    """A styled cell for a write-only sheet."""
    c = WriteOnlyCell(ws, value=value)
    c.font = Font(name="Calibri", bold=bold, size=size, color=color, italic=italic)
    if fill:
        c.fill = PatternFill("solid", fgColor=fill)
    if fmt:
        c.number_format = fmt
    if h or wrap:
        c.alignment = Alignment(horizontal=h, vertical="center", wrap_text=wrap)
    if line:
        c.border = Border(bottom=Side(style="medium", color=BLUE))
    return c


def _head(ws, headers):
    return [S(ws, h, bold=True, color="FFFFFF", fill=NAVY, wrap=True, h="center") for h in headers]


def _normal_header(ws, headers, fill=NAVY):
    ws.append(headers)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=fill)
        c.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 30


def _normal_sheet(wb, title, headers, rows, widths=None, formats=None, first=False, tab=None):
    """A normal (small) sheet with a styled header, filter and frozen top row."""
    ws = wb.active if first else wb.create_sheet()
    ws.title = title
    _normal_header(ws, headers)
    for r in rows:
        ws.append(r)
    for i, w in enumerate(widths or [], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for col, fmt in (formats or {}).items():
        for cell in ws[get_column_letter(col)][1:]:
            cell.number_format = fmt
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(ws.max_row, 2)}"
    if tab:
        ws.sheet_properties.tabColor = tab
    _print_setup(ws)
    return ws


def _status_rules(ws, column_letter, last_row):
    """Colour a Status column: Filed green, Parked blue, Unassigned amber, Unrecognised red."""
    for status, fill in STATUS_FILL.items():
        ws.conditional_formatting.add(f"{column_letter}2:{column_letter}{last_row}", CellIsRule(
            operator="equal", formula=[f'"{status}"'], fill=PatternFill("solid", bgColor=fill)))


# ----------------------------------------------------------------------------------------- the report

RECORD_HEADERS = ["Ref", "Supplier", "Document type", "Number", "Date", "Amount incl GST", "GST", "Amount excl GST", "Company folder", "Site",
                  "Status", "Remarks", "File name", "Saved in", "Open folder", "Mailbox", "Received", "Downloaded by bot", "Run marker",
                  "Warnings / reason"]
RECORD_WIDTHS = [10, 38, 17, 18, 12, 16, 11, 16, 28, 30, 13, 22, 44, 70, 13, 30, 17, 18, 36, 70]


def _stamp(value):
    """'2026-10-15 09:30:12' -> '15-10-2026 09:30' (the time the bot saved the document)."""
    text = str(value or "")
    try:
        return datetime.datetime.fromisoformat(text[:19]).strftime("%d-%m-%Y %H:%M")
    except ValueError:
        return text[:16]


def _source(rows):
    return rows if hasattr(rows, "attention") else db.ListSource(rows)


def _record_workbook(rows):
    """A streaming workbook with the Record (every document) and Needs attention sheets.
    `rows` is a list or a db.DocSource; rows are read one at a time, so memory stays flat however big the month is."""
    rows = _source(rows)
    total = rows.count()
    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Record")
    ws.sheet_properties.tabColor = BLUE
    _print_setup(ws)
    ws.print_title_rows = "1:1"
    for i, w in enumerate(RECORD_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    last = max(total + 1, 2)
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(RECORD_HEADERS))}{last}"
    ws.append(_head(ws, RECORD_HEADERS))
    markers = db.run_markers(rows.conn) if getattr(rows, "conn", None) is not None else {}

    def record_row(d, case="", remarks=None):
        fill = PatternFill("solid", fgColor=CASE_FILL[case]) if case in CASE_FILL else None
        date = WriteOnlyCell(ws, _date(d["doc_date"]))
        date.number_format = "dd/mm/yyyy"
        amount, gst, ex = WriteOnlyCell(ws, d["amount"]), WriteOnlyCell(ws, d["gst"]), WriteOnlyCell(ws, d["ex_amount"])
        amount.number_format = gst.number_format = ex.number_format = MONEY            # an empty cell means "not found"; 0.00 is a real value
        folder = (d["dest_dir"] or "").replace('"', "")
        row = [d["ref"], d["supplier"], d["doc_type"], d["number"], date, amount, gst, ex, d["company_folder"], d["site"],
               d["status"], d["remarks"] or "" if remarks is None else remarks, d["dest_name"] or "", d["dest_dir"] or "",
               f'=HYPERLINK("{folder}","Open folder")' if folder else "",
               d["mailbox"], str(d["received"] or "")[:16], _stamp(d.get("created_at")), markers.get(d["ref"], ""),
               d["warnings"] or d["reason"] or ""]
        marked = d["ref"] in markers and remarks is None
        if fill or marked:
            for i, c in enumerate(row):
                if not isinstance(c, Cell):
                    c = WriteOnlyCell(ws, c)
                if fill:
                    c.fill = fill
                if marked:                                  # the last e-mail of the latest run: bold with a box, easy to find
                    c.font = Font(bold=True)
                    c.border = Border(top=Side(style="medium"), bottom=Side(style="medium"))
                row[i] = c
            for c in (date, amount, gst, ex):
                if fill:
                    c.fill = fill
                if marked:
                    c.font = Font(bold=True)
                    c.border = Border(top=Side(style="medium"), bottom=Side(style="medium"))
        ws.append(row)

    for d in rows.rows():
        record_row(d, d.get("month_case") or "")
    for d in rows.early():                                     # arrived early: moved here from an earlier month's folder - listed, not counted
        moved = _date(d.get("moved_on"))
        note = f"ARRIVED EARLY: arrived {str(d['received'] or '')[:10]}, filed in the earlier month's folder and moved here" + (f" on {moved:%d-%m-%Y}" if moved else "")
        record_row(d, "early", note)
    _status_rules(ws, "K", last)
    ws.conditional_formatting.add(f"A2:T{last}", FormulaRule(formula=['AND(MOD(ROW(),2)=0,$L2="")'], fill=PatternFill("solid", bgColor=PALE)))

    wa = wb.create_sheet("Needs attention")
    wa.sheet_properties.tabColor = "C00000"
    _print_setup(wa)
    for i, w in enumerate([10, 38, 17, 18, 13, 30, 90, 70], start=1):
        wa.column_dimensions[get_column_letter(i)].width = w
    wa.freeze_panes = "A2"
    wa.append(_head(wa, ["Ref", "Supplier", "Document type", "Number", "Status", "Mailbox", "Reason", "Where the file is now"]))
    todo = list(rows.attention())                              # only the documents that need a person: always few
    wa.auto_filter.ref = f"A1:H{max(len(todo) + 1, 2)}"
    for d in todo:
        wa.append([d["ref"], d["supplier"], d["doc_type"], d["number"], d["status"], d["mailbox"], d["reason"], d["current_path"]])
    _status_rules(wa, "E", max(len(todo) + 1, 2))
    return wb


def write_record(rows, target):
    """The report of a month / period: the Record (every document) and Needs attention. Streaming writer (fast, low memory)."""
    return _save(_record_workbook(rows), target)


def write_archive(rows, corrections, target, title):
    """A read-only snapshot of a period's records: Record, Needs attention, Corrections and an About note."""
    rows = _source(rows)
    wb = _record_workbook(rows)
    wc = wb.create_sheet("Corrections")
    wc.sheet_properties.tabColor = "7F7F7F"
    for i, w in enumerate([10, 16, 40, 40, 20, 20], start=1):
        wc.column_dimensions[get_column_letter(i)].width = w
    wc.freeze_panes = "A2"
    wc.append(_head(wc, ["Ref", "Field", "Old value", "New value", "Changed by", "When"]))
    for c in corrections:
        wc.append([c["ref"], c["field"], c["old"], c["new"], c["by"], c["at"]])
    wa = wb.create_sheet("About")
    wa.column_dimensions["A"].width = 130
    for line in (title, f"Made {datetime.datetime.now():%Y-%m-%d %H:%M} by the Outlook Invoice Bot.",
                 "This is a read-only COPY of the records of the period. The live database still holds them; nothing was deleted.",
                 f"{rows.count()} documents, {len(corrections)} corrections. Status as it was when the file was made."):
        wa.append([line])
    tmp = target + ".part"
    wb.save(tmp)
    os.replace(tmp, target)
    try:
        os.chmod(target, stat.S_IREAD)                      # read-only: an archive is not edited
    except OSError:
        pass
    return target


def record_path(year, month):
    return os.path.join(reports_dir(), f"Record_{year:04d}-{month:02d}.xlsx")


def write_monthly(conn, year, month, basis="arrival"):
    return write_record(db.MonthSource(conn, year, month, basis), record_path(year, month))


def write_range(conn, start, end, basis="invoice"):
    source = db.DocSource(conn, *db.between_clause(start, end, basis))
    return write_record(source, os.path.join(reports_dir(), f"Report_{start}_{end}.xlsx")), source.count()


# ----------------------------------------------------------------------------------------- A/C Excel
AC_HEADERS = ["NAME OF COMPANY", "NAME OF SITE", "NAME OF SUPPLIER", "DOCKET DT.", "DOCKET NO.", "INVOICE DT.", "INVOICE NO.",
              "AMOUNT", "SHARE POINT", "REMARKS"]
AC_WIDTHS = [30, 30, 34, 13, 18, 13, 18, 14, 16, 40]
AC_TYPES = ("Invoice", "Credit note", "POD")


def ac_rows(conn, start, end):
    """Filed invoices, credit notes and dockets whose document date (or arrival date when the document has none) is in the period,
    sorted by company, site, date."""
    s, e = str(start)[:10], str(end)[:10]
    q = (f"SELECT * FROM docs WHERE status=? AND doc_type IN ({','.join('?' * len(AC_TYPES))}) "
         "AND substr(COALESCE(NULLIF(doc_date,''), received),1,10) BETWEEN ? AND ? "
         "ORDER BY company_folder, site, COALESCE(NULLIF(doc_date,''), received), id")
    for r in conn.execute(q, (db.FILED,) + AC_TYPES + (s, e)):
        yield dict(r)


def write_ac(conn, start, end):
    """The accounts sheet in the company's own A/C format (the MASTER layout): one row per document; a docket fills only the docket
    columns, an invoice / credit note only the invoice columns. AMOUNT = total incl. GST (credit notes negative, empty for dockets).
    SHARE POINT = a link that opens the filed file. Returns (path, number of rows)."""
    wb = Workbook(write_only=True)
    ws = wb.create_sheet("AC")
    _print_setup(ws)
    ws.print_title_rows = "1:1"
    for i, w in enumerate(AC_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    thin = Side(style="thin", color="000000")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    head = []
    for h in AC_HEADERS:
        c = WriteOnlyCell(ws, h)
        c.font, c.border = Font(bold=True), box
        c.alignment = Alignment(horizontal="center", vertical="center")
        head.append(c)
    ws.append(head)
    n = 0
    for d in ac_rows(conn, start, end):
        is_pod = d["doc_type"] == "POD"
        when = _date(d["doc_date"]) or _date(d["received"])
        path = os.path.join(d["dest_dir"] or "", d["dest_name"] or "").replace('"', "")
        values = [d["company_folder"] or "", d["site"] if d["site"] and d["site"] != "UNASSIGNED" else "", d["supplier"] or "",
                  when if is_pod else None, (d["number"] or "") if is_pod else "",
                  None if is_pod else when, "" if is_pod else (d["number"] or ""),
                  None if is_pod else d["amount"],
                  f'=HYPERLINK("{path}","Open file")' if d["dest_name"] else "", d["remarks"] or ""]
        row = []
        for i, v in enumerate(values):
            c = WriteOnlyCell(ws, v)
            c.border = box
            if i in (3, 5):
                c.number_format = "dd/mm/yyyy"
            if i == 7:
                c.number_format = MONEY
            if i == 8 and v:
                c.font = Font(color="0563C1", underline="single")
            row.append(c)
        ws.append(row)
        n += 1
    ws.auto_filter.ref = f"A1:J{max(n + 1, 2)}"
    target = os.path.join(reports_dir(), f"AC_{str(start)[:10]}_{str(end)[:10]}.xlsx")
    return _save(wb, target), n


# ----------------------------------------------------------------------------------------- To_Fix

TO_FIX_HEADERS = ["Ref", "Status", "Reason", "Received", "Mailbox", "Supplier", "Document type", "Number", "Date",
                  "Amount", "Where the file is now", "Company folder (TYPE HERE)", "Supplier (fix)", "Number (fix)",
                  "Document type (fix)", "Site folder (TYPE HERE)", "Open folder", "Statement missing data", "Deliver-to as printed"]
TO_FIX_INPUTS = {12: "company_folder", 13: "supplier", 14: "number", 15: "doc_type", 16: "site_folder"}

HOW_TO_FIX = [
    ("HOW TO USE THIS FILE", None),
    ("1", "Open the sheet  'To fix'.  Each row is one document the bot could not file by itself. 'Reason' tells you why."),
    ("2", "Click 'Open folder' to see the file. Open it and look at the document."),
    ("3", "Type the correction in the YELLOW cells only:"),
    ("", "•  Site folder (TYPE HERE): pick the site from the drop-down list. This is what a file needs when its status says 'site could not be determined'."),
    ("", "•  Company folder (TYPE HERE): pick the folder from the drop-down list, when the reason says the company folder could not be determined."),
    ("", "•  Statement missing data: an orange row. A statement or order confirmation was filed (renamed with NoSite) because no site name was found in it. Nothing to type: rename the file by hand if you want the site in its name."),
    ("", "•  Supplier (fix), Number (fix), Document type (fix): only for documents the bot could not read (status Unrecognised)."),
    ("4", "Save and CLOSE this file."),
    ("5", "At the next run the bot moves each file to its correct folder, records who changed what and when, and removes the row from this list."),
    ("", ""),
    ("COLOURS", None),
    ("", "Amber = waiting (known document, destination not available yet: the bot keeps retrying)."),
    ("", "Red = unrecognised (a person must look)."),
    ("", ""),
    ("OTHER SHEETS", None),
    ("", "Image-only emails: invoices that arrived as a picture. Open the e-mail in Outlook (orange category) and handle it by hand."),
    ("", "Overdue-urgent emails: e-mails with words like overdue / reminder / stop supply. They carry a red flag in Outlook."),
    ("", "New sites: documents the bot filed in the COMPANY folder because it could not find the site (Remarks says SITE NOT FOUND). Pick the site in the yellow column; at the next run the bot moves the file into that site's folder. A site that is not on the list is NOT created: the document stays where it is and Remarks says SITE NAME NOT ON LIST."),
    ("", "Sites by company: every company and its sites, for reference. To add a site, add it to the Sites sheet of Bot_Settings.xlsx."),
    ("", "Learned sites: sites the bot remembers from your fixes (a supplier's deliver-to number / address, or a Fleet Card merchant number). Type YES in the yellow column to make it forget one."),
    ("", "Deliver-to as printed (sheet To fix): the ship-to address exactly as printed, when the bot did not know that site. Choose the site once in 'Site folder (TYPE HERE)': the bot remembers it for the next documents."),
    ("", "New suppliers: supplier names the bot saw for the first time. If one is spelled wrongly, add a row to the Suppliers sheet of Bot_Settings.xlsx."),
]


NEW_SITES_HEADERS = ["Ref", "Company", "Supplier", "Number", "Date", "What the document says about the site", "Where the file is now",
                     "Site (TYPE HERE)", "Open folder"]


LEARNED_HEADERS = ["Id", "Supplier", "What the document says", "Remembered by", "Site", "Company folder", "Learned on", "By",
                   "Forget (type YES)"]


def _ship_text(value):
    """The stored ship-to / merchant text as staff should read it."""
    text = str(value or "")
    if text.startswith("merchant:"):
        return "Merchant number " + text[9:]
    return text.replace(" || ", "  /  ").replace(" | ", ", ")


def read_forget(path=None):
    """Ids of learned sites marked 'Forget (type YES)' in To_Fix > Learned sites."""
    path = path or os.path.join(reports_dir(), "To_Fix.xlsx")
    if not os.path.exists(path):
        return []
    try:
        wb = load_workbook(path, data_only=True)
    except Exception:
        return []
    out = []
    try:
        for row in wb["Learned sites"].iter_rows(min_row=2, values_only=True):
            if row and row[0] and len(row) >= 9 and str(row[8] or "").strip().lower() in ("yes", "y", "x"):
                try:
                    out.append(int(row[0]))
                except (TypeError, ValueError):
                    pass
    except KeyError:
        return []
    finally:
        wb.close()
    return out


def to_fix_conflicts():
    """Copies of To_Fix.xlsx made by OneDrive when two PCs saved it at once ('To_Fix-PCNAME.xlsx', 'To_Fix (1).xlsx').
    Fixes typed in such a copy are NOT read."""
    folder = reports_dir()
    try:
        return sorted(f for f in os.listdir(folder) if f.lower().startswith("to_fix") and f.lower().endswith(".xlsx")
                      and f.lower() not in ("to_fix.xlsx", "to_fix_new.xlsx") and not f.startswith("~$"))
    except OSError:
        return []


def site_not_found_docs(conn):
    """Filed documents that went to the company folder because no site could be found (Remarks says SITE NOT FOUND)."""
    return [dict(r) for r in conn.execute(
        "SELECT * FROM docs WHERE status=? AND remarks LIKE '%SITE NOT FOUND%' ORDER BY company_folder, id", (db.FILED,))]


def _new_sites_sheets(wb, conn):
    """'New sites': the documents filed without a site, grouped by company - type the site next to each. 'Sites by company': the full list."""
    docs = site_not_found_docs(conn)
    rows = []
    for d in docs:
        folder = (d["dest_dir"] or "").replace('"', "")
        rows.append([d["ref"], d["company_folder"], d["supplier"], d["number"], _date(d["doc_date"]),
                     d["site_source"] or "", d["current_path"], "", f'=HYPERLINK("{folder}","Open folder")' if folder else ""])
    ws = _normal_sheet(wb, "New sites", NEW_SITES_HEADERS, rows, [10, 28, 34, 16, 12, 70, 70, 34, 14], {5: "dd/mm/yyyy"}, tab="BF8F00")
    ws.cell(1, 8).fill = PatternFill("solid", fgColor="BF8F00")
    for r in range(2, max(ws.max_row, 2) + 200):
        ws.cell(r, 8).fill = PatternFill("solid", fgColor="FFF2CC")
    wl = wb.create_sheet("Sites by company")
    wl.sheet_properties.tabColor = "548235"
    _normal_header(wl, ["Company folder", "Company", "Site", "Site folder"])
    by_company = {}
    for s_ in sites.active_sites():
        by_company.setdefault(sites.company_of(s_) or s_.get("company") or "(no company)", []).append(s_)
    for c in settings.rows("Companies"):
        if not settings.yes(c.get("Active", ""), default=True):
            continue
        listed = sorted(by_company.pop(c["Company name"], []), key=lambda x: x["name"])
        wl.append([c["Folder name"], c["Company name"], f"{len(listed)} site(s)", ""])
        for cell in wl[wl.max_row]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor=BLUE)
        for s_ in listed:
            wl.append(["", "", s_["name"], sites.folder_for(s_["name"])])
    for name, listed in sorted(by_company.items()):                  # sites whose company is not in the Companies sheet
        wl.append(["(not in Companies sheet)", name, f"{len(listed)} site(s)", ""])
        for s_ in sorted(listed, key=lambda x: x["name"]):
            wl.append(["", "", s_["name"], sites.folder_for(s_["name"])])
    for i, w in enumerate([30, 36, 40, 40], start=1):
        wl.column_dimensions[get_column_letter(i)].width = w
    wl.freeze_panes = "A2"
    _print_setup(wl)


def write_to_fix(conn, new_suppliers=()):
    wb = Workbook()
    howto = wb.active
    howto.title = "How to use"
    howto.sheet_properties.tabColor = "548235"
    howto.sheet_view.showGridLines = False
    howto.column_dimensions["A"].width = 6
    howto.column_dimensions["B"].width = 130
    _print_setup(howto)
    for key, text in HOW_TO_FIX:
        if text is None:
            howto.append([key])
            howto.cell(howto.max_row, 1).font = Font(bold=True, size=14, color=NAVY)
        else:
            howto.append([key, text])
            howto.cell(howto.max_row, 1).font = Font(bold=True, color=BLUE)
            howto.cell(howto.max_row, 2).alignment = Alignment(wrap_text=True, vertical="top")

    docs = db.docs_by_status(conn, db.UNASSIGNED, db.UNRECOGNISED)
    docs += [d for d in db.docs_by_status(conn, db.PARKED) if settings.MISSING_SITE in (d.get("remarks") or "")]
    rows = []
    for d in docs:
        folder = (d["dest_dir"] or "").replace('"', "")
        rows.append([d["ref"], d["status"], d["reason"], str(d["received"] or "")[:16], d["mailbox"], d["supplier"], d["doc_type"],
                     d["number"], _date(d["doc_date"]), d["amount"], d["current_path"], "", "", "", "", "",
                     f'=HYPERLINK("{folder}","Open folder")' if folder else "",
                     "Site not found in this document" if settings.MISSING_SITE in (d.get("remarks") or "") else "",
                     _ship_text(d.get("ship_to"))])
    ws = _normal_sheet(wb, "To fix", TO_FIX_HEADERS, rows,
                       [10, 13, 70, 16, 30, 34, 16, 16, 12, 12, 70, 28, 28, 18, 18, 30, 14, 34, 60], {9: "dd/mm/yyyy", 10: MONEY}, tab="C00000")
    last = max(ws.max_row, 2)
    for r in range(2, last + 1):
        ws.cell(r, 2).fill = PatternFill("solid", fgColor=STATUS_FILL.get(ws.cell(r, 2).value, GREY))
        ws.cell(r, 3).alignment = Alignment(wrap_text=True, vertical="top")
        if ws.cell(r, 18).value:                               # a statement / order confirmation filed with NoSite: the whole row is highlighted
            for c in range(1, 19):
                if c not in TO_FIX_INPUTS:
                    ws.cell(r, c).fill = PatternFill("solid", fgColor="F8CBAD")
            ws.cell(r, 18).font = Font(bold=True, color="C00000")
    ws.cell(1, 18).fill = PatternFill("solid", fgColor="C00000")
    for col in TO_FIX_INPUTS:
        ws.cell(1, col).fill = PatternFill("solid", fgColor="BF8F00")
        for r in range(2, last + 200):
            ws.cell(r, col).fill = PatternFill("solid", fgColor="FFF2CC")
    emails = [dict(r) for r in conn.execute("SELECT * FROM emails ORDER BY received DESC")]
    _normal_sheet(wb, "Image-only emails", ["Received", "Mailbox", "Subject", "Pictures"],
                  [[e["received"][:16], e["mailbox"], e["subject"], e["detail"]] for e in emails if e["outcome"] == "image-only"],
                  [18, 30, 80, 50], tab="ED7D31")
    _normal_sheet(wb, "Overdue-urgent emails", ["Received", "Mailbox", "Subject", "Blocked word"],
                  [[e["received"][:16], e["mailbox"], e["subject"], e["detail"]] for e in emails if e["outcome"] == "blocked"],
                  [18, 30, 80, 24], tab="C00000")
    _normal_sheet(wb, "New suppliers", ["Supplier name as read", "How it was found"], [list(x) for x in new_suppliers], [50, 24], tab="7F7F7F")
    _new_sites_sheets(wb, conn)
    learned = db.learned_all(conn)
    wlrn = _normal_sheet(wb, "Learned sites", LEARNED_HEADERS,
                         [[x["id"], x["supplier"], _ship_text(x["example"] or x["key"]), x["key"], x["site"], x["company_folder"],
                           str(x["at"] or "")[:16], x["by"], ""] for x in learned], [8, 30, 60, 40, 34, 28, 17, 22, 18], tab="548235")
    wlrn.cell(1, 9).fill = PatternFill("solid", fgColor="BF8F00")
    for r in range(2, max(wlrn.max_row, 2) + 1):
        wlrn.cell(r, 9).fill = PatternFill("solid", fgColor="FFF2CC")
    # drop-down lists (typing something else is still allowed)
    lists = wb.create_sheet("Lists")
    lists.sheet_state = "hidden"
    folders = [r["Folder name"] for r in settings.rows("Companies") if r["Folder name"]]
    folders += sorted({r["Folder override"] for r in settings.rows("Suppliers") if r["Folder override"]} - set(folders))
    lists.append(["Company folders", "Document types", "Site folders"])
    types = [r["Document type"] for r in settings.rows("DocTypes")]
    site_folders = sites.folder_names()
    for i in range(max(len(folders), len(types), len(site_folders))):
        lists.append([folders[i] if i < len(folders) else None, types[i] if i < len(types) else None,
                      site_folders[i] if i < len(site_folders) else None])
    ns = wb["New sites"]
    for sheet, col, rng, end in ((ws, 12, f"=Lists!$A$2:$A${len(folders) + 1}", last + 200), (ws, 15, f"=Lists!$B$2:$B${len(types) + 1}", last + 200),
                                 (ws, 16, f"=Lists!$C$2:$C${max(len(site_folders), 1) + 1}", last + 200),
                                 (ns, 8, f"=Lists!$C$2:$C${max(len(site_folders), 1) + 1}", max(ns.max_row, 2) + 200)):
        dv = DataValidation(type="list", formula1=rng, allow_blank=True, showErrorMessage=True, errorStyle="warning",
                            errorTitle="Not in the list", error="This value is not in the list. Use it anyway?")
        sheet.add_data_validation(dv)
        dv.add(f"{get_column_letter(col)}2:{get_column_letter(col)}{end}")

    wb.active = wb.sheetnames.index("To fix")
    return _save(wb, os.path.join(reports_dir(), "To_Fix.xlsx"))


def read_new_sites(path=None):
    """[(ref, typed site)] from the 'New sites' sheet."""
    path = path or os.path.join(reports_dir(), "To_Fix.xlsx")
    if not os.path.exists(path):
        return []
    try:
        wb = load_workbook(path, data_only=True)
    except Exception:
        return []
    out = []
    try:
        for row in wb["New sites"].iter_rows(min_row=2, values_only=True):
            if row and row[0] and len(row) >= 8 and row[7] not in (None, ""):
                out.append((str(row[0]).strip(), str(row[7]).strip()))
    except KeyError:
        return []
    finally:
        wb.close()
    return out


def read_to_fix(path=None):
    """The fixes staff typed: [{'ref', 'company_folder', 'supplier', 'number', 'doc_type'}]."""
    path = path or os.path.join(reports_dir(), "To_Fix.xlsx")
    if not os.path.exists(path):
        return []
    try:
        wb = load_workbook(path, data_only=True)
    except Exception:
        return []
    edits = []
    try:
        for row in wb["To fix"].iter_rows(min_row=2, values_only=True):
            if not row or not row[0]:
                continue
            edit = {"ref": str(row[0]).strip()}
            for col, field in TO_FIX_INPUTS.items():
                v = row[col - 1] if len(row) >= col else None
                if v not in (None, ""):
                    edit[field] = str(v).strip()
            if len(edit) > 1:
                edits.append(edit)
    except KeyError:
        return []
    finally:
        wb.close()
    return edits


# ----------------------------------------------------------------------------------------- shadow

SHADOW_HEADERS = ["Mailbox", "Received", "Subject", "What the bot would do in Outlook", "Attachment", "Document type", "Supplier",
                  "Number", "Date", "Amount incl GST", "GST", "Amount excl GST", "Company folder", "Site", "Would be saved in", "As file name",
                  "Result", "Remarks", "Reason / warnings"]


def _tile_row(ws, row, tiles):
    """KPI tiles across columns A..F of a normal sheet: label, big number, caption."""
    for col, (label, value, caption, color) in enumerate(tiles, start=1):
        for r, (text, font) in enumerate([(label, Font(bold=True, size=9, color="595959")),
                                          (value, Font(bold=True, size=20, color=color)),
                                          (caption, Font(italic=True, size=9, color="7F7F7F"))]):
            c = ws.cell(row + r, col, text)
            c.font, c.fill = font, PatternFill("solid", fgColor=LIGHT)
            c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row + 1].height = 34


def write_shadow_report(rows, summary_text, extras):
    """rows = per document / per e-mail 'would do' lines; extras = {'blocked', 'new_suppliers', 'skipped', 'counters', 'mailboxes', 'notes'}."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws.sheet_properties.tabColor = NAVY
    ws.sheet_view.showGridLines = False
    _print_setup(ws)
    for i, w in enumerate([30, 27, 27, 27, 27, 27], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    c = extras.get("counters", {})
    ws.append(["SHADOW RUN - REHEARSAL"] + [None] * 5)
    for col in range(1, 7):
        ws.cell(1, col).fill = PatternFill("solid", fgColor=NAVY)
    ws["A1"].font = Font(bold=True, size=20, color="FFFFFF")
    ws.row_dimensions[1].height = 32
    ws.append([f"Generated {datetime.datetime.now():%Y-%m-%d %H:%M}  |  mode: {_mode_label()}"])
    ws["A2"].font = Font(italic=True, color="595959")
    ws.append(["Nothing was changed in Outlook and no file was written to a destination or holding folder. This is what a real run WOULD do."])
    ws["A3"].font = Font(bold=True, italic=True, color="C00000")
    ws.append([])
    _tile_row(ws, 5, [("E-MAILS READ", f"{c.get('emails scanned', 0):,}", "in the chosen period", NAVY),
                      ("WOULD BE FILED", f"{c.get('Filed', 0):,}", "to company folders", "2E7D32"),
                      ("WOULD BE PARKED", f"{c.get('Parked', 0):,}", "statements, summaries ...", NAVY),
                      ("WOULD WAIT", f"{c.get('Unassigned', 0):,}", "folder not available yet", "B26A00" if c.get("Unassigned") else NAVY),
                      ("UNRECOGNISED", f"{c.get('Unrecognised', 0):,}", "a person must look", "C00000" if c.get("Unrecognised") else NAVY),
                      ("RED-FLAGGED E-MAILS", f"{c.get('blocked', 0):,}", "overdue / urgent / stop supply", "C00000" if c.get("blocked") else NAVY)])
    ws.append([])
    ws.append([])
    ws.append([])
    ws.append(["By mailbox"])
    ws.cell(ws.max_row, 1).font = Font(bold=True, size=14, color=NAVY)
    head = ws.max_row + 1
    ws.append(["Mailbox (company)", "E-mails read", "Period", "Red-flagged", "Document e-mails", "Documents"])
    for col in range(1, 7):
        cell = ws.cell(head, col)
        cell.font, cell.fill, cell.alignment = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor=NAVY), Alignment(horizontal="center")
    for label, n, start, end, kind, blocked, wanted, docs, note in extras.get("mailboxes", []):
        ws.append([label, n, f"{start[:10]} to {end[:10]}", blocked, wanted, docs])
        for col in (2, 4, 5, 6):
            ws.cell(ws.max_row, col).alignment = Alignment(horizontal="center")
    extra_notes = [n for n in extras.get("notes", []) if n]
    if extras.get("new_suppliers"):
        extra_notes.append("First-time supplier names (check them): " + "; ".join(n for n, _ in extras["new_suppliers"][:12]))
    if extra_notes:
        ws.append([])
        for note in extra_notes:
            ws.append([note])
            ws.cell(ws.max_row, 1).font = Font(italic=True, color="595959")
    ws.append([])
    ws.append([summary_text.splitlines()[-1] if summary_text else ""])
    ws.cell(ws.max_row, 1).font = Font(italic=True, size=9, color="7F7F7F")
    sheet = _normal_sheet(wb, "Would do", SHADOW_HEADERS, rows, [28, 16, 60, 40, 40, 16, 34, 16, 12, 14, 10, 14, 26, 28, 70, 44, 14, 20, 80],
                          {9: "dd/mm/yyyy", 10: MONEY, 11: MONEY, 12: MONEY}, tab=BLUE)
    for r in range(2, sheet.max_row + 1):
        result = str(sheet.cell(r, 17).value or "")
        fill = STATUS_FILL.get(result) or (AMBER if result == "duplicate" else None)
        if fill:
            sheet.cell(r, 17).fill = PatternFill("solid", fgColor=fill)
        if "NOT downloaded" in str(sheet.cell(r, 19).value or ""):
            sheet.cell(r, 19).font = Font(color="C00000", bold=True)
    _normal_sheet(wb, "Blocked words", ["Word", "E-mails"], [[k, v] for k, v in sorted(extras.get("blocked", {}).items())], [30, 12], tab="C00000")
    _normal_sheet(wb, "New suppliers", ["Supplier name as read", "How it was found"], [list(x) for x in extras.get("new_suppliers", [])], [50, 24], tab="7F7F7F")
    _normal_sheet(wb, "Skipped with attachments", ["Mailbox", "Received", "Subject", "Attachments"],
                  [list(x) for x in extras.get("skipped", [])], [28, 16, 80, 70], tab="ED7D31")
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    return _save(wb, os.path.join(reports_dir(), f"Shadow_Report_{stamp}.xlsx"))


# ----------------------------------------------------------------------------------------- Fuel invoice list
FUEL_HEADERS = ["Download Date", "Contact", "Invoice Number", "Reference", "Invoice Date", "XERO Bill", "ABA", "QUICK FUEL", "REMARKS"]
FUEL_WIDTHS = [14, 52, 15.5, 44, 12.5, 12.5, 17.5, 18, 30]
MONTH_ABBR = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]


def fuel_list_path(year, month):
    return os.path.join(reports_dir(), f"PI_Invoice_list_{MONTH_ABBR[month - 1]}_{year}.xlsx")


def _new_fuel_book():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(FUEL_HEADERS)
    for i, w in enumerate(FUEL_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
        c = ws.cell(1, i)
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"
    return wb


def append_fuel_list(conn):
    """Adds every filed fuel invoice that is not on its month's list yet at the BOTTOM of
    Reports\\PI_Invoice_list_<MON>_<YEAR>.xlsx (month = the day the bot downloaded it). Rows already there - and the staff
    columns XERO Bill / ABA / QUICK FUEL - are never changed. A list that is open in Excel is left alone: its rows are added
    on the next run. Returns (files written, rows added, files that were busy)."""
    todo = [dict(r) for r in conn.execute(
        "SELECT * FROM docs WHERE doc_type='Fuel Invoice' AND status=? AND (listed_at IS NULL OR listed_at='') ORDER BY number, id",
        (db.FILED,))]
    by_month = {}
    for d in todo:
        when = str(d["created_at"] or db.now())[:10]
        by_month.setdefault((int(when[:4]), int(when[5:7])), []).append(d)
    written, added, busy = [], 0, []
    for (year, month), rows in sorted(by_month.items()):
        path = fuel_list_path(year, month)
        try:
            wb = load_workbook(path) if os.path.exists(path) else _new_fuel_book()
        except Exception:
            busy.append(os.path.basename(path))
            continue
        ws = wb.worksheets[0]
        have = {str(ws.cell(r, 3).value or "").strip().upper() for r in range(2, ws.max_row + 1)}
        new = 0
        for d in rows:
            if (d["number"] or "").upper() in have:
                continue
            r = ws.max_row + 1
            ws.cell(r, 1, _date(str(d["created_at"])[:10])).number_format = "dd/mm/yyyy"
            ws.cell(r, 2, _fuel_contact(d))
            ws.cell(r, 3, d["number"])
            ws.cell(r, 4, d["reference"] or "")
            ws.cell(r, 5, _date(d["doc_date"])).number_format = "dd/mm/yyyy"
            ws.cell(r, 9, "")
            have.add((d["number"] or "").upper())
            new += 1
        ws.auto_filter.ref = f"A1:I{max(ws.max_row, 2)}"
        try:
            if os.path.exists(path):
                os.chmod(path, stat.S_IWRITE | stat.S_IREAD)
            wb.save(path)
        except OSError:
            busy.append(os.path.basename(path))
            continue
        stamp = db.now()
        for d in rows:
            conn.execute("UPDATE docs SET listed_at=? WHERE ref=?", (stamp, d["ref"]))
        conn.commit()
        written.append(path)
        added += new
    return written, added, busy


def _fuel_contact(d):
    """'Highland Energy 1 Pty Ltd - PEARL ENERGY TARAGO', exactly as printed in the invoice's Customer (stored with the document)."""
    return str(d.get("ship_to") or "").strip() or f"{d.get('company_folder') or ''} - {d.get('site') or ''}".strip(" -")


# ----------------------------------------------------------------------------------------- Fleet Card record
FLEET_HEADERS = ["Merchant No", "Report Issue Date", "Total payment", "Site", "Company", "Status", "File name"]
FLEET_WIDTHS = [14, 18, 16, 34, 30, 13, 34]


def fleet_record_path(year, month):
    return os.path.join(reports_dir(), f"FleetCard_Record_{MONTH_ABBR[month - 1]}_{year}.xlsx")


def write_fleet_record(conn, year, month):
    """Reports\\FleetCard_Record_<MON>_<YEAR>.xlsx: every Fleet Card statement the bot downloaded that month (sheet Record)
    and the ones still waiting (sheet Needs attention). Rebuilt from the records each time."""
    start, end = f"{year:04d}-{month:02d}-01", (f"{year + 1:04d}-01-01" if month == 12 else f"{year:04d}-{month + 1:02d}-01")
    rows = [dict(r) for r in conn.execute(
        "SELECT * FROM docs WHERE doc_type='Fleet Card Statement' AND created_at>=? AND created_at<? ORDER BY site, doc_date, id", (start, end))]
    wb = Workbook(write_only=True)
    for title, keep in (("Record", lambda d: True), ("Needs attention", lambda d: d["status"] in (db.UNASSIGNED, db.UNRECOGNISED))):
        ws = wb.create_sheet(title)
        _print_setup(ws)
        for i, w in enumerate(FLEET_WIDTHS, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A2"
        ws.append(_head(ws, FLEET_HEADERS))
        n = 0
        for d in rows:
            if not keep(d):
                continue
            merchant = str(d["ship_to"] or "")[9:] if str(d["ship_to"] or "").startswith("merchant:") else ""
            date = WriteOnlyCell(ws, _date(d["doc_date"]))
            date.number_format = "dd/mm/yyyy"
            amount = WriteOnlyCell(ws, d["amount"])
            amount.number_format = MONEY
            status = WriteOnlyCell(ws, d["status"])
            if d["status"] in STATUS_FILL:
                status.fill = PatternFill("solid", fgColor=STATUS_FILL[d["status"]])
            site = d["site"] if d["site"] and d["site"] != settings.UNASSIGNED else ""
            ws.append([merchant, date, amount, site, d["company_folder"] or "", status, d["dest_name"] or d["orig_name"] or ""])
            n += 1
        ws.auto_filter.ref = f"A1:G{max(n + 1, 2)}"
    return _save(wb, fleet_record_path(year, month))
