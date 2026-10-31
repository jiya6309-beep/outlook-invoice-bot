"""The quarterly records archive.

Every three months (calendar quarters: Jan-Mar, Apr-Jun, Jul-Sep, Oct-Dec) the bot writes the records of the quarter that has just ended
to a read-only Excel file  Archive_2026-07_to_2026-09.xlsx  in Reports\\Archive (or the folder in the Run sheet).
NOTHING is deleted from the database: the database is small (about 1 GB a year) and must keep every record so that a re-sent invoice
is still recognised as a duplicate. The archive is the human-readable, read-only copy that can be kept, moved or handed over."""
import datetime
import os
import re

from . import db, exports, paths, settings


def quarter_of(d):
    return d.year, (d.month - 1) // 3 + 1


def bounds(year, quarter):
    """(first day, last day) of a calendar quarter."""
    start = datetime.date(year, 3 * (quarter - 1) + 1, 1)
    nxt = datetime.date(year + (1 if quarter == 4 else 0), 1 if quarter == 4 else 3 * quarter + 1, 1)
    return start, nxt - datetime.timedelta(days=1)


def last_finished(today):
    year, q = quarter_of(today)
    return (year - 1, 4) if q == 1 else (year, q - 1)


def parse(text):
    """'2026-Q3' -> (2026, 3)."""
    m = re.fullmatch(r"\s*(\d{4})\s*-?\s*[Qq]([1-4])\s*", str(text))
    if not m:
        raise ValueError(f"'{text}' is not a quarter. Write it like 2026-Q3.")
    return int(m.group(1)), int(m.group(2))


def folder():
    return settings.run_value("Records archive folder") or os.path.join(settings.reports_root(), "Archive")


def path_for(year, quarter):
    start, end = bounds(year, quarter)
    return os.path.join(folder(), f"Archive_{start:%Y-%m}_to_{end:%Y-%m}.xlsx")


def make(conn, year, quarter, basis="arrival", force=False):
    """Write the archive of one quarter. Returns (path, documents) or (None, 0) if it exists already (and not force) or there is nothing to archive."""
    target = path_for(year, quarter)
    if os.path.exists(target):
        if not force:
            return None, 0
        os.chmod(target, 0o644)
        os.remove(target)
    start, end = bounds(year, quarter)
    source = db.DocSource(conn, *db.between_clause(start, end, basis))
    count = source.count()
    if not count:
        return None, 0
    os.makedirs(os.path.dirname(target), exist_ok=True)
    exports.write_archive(source, db.corrections_between(conn, start, end), target,
                          f"Archive of the records of {start:%B %Y} to {end:%B %Y} (by {'invoice date' if basis == 'invoice' else 'e-mail arrival'})")
    return target, count


def make_due(conn, today, basis="arrival"):
    """The archive of the quarter that ended most recently, if it has not been written yet."""
    year, q = last_finished(today)
    return make(conn, year, q, basis)


# ------------------------------------------------------------------ the new features (same rule as the invoices)
def _feature_path(name, year, quarter):
    start, end = bounds(year, quarter)
    return os.path.join(folder(), f"Archive_{name}_{start:%Y-%m}_to_{end:%Y-%m}.xlsx")


def _write_simple(target, title, headers, rows):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    import stat as _stat
    wb = Workbook()
    ws = wb.active
    ws.title = "Records"
    ws.append([title])
    ws["A1"].font = Font(bold=True, size=13)
    ws.append(headers)
    for c in ws[2]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
    for r in rows:
        ws.append(list(r))
    for i, h in enumerate(headers, start=1):
        ws.column_dimensions[ws.cell(2, i).column_letter].width = max(12, min(60, len(h) + 8))
    ws.freeze_panes = "A3"
    os.makedirs(os.path.dirname(target), exist_ok=True)
    wb.save(target)
    os.chmod(target, _stat.S_IREAD)                      # read-only, like the invoice archive


def make_features(conn, year, quarter):
    """Archive_FuelInvoices_..., Archive_FleetCard_... and Archive_Security_... for a quarter (each only once, nothing deleted).
    Returns [(path, rows)]."""
    start, end = bounds(year, quarter)
    s, e = start.isoformat(), (end + datetime.timedelta(days=1)).isoformat()
    made = []
    fuel = [(str(d["created_at"])[:10], d["ship_to"] or "", d["number"], d["reference"] or "", d["doc_date"], d["company_folder"],
             d["status"], d["dest_name"] or "") for d in conn.execute(
        "SELECT * FROM docs WHERE doc_type='Fuel Invoice' AND created_at>=? AND created_at<? ORDER BY number", (s, e))]
    fleet = [((d["ship_to"] or "")[9:] if str(d["ship_to"] or "").startswith("merchant:") else "", d["doc_date"], d["amount"], d["site"],
              d["company_folder"], d["status"], d["dest_name"] or "") for d in conn.execute(
        "SELECT * FROM docs WHERE doc_type='Fleet Card Statement' AND created_at>=? AND created_at<? ORDER BY site, doc_date", (s, e))]
    try:
        sec = list(conn.execute("SELECT client, at, kind, line, file FROM security_events WHERE at>=? AND at<? ORDER BY client, at", (s, e)))
    except Exception:
        sec = []
    for name, rows, headers in (
            ("FuelInvoices", fuel, ["Download Date", "Contact", "Invoice Number", "Reference", "Invoice Date", "Company folder", "Status", "File name"]),
            ("FleetCard", fleet, ["Merchant No", "Report Issue Date", "Total payment", "Site", "Company", "Status", "File name"]),
            ("Security", sec, ["Client No", "Time", "Event", "Line in the report", "Report file"])):
        target = _feature_path(name, year, quarter)
        if rows and not os.path.exists(target):
            _write_simple(target, f"{name} - {start:%B %Y} to {end:%B %Y} (read-only archive)", headers, rows)
            made.append((target, len(rows)))
    return made
