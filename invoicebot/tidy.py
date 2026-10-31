"""One-time tidy of the holding folders (5.8.0): files the bot put in Holding\\_Statements, _Order Confirmations and _Unrecognised before the
month folders existed are moved into their month folder, and statements / order confirmations get the new name  supplier_site_date_TAG.

    python run_bot.py --tidy-holding            LIST ONLY: shows what would move, changes nothing (also written to Reports\\Tidy_Holding_Preview.xlsx)
    python run_bot.py --tidy-holding --apply    does the move

Only files the bot filed itself (they have a record) are touched, and only when they are still where the bot put them.
A file is copied, size-checked, and only then the original removed. The record, and with it the Excel reports, follow."""
import datetime
import os
import re

from . import db, doctypes, exports, naming, paths, settings
from .router import PlacementError, Router, UNRECOGNISED_DIR, folder_clean, holding_month

MONTH_FOLDER = re.compile(r"^(January|February|March|April|May|June|July|August|September|October|November|December) \d{4}$")


def _holding_folders():
    names = {UNRECOGNISED_DIR.lower(), "_statements", "_order confirmations"}
    for r in settings.rows("DocTypes"):
        if r["Route"].strip().lower().startswith("hold") and r["Subfolder"].strip():
            names.add(folder_clean(r["Subfolder"].strip()).lower())
    return names


def plan(conn):
    """[{'ref','type','from','to_dir','to_name','missing_site'}] for every record still sitting outside a month folder."""
    from .pipeline import Pipeline
    holding = os.path.abspath(settings.holding_root())
    window = max(0, settings.run_int("Month window (months)", 3))
    date_fmt = settings.run_value("File name date format", naming.DEFAULT_DATE_FORMAT) or naming.DEFAULT_DATE_FORMAT
    top = _holding_folders()
    out = []
    for d in db.docs_by_status(conn, db.PARKED, db.UNRECOGNISED):
        path = d.get("current_path") or ""
        if not path or not os.path.isfile(path):
            continue
        try:
            rel = os.path.relpath(os.path.abspath(path), holding)
        except ValueError:
            continue
        parts = rel.split(os.sep)
        if parts[0].startswith("..") or len(parts) < 2 or parts[0].lower() not in top:
            continue
        if len(parts) >= 3 and MONTH_FOLDER.match(parts[1]):
            continue                                                    # already tidy
        f = Pipeline._fields_from_row(d)
        month = holding_month(f["received"], f["date"], window)
        name, missing = os.path.basename(path), False
        rule = doctypes.rule(d["doc_type"]) if d["doc_type"] else None
        if rule and "{site}" in rule["pattern"]:
            site = f["site"] if f["site"] and f["site"] != settings.UNASSIGNED else ""
            shown = naming.format_date(f["date"] or f["received"].date(), date_fmt)
            name = naming.build(rule["pattern"], f["supplier"], f["short"], f["number"] or "", shown, rule["tag"], rule["type"],
                                f["orig_name"], f["ext"] or os.path.splitext(path)[1], site=site)
            missing = not site
        to_dir = os.path.join(holding, parts[0], month, *parts[1:-1])
        out.append({"ref": d["ref"], "type": d["doc_type"] or "", "from": path, "to_dir": to_dir, "to_name": name, "missing_site": missing,
                    "remarks": d.get("remarks") or ""})
    return out


def run(apply=False, log=print):
    dbfile = paths.p("data", "bot.db")
    if not os.path.exists(dbfile):
        log("There are no records yet: nothing to tidy.")
        return 0, 0
    conn = db.connect(dbfile)
    try:
        items = plan(conn)
        if not apply:
            _preview_file(items)
            for it in items:
                log(f"{it['ref']}  {it['type']:<22} {it['from']}\n      -> {os.path.join(it['to_dir'], it['to_name'])}")
            log(f"\nLIST ONLY: {len(items)} file(s) would be moved. Nothing was changed. To do it: python run_bot.py --tidy-holding --apply")
            return len(items), 0
        router = Router(dry=False, log=log)
        moved = failed = 0
        for it in items:
            try:
                new_path = router.move(it["from"], it["to_dir"], it["to_name"])
            except PlacementError as e:
                failed += 1
                log(f"NOT moved {it['ref']}: {e.reason}")
                continue
            fields = dict(current_path=new_path, dest_dir=os.path.dirname(new_path), dest_name=os.path.basename(new_path))
            if it["missing_site"] and settings.MISSING_SITE not in it["remarks"]:
                fields["remarks"] = "; ".join(x for x in (it["remarks"], settings.MISSING_SITE) if x)
            db.update_doc(conn, it["ref"], **fields)
            conn.commit()
            moved += 1
        log(f"Tidy done: {moved} moved, {failed} not moved. Run the bot (or press Make report) to refresh To_Fix.xlsx and the reports.")
        return moved, failed
    finally:
        conn.close()


def _preview_file(items):
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Would move"
    ws.append(["Ref", "Document type", "Now in", "Would go to", "New file name", "Site missing"])
    for it in items:
        ws.append([it["ref"], it["type"], it["from"], it["to_dir"], it["to_name"], "yes" if it["missing_site"] else ""])
    for col, w in zip("ABCDEF", (10, 22, 80, 70, 50, 12)):
        ws.column_dimensions[col].width = w
    wb.save(os.path.join(exports.reports_dir(), "Tidy_Holding_Preview.xlsx"))
