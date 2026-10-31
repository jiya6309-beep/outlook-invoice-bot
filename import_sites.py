"""Loads the company's site list (Excel) into the Sites sheet of Bot_Settings.xlsx.

    python import_sites.py NEW_LIST.xlsx
    python import_sites.py NEW_LIST.xlsx --old OLD_LIST.xlsx
    python import_sites.py LIST.xlsx --new-sheet 0 --old-sheet 1      (new and old lists are two sheets of one file)

The NEW list is the master. The OLD list adds the former names of renamed sites as aliases and keeps sites that
only exist there. Handler e-mails and aliases already typed in the Sites sheet are kept. A backup of
Bot_Settings.xlsx is made first."""
import argparse
import datetime
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def main(argv=None):
    ap = argparse.ArgumentParser(description="Import the site list into Bot_Settings.xlsx")
    ap.add_argument("new", help="the current site list (.xlsx)")
    ap.add_argument("--old", help="the older site list (.xlsx) - optional")
    ap.add_argument("--new-sheet", default="0", help="sheet position (0 = first) or name in the new file")
    ap.add_argument("--old-sheet", default="0", help="sheet position or name in the old file")
    args = ap.parse_args(argv)

    from openpyxl import load_workbook
    from invoicebot import settings, sites

    pick = lambda s: int(s) if str(s).isdigit() else s
    new = sites.read_site_list(args.new, pick(args.new_sheet))
    old = sites.read_site_list(args.old, pick(args.old_sheet)) if args.old else None
    if not new:
        print("No sites were found. The sheet needs rows with 'SITE NAME' and 'ADDRESS' headings under 'Company Name : ...' lines.")
        return 1
    rows, report = sites.build_rows(new, old)

    # keep what people already typed: handler e-mails and extra aliases
    settings.reset()
    existing = {r["Site name"]: r for r in settings.rows("Sites")}
    for row in rows:
        old_row = existing.get(row[0])
        if old_row:
            if old_row.get("Handler email"):
                row[5] = old_row["Handler email"]
            merged = sorted({a for a in (row[2].split("|") + old_row.get("Aliases", "").split("|")) if a})
            row[2] = "|".join(merged)
            if not row[4]:
                row[4] = old_row.get("Handler", "")
            if old_row.get("SharePoint folder"):                      # a folder name already matched is never lost on re-import
                row.extend([""] * (10 - len(row)))
                row[9] = old_row["SharePoint folder"]

    path = settings.settings_file()
    if os.path.exists(path):
        backup = path.replace(".xlsx", f"_backup_{datetime.datetime.now():%Y%m%d_%H%M%S}.xlsx")
        shutil.copyfile(path, backup)
        print("Backup of the settings:", backup)
    else:
        settings.create_default_file(path)
    wb = load_workbook(path)
    ws = wb["Sites"]
    ws.delete_rows(2, ws.max_row)
    for row in rows:
        ws.append(row)
    wb.save(path)
    print(f"{len(new)} sites in the new list" + (f", {len(old)} in the old list" if old else "") + f"  ->  {len(rows)} rows written to the Sites sheet.")
    aliased = [r for r in rows if r[2]]
    print(f"{len(aliased)} sites have extra names (renamed sites / spelling variants).")
    if report:
        print("\nPlease look at:")
        for line in report:
            print("  -", line)
    return 0


if __name__ == "__main__":
    sys.exit(main())
