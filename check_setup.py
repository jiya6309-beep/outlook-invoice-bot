"""Checks that this PC and the settings are ready.   python check_setup.py   (or check_setup.bat)

OK = fine,  WARN = works but look at it,  FAIL = must be fixed before a live run."""
import importlib
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

results = []


def report(level, text, hint=""):
    results.append(level)
    print(f"  [{level:4}] {text}" + (f"\n         -> {hint}" if hint else ""))


def main():
    from invoicebot import __version__, companies, db, memory, ocr, paths, settings, sites

    print(f"\nOutlook Invoice Bot {__version__} - setup check\nProgram folder: {paths.home()}\n")
    try:
        settings.get()
    except settings.SettingsError as e:
        print("  [FAIL]", e, "\n\nResult: NOT READY.")
        return 1

    print("1. Python and libraries")
    v = sys.version_info
    report("OK" if v >= (3, 9) else "FAIL", f"Python {v.major}.{v.minor}.{v.micro}", "" if v >= (3, 9) else "Install Python 3.10 or newer")
    for module, need, why in [("openpyxl", True, "Excel files"), ("pdfplumber", True, "reading PDFs"), ("PIL", True, "images (Pillow)"),
                              ("pytesseract", False, "OCR for scanned PDFs / logos"), ("win32com", sys.platform.startswith("win"), "talking to Outlook (pywin32)")]:
        try:
            importlib.import_module(module)
            report("OK", f"{module} ({why})")
        except ImportError:
            report("FAIL" if need else "WARN", f"{module} is missing ({why})", "double-click Start_Bot.bat" if need else "optional / Windows only")

    print("\n2. OCR (optional)")
    settings.reset()
    ocr.configure(settings.run_bool("OCR enabled", True), settings.run_value("Tesseract path"), settings.run_int("OCR max pages", 3))
    report("OK" if ocr.available() else "WARN", "Tesseract OCR ready" if ocr.available() else "OCR not available: " + ocr.why_unavailable(),
           "" if ocr.available() else "Scanned PDFs will go to _Unrecognised. See README 'Optional: OCR'.")

    print("\n3. Settings (Bot_Settings.xlsx)")
    path = settings.settings_file()
    report("OK" if os.path.exists(path) else "WARN", f"settings file: {path}", "" if os.path.exists(path) else "it will be created on the first run")
    rows = companies.all_active()
    report("OK" if len(rows) == 13 else "WARN", f"{len(rows)} active companies (expected 13)")
    n_sites = len(sites.active_sites())
    report("OK" if n_sites else "WARN", f"{n_sites} sites in the Sites sheet", "" if n_sites else "run import_sites.py to load your site list")

    print("\n4. Folders")
    mode, root = settings.destination_mode(), settings.destination_root()
    report("OK", f"destination mode: {mode}")
    if mode == "live" and not root and settings.run_bool("Find the destination folder automatically", True):
        from invoicebot import autofind
        print("         (looking for the synced folder on this PC ...)")
        found, hits = autofind.discover()
        root = found
        report("OK" if found else "FAIL", f"destination folder found automatically: {found}" if found else "the synced destination folder was not found automatically",
               "" if len(hits) < 2 else "several matches: " + " | ".join(hits[:3]))
    if mode == "live" and not root:
        report("FAIL", "live mode but no SharePoint path set",
               "fill in 'Live destination root' (Bot_Settings.xlsx, sheet Run) or LIVE_DESTINATION_ROOT in config.py")
    else:
        report("OK" if os.path.isdir(root) or mode == "test" else "FAIL", f"destination root: {root}",
               "" if os.path.isdir(root) or mode == "test" else "the folder does not exist - is OneDrive synced and signed in?")
        if mode == "live" and os.path.isdir(root):
            from invoicebot.router import Router
            import datetime
            r = Router()
            month = r.month_dir(datetime.date.today())
            report("OK" if os.path.isdir(month) else "WARN", f"this month's folder: {month}",
                   "" if os.path.isdir(month) else "missing: files will wait in _Unassigned until it exists")
            missing = [c["Folder name"] for c in rows if not os.path.isdir(os.path.join(month, c["Folder name"]))]
            if os.path.isdir(month):
                report("OK" if not missing else "WARN", "company folders present" if not missing else "missing company folders: " + ", ".join(missing))
    pattern = settings.run_value("Destination pattern", "{MONTH_FOLDER}\\{COMPANY}")
    report("OK", f"folder layout: <root>\\{pattern}")
    if "{SITE}" in pattern and root and os.path.isdir(root):
        names = sites.active_sites()
        have = [s["name"] for s in names if os.path.isdir(os.path.join(root, sites.folder_for(s["name"]))) ] if pattern.strip() == "{SITE}" else []
        if pattern.strip() == "{SITE}":
            report("OK" if len(have) == len(names) else "WARN", f"site folders found: {len(have)} of {len(names)} sites",
                   "" if len(have) == len(names) else "press 'Match my sites to the folders' in the panel (Setup 2) to fill in the real folder names")
    for what, folder in (("holding", settings.holding_root()), ("reports", settings.reports_root())):
        try:
            settings.check_output_root(folder, what.title())
            os.makedirs(folder, exist_ok=True)
            probe = os.path.join(folder, ".probe")
            with open(probe, "w") as f:
                f.write("x")
            os.remove(probe)
            report("OK", f"{what} folder is writable: {folder}")
        except settings.SettingsError as e:
            report("FAIL", f"{what} folder not found: {folder}", "check that OneDrive / SharePoint has synced it, or correct the path (sheet Run)")
        except OSError as e:
            report("FAIL", f"{what} folder is not writable: {folder}", str(e))
    for sp in companies.special_rows():
        if sp["role"] in ("fleet card", "statements"):
            ok = bool(sp["dest"]) and os.path.isdir(sp["dest"])
            report("OK" if ok else "WARN", f"{sp['label']} ({sp['role']}): {sp['dest'] or '(no Destination folder)'}",
                   "" if ok else "check the SharePoint sync, or the Destination folder in the SpecialMailboxes sheet")
        else:
            report("OK", f"{sp['label']}: {sp['role'].upper()} - not read")

    print("\n5. Database and PC")
    try:
        conn = db.connect(paths.p("data", "bot.db"))
        n = conn.execute("SELECT COUNT(*) FROM docs").fetchone()[0]
        conn.close()
        report("OK", f"database opens ({n} documents recorded)")
    except Exception as e:
        report("FAIL", "database cannot be opened", str(e))
    free = memory.free_ram_mb()
    report("OK" if (free or 9999) >= 1500 else "WARN", f"free memory: {free} MB, CPU cores: {memory.cpu_count()}",
           "" if (free or 9999) >= 1500 else "low memory: the bot will read one document at a time")

    print("\n6. Outlook")
    if not sys.platform.startswith("win"):
        report("WARN", "not Windows: Outlook cannot be checked here")
    else:
        try:
            from invoicebot.outlook_io import ClassicOutlook
            gw = ClassicOutlook()
            names = gw.list_mailboxes()
            report("OK", f"connected to Outlook; {len(names)} mailbox(es) open")
            missing = []
            for row in companies.reading_list():
                try:
                    gw.resolve_store(companies.mailbox_names(row))
                except LookupError:
                    missing.append(f"{row['Company name']} ({companies.mailbox_names(row)[0]})")
            report("OK" if not missing else "WARN", "all company mailboxes found" if not missing else "not found in Outlook: " + "; ".join(missing),
                   "" if not missing else "open the mailbox in Outlook, or correct 'Mailbox (as in Outlook)' in the Companies sheet. Names in Outlook: " + ", ".join(names))
        except Exception as e:
            report("FAIL", f"cannot connect to Outlook: {e}", "classic desktop Outlook must be installed, OPEN and signed in (not 'New Outlook')")

    fails, warns = results.count("FAIL"), results.count("WARN")
    print(f"\nResult: {fails} to fix, {warns} to look at.  " + ("READY." if not fails else "NOT READY."))
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
