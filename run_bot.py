r"""Outlook Invoice Bot - main program.

    python run_bot.py                      one live run (automatic catch-up since the last successful run)
    python run_bot.py --shadow             a rehearsal: shows what WOULD happen, changes nothing
    python run_bot.py --from 2026-10-01 --to 2026-10-05      read e-mails that ARRIVED in this period (inclusive)
    python run_bot.py --mailbox "One@pearlenergy.com.au"      only one mailbox
    python run_bot.py --demo               try it with sample e-mails (no Outlook needed)
    python run_bot.py --msg-folder C:\Msgs  run the bot on a folder of saved .msg e-mails (no Outlook needed)
    python run_bot.py --report --from 2026-10-01 --to 2026-10-31      Excel report of the documents handled in a period
    python run_bot.py --ac --from 2026-10-01 --to 2026-10-31          A/C Excel (invoices, credit notes, dockets) for a period
    python run_bot.py --match-folders      match the site list to the real SharePoint site folders (fills Sites > SharePoint folder)
    python run_bot.py --check              check that this PC is ready

See README.md.
"""
import argparse
import datetime
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def _date(text):
    try:
        return datetime.date.fromisoformat(text)
    except ValueError:
        raise argparse.ArgumentTypeError(f"'{text}' is not a date - use YYYY-MM-DD, e.g. 2026-10-31")


def match_folders_command(folder):
    from invoicebot import settings, sites
    root = (folder or "").strip().strip('"') or settings.destination_root()
    if not root or not os.path.isdir(root):
        print(f"The destination folder was not found: '{root}'.\nSet the real SharePoint folder first (panel: Setup 2, or Bot_Settings.xlsx > Run > Live destination root) "
              "and make sure OneDrive has synced it.")
        return 2
    result = sites.match_folders(root)
    written = sites.apply_folder_matches(result, root)
    total = len(result["matched"]) + len(result["fuzzy"]) + len(result["missing"]) + len(result["ambiguous"])
    print(f"Folder: {root}\nLooked in: {result.get('looked_in', root)}\n{len(result['matched']) + len(result['fuzzy'])} of {total} sites matched to a folder "
          f"({len(result['matched'])} exactly, {len(result['fuzzy'])} by similar spelling). {written} folder name(s) written to the Sites sheet.\n")
    def show(title, items, limit=70):
        if items:
            print(f"{title} ({len(items)}):")
            for item in items[:limit]:
                print("   -", item)
            if len(items) > limit:
                print(f"   ... and {len(items) - limit} more")
            print()
    show("Matched by similar spelling - please check", [f"{s}  ->  {f}   (similarity {r})" for s, (f, r) in result["fuzzy"].items()])
    show("Sites WITHOUT a folder: documents for these wait in _Unassigned until the folder exists or you correct the name", result["missing"])
    show("Folders WITHOUT a site in the list: add the site to the Sites sheet, or an alias to an existing site", result["extra"])
    show("Several folders fit one site - choose by hand in the SharePoint folder column", [f"{s}: {', '.join(fs)}" for s, fs in result["ambiguous"].items()])
    return 0


def settings_error():
    from invoicebot import settings as _s
    return _s.SettingsError


def main(argv=None):
    ap = argparse.ArgumentParser(description="Outlook Invoice Bot", formatter_class=argparse.RawDescriptionHelpFormatter,
                                 epilog="Settings are in Bot_Settings.xlsx. Close it before a run.")
    ap.add_argument("--shadow", action="store_true", help="rehearsal: change nothing, write a report of what would happen")
    ap.add_argument("--from", dest="from_date", type=_date, metavar="YYYY-MM-DD", help="first day (e-mail arrival date)")
    ap.add_argument("--to", dest="to_date", type=_date, metavar="YYYY-MM-DD", help="last day, inclusive")
    ap.add_argument("--mailbox", help="only this mailbox (part of its name is enough)")
    ap.add_argument("--workers", type=int, help="reader processes (default: automatic)")
    ap.add_argument("--subject", metavar="TEXT", help="TESTING: only e-mails whose subject contains TEXT are considered, all others are left alone")
    ap.add_argument("--demo", action="store_true", help="run with sample e-mails in a simulated Outlook")
    ap.add_argument("--msg-folder", metavar="FOLDER", help="test with a folder of saved .msg e-mails (simulated Outlook, own working folder)")
    ap.add_argument("--msg-mailbox", metavar="NAME", help="with --msg-folder: company mailbox for e-mails whose recipient is not recognised")
    ap.add_argument("--report", action="store_true", help="make the Excel report for --from/--to from the records (reads no e-mail)")
    ap.add_argument("--ac", action="store_true", help="make the A/C Excel (filed invoices, credit notes, dockets) for --from/--to")
    ap.add_argument("--match-folders", nargs="?", const="", metavar="FOLDER",
                    help="match the Sites sheet to the real site folders under FOLDER (default: the destination root) and write the folder names")
    ap.add_argument("--archive-quarter", nargs="?", const="", metavar="YYYY-Qn",
                    help="write the read-only records archive of a quarter (default: the last finished one). Example: 2026-Q3")
    ap.add_argument("--test-outlook-marking", action="store_true",
                    help="try every kind of Outlook marking (read, tick, flag, category) on ONE test e-mail and report what works (use with --mailbox and --subject)")
    ap.add_argument("--tidy-holding", action="store_true",
                    help="one-time tidy: move old files of _Statements, _Order Confirmations and _Unrecognised into month folders (LIST ONLY unless --apply)")
    ap.add_argument("--apply", action="store_true", help="with --tidy-holding: really move the files")
    ap.add_argument("--security", action="store_true",
                    help="SECURITY RUN: read the alarm reports of the Security mailbox and fill the fortnightly STORE OPEN & CLOSE Excel "
                         "(with --shadow: rehearsal; with --from/--to: a chosen period). Never reads or files invoices.")
    ap.add_argument("--import-security-sites", metavar="XLSX",
                    help="build Security_Settings.xlsx from a filled STORE OPEN & CLOSE workbook (the sheet with Client No)")
    ap.add_argument("--check", action="store_true", help="check this PC and the settings")
    args = ap.parse_args(argv)

    if args.msg_folder:
        args.msg_folder = args.msg_folder.strip('"').strip()
    if args.from_date and args.to_date and args.from_date > args.to_date:
        ap.error("--from is later than --to")

    if args.check:
        import check_setup
        return check_setup.main()

    if args.tidy_holding:
        from invoicebot import paths, runlock, tidy
        lock = None
        if args.apply:
            lock = runlock.RunLock(paths.p("data", "run.lock"))
            try:
                lock.__enter__()
            except runlock.LockBusy as e:
                print(f"Not started: {e}")
                return 2
        try:
            tidy.run(apply=args.apply)
        finally:
            if lock and lock.held:
                lock.__exit__(None, None, None)
        return 0

    if args.import_security_sites:
        from invoicebot import security_settings as _ss
        if os.path.exists(_ss.path()):
            print(f"{_ss.FILE_NAME} already exists. Rename it first if you really want to rebuild it from the workbook.")
            return 2
        target, n = _ss.import_from_workbook(args.import_security_sites.strip('"'))
        print(f"{n} sites written to {target}. Check the opening hours, then save.")
        return 0

    if args.security:
        from invoicebot import security, security_settings as _ss
        try:
            if args.demo:
                raise _ss.SecurityError("--security cannot be combined with --demo")
            from invoicebot.outlook_io import ClassicOutlook
            gateway = ClassicOutlook()
        except _ss.SecurityError as e:
            print(e)
            return 2
        except Exception as e:
            print(f"Outlook could not be opened: {e}\nClassic Outlook must be open and signed in on this PC.")
            return 2
        try:
            r = security.SecurityRun(gateway, shadow=args.shadow, from_date=args.from_date, to_date=args.to_date).run()
        except (_ss.SecurityError, settings_error()) as e:
            print(f"Security run stopped: {e}")
            return 2
        text = r.summary()
        print(text)
        try:
            from invoicebot import paths as _p
            os.makedirs(_p.p("logs"), exist_ok=True)
            import datetime as _d
            with open(_p.p("logs", f"security_{_d.datetime.now():%Y%m%d_%H%M%S}.txt"), "w", encoding="utf-8") as fh:
                fh.write(text)
        except OSError:
            pass
        return 1 if r.warnings else 0

    if args.test_outlook_marking:
        if not args.mailbox or not args.subject:
            print("Tell me which mailbox and which e-mail: use --mailbox jiya --subject BOTTEST  (the panel boxes 'Only one mailbox' and 'Only touch e-mails whose subject contains').")
            return 2
        try:
            from invoicebot.outlook_io import ClassicOutlook
            gateway = ClassicOutlook()
        except Exception as e:
            print(f"Outlook could not be opened: {e}\nClassic Outlook must be open and signed in on this PC ('New Outlook' cannot be used).")
            return 2
        lines, ok = gateway.diagnose_marking(args.mailbox, args.subject)
        print("\n".join(lines))
        return 0 if ok else 1

    if args.archive_quarter is not None:
        from invoicebot import db, paths, quarterly, settings
        import datetime as _dt
        try:
            year, q = quarterly.parse(args.archive_quarter) if args.archive_quarter else quarterly.last_finished(_dt.date.today())
        except ValueError as e:
            ap.error(str(e))
        dbfile = paths.p("data", "bot.db")
        if not os.path.exists(dbfile):
            print("There are no records yet: the bot has not processed any document.")
            return 1
        conn = db.connect(dbfile)
        try:
            basis = "invoice" if settings.run_value("Month basis", "arrival").lower() == "invoice" else "arrival"
            path, n = quarterly.make(conn, year, q, basis, force=bool(args.archive_quarter))
        finally:
            conn.close()
        print(f"Archive of {year}-Q{q}: " + (f"{n} documents written to {path}" if path else "nothing written (it exists already, or the quarter has no documents). Name the quarter to write it again."))
        return 0

    if args.match_folders is not None:
        return match_folders_command(args.match_folders)

    if args.ac:
        if not (args.from_date and args.to_date):
            ap.error("--ac needs --from and --to")
        from invoicebot import db, exports, paths
        dbfile = paths.p("data", "bot.db")
        if not os.path.exists(dbfile):
            print("There are no records yet: the bot has not processed any document.")
            return 1
        conn = db.connect(dbfile)
        try:
            file, n = exports.write_ac(conn, args.from_date, args.to_date)
        finally:
            conn.close()
        print(f"{n} documents (invoices, credit notes, dockets) from {args.from_date} to {args.to_date}\nA/C Excel saved: {file}")
        return 0

    if args.report:
        if not (args.from_date and args.to_date):
            ap.error("--report needs --from and --to")
        from invoicebot import db, exports, paths, settings
        dbfile = paths.p("data", "bot.db")
        if not os.path.exists(dbfile):
            print("There are no records yet: the bot has not processed any document.")
            return 1
        conn = db.connect(dbfile)
        try:
            file, n = exports.write_range(conn, args.from_date, args.to_date, settings.run_value("Report date basis", "invoice"))
        finally:
            conn.close()
        print(f"{n} documents from {args.from_date} to {args.to_date}\nReport saved: {file}")
        return 0

    from invoicebot import demo, paths
    if args.msg_folder and not os.path.isdir(args.msg_folder):
        print(f"The folder '{args.msg_folder}' does not exist.")
        return 2
    if args.demo or args.msg_folder:
        # keep the real Bot_Settings.xlsx (companies, keywords, sites) but work in a separate folder
        import shutil
        real_settings = os.path.join(paths.project_dir(), "Bot_Settings.xlsx") if not os.environ.get("BOT_HOME") else None
        demo.reset_workspace("demo_run" if args.demo else "msg_test_run")
        if args.msg_folder and real_settings and os.path.exists(real_settings):
            shutil.copyfile(real_settings, os.path.join(paths.home(), "Bot_Settings.xlsx"))
            from invoicebot import settings as _settings          # the test never writes into SharePoint: every output folder -> the test folder
            _settings.reset()

            def _local(wb):
                for key in ("Holding root", "Reports root", "Records archive folder", "Backup folder", "Archive folder", "Live destination root"):
                    _settings.set_run_value(wb, key, "")
                _settings.set_run_value(wb, "Destination mode", "test")
                if "SpecialMailboxes" in wb.sheetnames:
                    ws = wb["SpecialMailboxes"]
                    for r in range(2, ws.max_row + 1):
                        if ws.cell(r, 3).value:
                            ws.cell(r, 3).value = os.path.join(paths.home(), "TestOutput", "_" + str(ws.cell(r, 2).value or "special").upper())
                            os.makedirs(ws.cell(r, 3).value, exist_ok=True)
            _settings.edit(_local)
            _settings.reset()
    from invoicebot.pipeline import Pipeline

    if args.demo:
        gateway = demo.build()
        print("DEMO - sample e-mails in a simulated Outlook. Working folder:", paths.home())
    elif args.msg_folder:
        from invoicebot import msgtest
        gateway, first, last, n, problems = msgtest.build(args.msg_folder, args.msg_mailbox)
        print(f"{n} .msg e-mails loaded from {args.msg_folder} (simulated Outlook; your real Outlook is not touched).")
        for name, why in problems:
            print(f"  could not read {name}: {why}")
        if not n:
            print("No readable .msg files found.")
            return 1
        args.from_date, args.to_date = args.from_date or first, args.to_date or last
        print(f"Reading e-mails that arrived {args.from_date} to {args.to_date}. Working folder: {paths.home()}")
    else:
        try:
            from invoicebot.outlook_io import ClassicOutlook
            gateway = ClassicOutlook()
        except ImportError:
            print("The Python package 'pywin32' is missing. Double-click Start_Bot.bat (Windows only).")
            return 2
        except Exception as e:
            print("Could not connect to Outlook:", e)
            print("Classic desktop Outlook must be installed, OPEN and signed in on this PC (not 'New Outlook').")
            return 2

    run = Pipeline(gateway, shadow=args.shadow, from_date=args.from_date, to_date=args.to_date, only=args.mailbox,
                   workers=args.workers, subject_filter=args.subject)
    summary = run.run()

    if args.demo or args.msg_folder:
        print("\nWHAT HAPPENED TO EACH " + ("SAMPLE " if args.demo else "") + "E-MAIL" + ("   (simulated: the real e-mails are unchanged)" if args.msg_folder else ""))
        print("-" * 118)
        for subject, state, flag, cats, expect in demo.describe(gateway):
            print(f"{subject:52} {state:7} {flag:9} {cats:34} | {expect[:60]}")
        print("\nFiles were written under:", paths.home())
        print("  TestOutput\\<MONTH> <YEAR> EMAIL INVOICES\\<COMPANY FOLDER>   (the filed documents)")
        print("  Holding\\_Unassigned and Holding\\_Unrecognised\\...              (parked / unreadable)")
        print("  Reports\\                                                      (Excel files for people)")
    return 0 if not summary.stop_reason or summary.stop_reason.startswith("time limit") else 1


if __name__ == "__main__":
    sys.exit(main())
