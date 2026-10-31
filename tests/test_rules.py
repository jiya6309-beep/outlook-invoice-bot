"""Rule tests: keywords, blocked words, naming, companies, month folders, windows, locks, database, router."""
import datetime
import os
import subprocess
import sys
import time
import unittest

from tests.helpers import Isolated, ONE, read_bytes, write

from invoicebot import companies, db, doctypes, naming, settings, sites
from invoicebot.pipeline import compute_window
from invoicebot.router import MONTHS, ConfigError, PlacementError, Router, month_folder_name
from invoicebot.runlock import LockBusy, RunLock

D = datetime.date


class DocTypeTests(Isolated, unittest.TestCase):
    def test_subject_keywords(self):
        for subject, expected in [("Tax Invoice 778899", "Invoice"), ("INV:8024871400", "Invoice"), ("Credit Note CN-1", "Credit note"),
                                  ("POD D-1", "POD"), ("Weekly Statement", "Statement"), ("Account Summary for 0001", "Account Summary"),
                                  ("Hire Statement Sept", "Hire Statement"), ("DT:06.09.2026 C:1202316 INV:8024871400", "Invoice")]:
            with self.subTest(subject=subject):
                self.assertEqual(doctypes.best_match(subject)[0], expected)
        self.assertFalse(doctypes.subject_matches("Lunch on Friday"))
        self.assertFalse(doctypes.subject_matches("Winning the argument"))     # 'inv' / 'cn' must be whole words only

    def test_blocked_words(self):
        for subject in ["PAYMENT REMINDER", "Overdue Dairy and Drinks Account", "07254094 **Stop Supply Warning**",
                        "URGENT: invoice", "Your remainder notice"]:
            with self.subTest(subject=subject):
                self.assertIsNotNone(doctypes.blocked_word(subject))
        self.assertIsNone(doctypes.blocked_word("Tax Invoice 1234"))
        self.assertIsNone(doctypes.blocked_word("Warningham Road delivery"))   # whole words only

    def test_blocked_beats_keyword_in_the_pipeline_order(self):
        # a subject with both a document keyword and a blocked word is blocked (the pipeline checks blocked first)
        self.assertTrue(doctypes.subject_matches("Service Due Reminder"))
        self.assertEqual(doctypes.blocked_word("Service Due Reminder"), "reminder")

    def test_exception_sender_is_matched_ignoring_spaces(self):
        row = doctypes.sender_row("Rarity Wholesale", "accounts@raritywholesale.com.au")
        self.assertIsNotNone(row)
        self.assertTrue(doctypes.always_download(row))
        hint_only = doctypes.sender_row("", "aus.coke.credit@ccamatil.com")
        self.assertIsNotNone(hint_only)
        self.assertFalse(doctypes.always_download(hint_only))

    def test_title_decides_before_subject(self):
        typ, how = doctypes.decide(subject="Invoice 5", header="WEEKLY STATEMENT\nAcme")
        self.assertEqual(typ, "Statement")
        self.assertIn("title", how)

    def test_rules_loaded(self):
        self.assertEqual(doctypes.rule("Credit note")["sign"], "-")
        self.assertEqual((doctypes.rule("Statement")["route"], doctypes.rule("Statement")["subfolder"]), ("holding", "_Statements"))
        self.assertEqual(doctypes.rule("Account Summary")["route"], "unrecognised")
        self.assertEqual(doctypes.rule("Account Summary")["subfolder"], "Account Summary")
        self.assertFalse(doctypes.rule("POD")["read_amounts"])


class NamingTests(unittest.TestCase):
    def test_patterns(self):
        n = lambda p, **k: naming.build(p, **{"supplier": "Acme Foods Pty Ltd", "short": "Acme", "number": "INV-1", "date": "2026-10-02", **k})
        self.assertEqual(n("{supplier}_{number}", ext=".pdf"), "Acme Foods Pty Ltd_INV-1.pdf")
        self.assertEqual(n("{supplier}_{number}_CN", ext=".PDF"), "Acme Foods Pty Ltd_INV-1_CN.pdf")
        self.assertEqual(n("{supplier}_{date}_STM", ext=".pdf"), "Acme Foods Pty Ltd_2026-10-02_STM.pdf")

    def test_illegal_characters_and_fallbacks(self):
        self.assertEqual(naming.build("{supplier}_{number}", supplier="A/B*C", number='12"3:4', ext=".pdf"), "A-B-C_12-3-4.pdf")
        name = naming.build("{supplier}_{number}", supplier="", number="", original="scan 001.pdf", ext=".pdf")
        self.assertEqual(name, "Unknown supplier_scan 001.pdf")
        self.assertLessEqual(len(naming.build("{supplier}_{number}", supplier="X" * 400, number="1", ext=".pdf")), 120)

    def test_dates_in_names_are_dd_mm_yyyy(self):
        import datetime
        d = datetime.date(2026, 10, 4)
        self.assertEqual(naming.format_date(d), "04-10-2026")
        self.assertEqual(naming.format_date(datetime.datetime(2026, 1, 9, 14, 30)), "09-01-2026")
        self.assertEqual(naming.format_date(d, "yyyymmdd"), "20261004")
        self.assertEqual(naming.format_date(d, "DD.MM.YY"), "04.10.26")
        self.assertEqual(naming.format_date(d, ""), "04-10-2026")                             # empty cell = the default
        self.assertEqual(naming.build("{supplier}_{date}_STM", supplier="Acme", date=naming.format_date(d, "dd/mm/yyyy"), ext=".pdf"),
                         "Acme_04-10-2026_STM.pdf")                                           # a slash can never reach the file system

    def test_folder_names_are_used_exactly_as_written(self):
        for name in ("P & O INVESTMENT", "TWEED HEADS ENERGY NO.1", "CENTRAL COAST ENERGY 1", "OCTOBER 2026 EMAIL INVOICES", "_Statements"):
            self.assertEqual(naming.clean_folder(name), name)
        self.assertEqual(naming.clean_folder('A/B:C*  D. '), "A-B-C-  D")              # only what Windows forbids is replaced; inner spaces are kept
        self.assertEqual(naming.clean_folder("AMPOL  KELSO"), "AMPOL  KELSO")            # the real folder has two spaces
        self.assertEqual(naming.build("{supplier}_{number}", supplier="Suntory Beverage & Food Australia", number="1", ext=".pdf"),
                         "Suntory Beverage & Food Australia_1.pdf")                    # '&' is legal in a file name too

    def test_unique_name(self):
        taken = {"a.pdf", "a_2.pdf"}
        self.assertEqual(naming.unique_name("a.pdf", lambda c: c in taken), "a_3.pdf")


class CompanyTests(Isolated, unittest.TestCase):
    def test_mailbox_lookup_is_lenient(self):
        self.assertEqual(companies.for_mailbox(ONE)["Folder name"], "ONE PEARL ENERGY")
        self.assertEqual(companies.for_mailbox("one pearl energy")["Folder name"], "ONE PEARL ENERGY")
        self.assertEqual(companies.for_mailbox("Pearl Energy Colyton")["Folder name"], "COLYTON FUELS")
        self.assertIsNone(companies.for_mailbox("Somebody Else"))

    def test_company_named_on_document(self):
        hit = companies.find_in_text("TAX INVOICE\nInvoice To: RADHE KRISHNA INVESTMENT PTY LTD\n1 Main St")
        self.assertEqual(hit[0][0]["Folder name"], "RADHE KRISHNA INVESTMENT")
        hit = companies.find_in_text("PEARL ENERGY NO.5 PTY LTD")
        self.assertEqual(hit[0][0]["Folder name"], "FIVE PEARL ENERGY")

    def test_a_test_mailbox_row_with_no_text_to_look_for(self):
        self.add_rows("Companies", [["TEST - JIYA", "", "TEST JIYA", "jiya@pearlenergy.com.au", "Yes"]])
        self.assertEqual(companies.by_name("TEST - JIYA")["Folder name"], "TEST JIYA")
        self.assertEqual(companies.for_mailbox("jiya@pearlenergy.com.au")["Folder name"], "TEST JIYA")
        self.assertIsNone(companies.by_name("Some Unknown Company"))              # must NOT fall onto the empty row
        self.assertIsNone(companies.by_name(""))
        row, src, _ = companies.resolve("", "Colyton Fuels Pty Ltd", "TEST - JIYA")
        self.assertEqual((row["Folder name"], src), ("COLYTON FUELS", "site"))
        row, src, _ = companies.resolve("", "", "TEST - JIYA")
        self.assertEqual((row["Folder name"], src), ("TEST JIYA", "mailbox"))

    def test_the_company_an_email_came_from(self):
        for sender, folder in [("P & O Investment", "P & O INVESTMENT"), ("Highland Energy 1", "HIGHLAND ENERGY 1"),
                               ("Pearl Energy Colyton", "COLYTON FUELS"), ("Shri Hari Investment", "SHRI HARI INVESTMENT"),
                               ("One Pearl Energy", "ONE PEARL ENERGY"), ("Gopinath Investment", "GOPINATH INVESTMENT"),
                               ("Tweed Heads Energy No.1", "TWEED HEADS ENERGY NO.1")]:
            self.assertEqual(companies.for_sender(sender)["Folder name"], folder, sender)
        self.assertEqual(companies.for_sender("", "pearl.colyton@pearlenergy.com.au")["Folder name"], "COLYTON FUELS")
        for generic in ("One", "Energy", "Investment", "Accounts", "UNILEVER AR", ""):               # short or unrelated senders never match
            self.assertIsNone(companies.for_sender(generic), generic)

    def test_sender_in_the_cascade(self):
        mb = "PEARL ENERGY NO: 1 PTY LTD"
        row, src, _ = companies.resolve("", "", mb, sender_company="COLYTON FUELS PTY LTD")
        self.assertEqual((row["Folder name"], src), ("COLYTON FUELS", "sender"))                    # beats the mailbox
        row, src, _ = companies.resolve("RADHE KRISHNA INVESTMENT PTY LTD", "", mb, sender_company="COLYTON FUELS PTY LTD")
        self.assertEqual((row["Folder name"], src), ("RADHE KRISHNA INVESTMENT", "document"))       # the document beats the sender
        row, src, _ = companies.resolve("", "Pearl Energy No 2 Pty Ltd", mb, sender_company="COLYTON FUELS PTY LTD")
        self.assertEqual((row["Folder name"], src), ("TWO PEARL ENERGY", "site"))
        self.set_run(Company_source_order="document,site,mailbox")                                    # a workbook from before the sender existed
        row, src, _ = companies.resolve("", "", mb, sender_company="COLYTON FUELS PTY LTD")
        self.assertEqual(src, "sender")
        self.set_run(Use_sender_as_company_clue="No")
        row, src, _ = companies.resolve("", "", mb, sender_company="COLYTON FUELS PTY LTD")
        self.assertEqual((row["Folder name"], src), ("ONE PEARL ENERGY", "mailbox"))

    def test_cascade_order(self):
        row, src, warns = companies.resolve("COLYTON FUELS PTY LTD", "", "PEARL ENERGY NO: 1 PTY LTD")
        self.assertEqual((row["Folder name"], src), ("COLYTON FUELS", "document"))
        self.assertTrue(warns)                                               # the mailbox disagrees: reported
        row, src, _ = companies.resolve("COLYTON FUELS PTY LTD", "", "PEARL ENERGY NO: 1 PTY LTD", order=["mailbox", "document"])
        self.assertEqual((row["Folder name"], src), ("ONE PEARL ENERGY", "mailbox"))
        row, src, _ = companies.resolve("", "Pearl Energy No 2 Pty Ltd", "PEARL ENERGY NO: 1 PTY LTD")
        self.assertEqual((row["Folder name"], src), ("TWO PEARL ENERGY", "site"))
        row, src, _ = companies.resolve("", "", "PEARL ENERGY NO: 1 PTY LTD")
        self.assertEqual(src, "mailbox")
        self.assertIsNone(companies.resolve("", "", "unknown")[0])


class SiteTests(Isolated, unittest.TestCase):
    def setUp(self):
        super().setUp()
        rows = [["Ampol Beresfield North", "Pearl Energy No 1 Pty Ltd", "", "|".join(sites.make_fragments("19 Balook Drive, Beresfield NSW 2322")), "", "", "Yes", "", ""],
                ["Pearl Energy Beresfield", "Pearl Energy No 1 Pty Ltd", "Ampol Beresfield", "|".join(sites.make_fragments("6 Yangan Avenue, Beresfield NSW 2322")), "", "", "Yes", "", ""],
                ["Ampol Greta", "Pearl Energy No 2 Pty Ltd", "", "|".join(sites.make_fragments("2021 New England Highway, Greta NSW 2334")), "", "", "Yes", "", ""],
                ["Pearl Energy Greta", "Pearl Energy No 2 Pty Ltd", "", "|".join(sites.make_fragments("122 High Street, Greta NSW 2334")), "", "", "Yes", "", ""],
                ["Pearl Energy Barleigh Heads", "Pearl Energy No 3 Pty Ltd", "Pearl Energy Burleigh Heads", "|".join(sites.make_fragments("1 Cnr Oyster Creek Dr & Pacific Mwy BURLEIGH HEADS QLD 4220")), "", "", "Yes", "", ""]]
        self.add_rows("Sites", rows)

    def test_fragments(self):
        self.assertEqual(sites.make_fragments("19 Balook Drive , Beresfield NSW 2322 (north)"), ["19 balook", "balook drive beresfield"])
        self.assertIn("250 bridge", sites.make_fragments("242 - 250 Bridge Street, West Tamworth NSW 2340"))
        self.assertIn("pacific highway south kempsey", sites.make_fragments("Lot 1 & 2 Pacific Highway, South Kempsey, NSW,2440"))

    def test_two_sites_in_one_town_are_told_apart_by_address(self):
        self.assertEqual(sites.match("Deliver to 2021 New England Highway Greta", blocks=["2021 New England Highway Greta"])["site"], "Ampol Greta")
        self.assertEqual(sites.match("122 High Street Greta", blocks=["122 High Street Greta"])["site"], "Pearl Energy Greta")

    def test_town_alone_is_not_enough(self):
        m = sites.match("Delivery to Greta NSW", blocks=["Greta NSW"])
        self.assertEqual(m["site"], settings.UNASSIGNED)

    def test_old_name_alias_and_spelling(self):
        self.assertEqual(sites.match("Ampol Beresfield 6 Yangan Avenue")["site"], "Pearl Energy Beresfield")
        self.assertEqual(sites.match("Ship to PEARL ENERGY BURLEIGH HEADS")["site"], "Pearl Energy Barleigh Heads")

    def test_conflicting_address_and_name_is_not_guessed(self):
        m = sites.match("Pearl Energy Greta, 2021 New England Highway Greta", blocks=["2021 New England Highway Greta"])
        self.assertEqual(m["site"], settings.UNASSIGNED)
        self.assertIn("conflict", m["source"])

    def test_recipient_name_is_a_clue(self):
        self.assertEqual(sites.match("nothing useful", to_name="Pearl Energy Greta")["site"], "Pearl Energy Greta")


class WindowTests(unittest.TestCase):
    NOW = datetime.datetime(2026, 10, 3, 14, 0)

    def test_first_run_uses_default_days(self):
        s, e, kind = compute_window(self.NOW, default_days=2)
        self.assertEqual((s, e, kind), (self.NOW - datetime.timedelta(days=2), self.NOW, "catch-up"))

    def test_catch_up_starts_before_the_last_success(self):
        last = datetime.datetime(2026, 10, 3, 9, 0)
        s, _, _ = compute_window(self.NOW, last_success=last, overlap_minutes=60)
        self.assertEqual(s, datetime.datetime(2026, 10, 3, 8, 0))

    def test_catch_up_is_capped(self):
        s, _, _ = compute_window(self.NOW, last_success=datetime.datetime(2026, 1, 1), max_days=31)
        self.assertEqual(s, self.NOW - datetime.timedelta(days=31))

    def test_explicit_dates_are_inclusive(self):
        s, e, kind = compute_window(self.NOW, from_date=D(2026, 10, 1), to_date=D(2026, 10, 2))
        self.assertEqual((s, e, kind), (datetime.datetime(2026, 10, 1, 0, 0), datetime.datetime(2026, 10, 2, 23, 59, 59), "dates"))

    def test_future_end_is_capped_at_now(self):
        _, e, _ = compute_window(self.NOW, from_date=D(2026, 10, 1), to_date=D(2026, 12, 31))
        self.assertEqual(e, self.NOW)

    def test_single_day(self):
        s, e, _ = compute_window(self.NOW, from_date=D(2026, 9, 30), to_date=D(2026, 9, 30))
        self.assertEqual((s.day, e.day, e.hour), (30, 30, 23))


class MonthFolderTests(Isolated, unittest.TestCase):
    def test_names(self):
        self.assertEqual(month_folder_name(D(2026, 10, 3)), "OCTOBER 2026 EMAIL INVOICES")
        self.assertEqual(month_folder_name(D(2026, 1, 31)), "JANUARY 2026 EMAIL INVOICES")
        self.assertEqual(month_folder_name(D(2026, 9, 1), "{Month} {YEAR}"), "September 2026")
        self.assertEqual(month_folder_name(D(2026, 9, 1), "{YEAR}-{MM}"), "2026-09")
        self.assertEqual(len(MONTHS), 12)

    def test_pattern_is_read_from_the_settings(self):
        self.set_run(Month_folder_pattern="INVOICES {MM}-{YEAR}")
        self.assertEqual(month_folder_name(D(2026, 3, 9)), "INVOICES 03-2026")


class RouterTests(Isolated, unittest.TestCase):
    def test_live_mode_without_path_refuses(self):
        self.set_run(Destination_mode="live")
        with self.assertRaises(ConfigError):
            Router().check_config()

    def test_test_mode_builds_the_folders(self):
        r = Router()
        folder = r.prepare_company_dir(D(2026, 10, 3), "ONE PEARL ENERGY")
        self.assertTrue(os.path.isdir(folder))
        self.assertTrue(folder.endswith(os.path.join("OCTOBER 2026 EMAIL INVOICES", "ONE PEARL ENERGY")))

    def test_live_mode_never_creates_folders_when_the_like_test_switch_is_off(self):
        root = self.path("sharepoint")
        os.makedirs(root)
        self.set_run(Destination_mode="live", Live_destination_root=root, Live_mode_creates_missing_folders_like_test_mode="No",
                     Create_month_folder_in_live_mode="No", Create_company_folder_in_live_mode="No")
        with self.assertRaises(PlacementError) as cm:
            Router().prepare_company_dir(D(2026, 10, 3), "ONE PEARL ENERGY")
        self.assertIn("month folder", cm.exception.reason)
        os.makedirs(os.path.join(root, "OCTOBER 2026 EMAIL INVOICES"))
        with self.assertRaises(PlacementError) as cm:
            Router().prepare_company_dir(D(2026, 10, 3), "ONE PEARL ENERGY")
        self.assertIn("company folder", cm.exception.reason)

    def test_live_mode_can_create_when_allowed(self):
        root = self.path("sharepoint")
        os.makedirs(root)
        self.set_run(Destination_mode="live", Live_destination_root=root, Create_month_folder_in_live_mode="Yes",
                     Create_company_folder_in_live_mode="Yes")
        self.assertTrue(os.path.isdir(Router().prepare_company_dir(D(2026, 10, 3), "ONE PEARL ENERGY")))

    def test_live_root_missing_is_reported(self):
        self.set_run(Destination_mode="live", Live_destination_root=self.path("not_synced_yet"))
        with self.assertRaises(PlacementError) as cm:
            Router().prepare_company_dir(D(2026, 10, 3), "ONE PEARL ENERGY")
        self.assertIn("not reachable", cm.exception.reason)

    def test_config_py_path_is_used_when_the_sheet_is_empty(self):
        import config
        self.set_run(Destination_mode="live", Live_destination_root="")              # 5.15.0: the sheet has a default; emptied here
        old = config.LIVE_DESTINATION_ROOT
        try:
            config.LIVE_DESTINATION_ROOT = self.path("from_config")
            self.assertEqual(settings.destination_root(), self.path("from_config"))
            self.set_run(Live_destination_root=self.path("from_sheet"))
            self.assertEqual(settings.destination_root(), self.path("from_sheet"))       # the sheet wins
        finally:
            config.LIVE_DESTINATION_ROOT = old

    def test_place_never_overwrites(self):
        folder = self.path("dest")
        os.makedirs(folder)
        a, b = self.path("a.pdf"), self.path("b.pdf")
        write(a, b"AAAA")
        write(b, b"BBBBBB")
        r = Router()
        p1, how1 = r.place(a, folder, "x.pdf")
        p2, how2 = r.place(a, folder, "x.pdf")            # the same content: nothing to do
        p3, how3 = r.place(b, folder, "x.pdf")            # different content, same name: suffix
        self.assertEqual((how1, how2, how3), ("copied", "identical", "renamed"))
        self.assertEqual(os.path.basename(p3), "x_2.pdf")
        self.assertEqual(read_bytes(p1), b"AAAA")
        self.assertFalse([f for f in os.listdir(folder) if f.endswith(".part")])

    def test_dry_router_writes_nothing(self):
        r = Router(dry=True)
        r.prepare_company_dir(D(2026, 10, 3), "ONE PEARL ENERGY")
        src = self.path("a.pdf")
        write(src, b"A")
        r.place(src, r.company_dir(D(2026, 10, 3), "ONE PEARL ENERGY"), "x.pdf")
        r.hold(src, ("_Unassigned",), "x.pdf")
        self.assertFalse(os.path.exists(self.path("TestOutput")))
        self.assertFalse(os.path.exists(self.path("Holding")))

    def test_path_too_long_is_refused(self):
        r = Router()
        with self.assertRaises(PlacementError) as cm:
            r.prepare_company_dir(D(2026, 10, 3), "X" * 380)
        self.assertFalse(cm.exception.retryable)

    def test_month_end_warning(self):
        root = self.path("sp")
        os.makedirs(root)
        self.set_run(Destination_mode="live", Live_destination_root=root)
        r = Router()
        self.assertIsNotNone(r.month_warning(D(2026, 10, 30)))
        self.assertIsNone(r.month_warning(D(2026, 10, 10)))
        os.makedirs(os.path.join(root, "NOVEMBER 2026 EMAIL INVOICES"))
        self.assertIsNone(r.month_warning(D(2026, 10, 30)))

    def test_archive_and_cleanup(self):
        os.makedirs(self.path("Scratch"))
        src = self.path("Scratch", "f.pdf")
        write(src, b"X")
        kept = Router.archive(src, "R000001", "Invoice 1.pdf", 7)
        self.assertTrue(os.path.exists(kept))
        old = os.path.join(settings.archive_root(), "2020-01-01")
        os.makedirs(old)
        write(os.path.join(old, "x"), "x")
        os.makedirs(os.path.join(settings.archive_root(), "keep_me_not_a_date"))
        Router.cleanup_archive(7)
        self.assertFalse(os.path.exists(old))
        self.assertTrue(os.path.exists(kept))
        self.assertTrue(os.path.isdir(os.path.join(settings.archive_root(), "keep_me_not_a_date")))   # only date folders are ever deleted


class PluralKeywordTests(Isolated, unittest.TestCase):
    """Found in a real mailbox: 'Sales Invoices for Sales Order' was ignored because 'Invoices' is not the word 'invoice'."""

    def test_plural_subjects_are_found(self):
        for subject, typ in [("Sales Invoices for Sales Order", "Invoice"), ("Tax Invoices attached", "Invoice"), ("Credit Notes for September", "Credit note"),
                             ("Statements for Pearl Energy", "Statement"), ("Delivery Dockets 4455", "POD"), ("Adjustment Notes", "Credit note")]:
            self.assertEqual(doctypes.best_match(subject)[0], typ, subject)

    def test_short_keywords_and_other_words_stay_exact(self):
        for subject in ("Weekly stats report", "INVs attached", "Invoiced amounts", "Statementary notes", "Pods for sale"):
            self.assertIsNone(doctypes.best_match(subject)[0], subject)
        self.assertEqual(doctypes.best_match("Invoice 1234")[0], "Invoice")                           # the singular still works
        self.assertEqual(doctypes.best_match("INV8024")[0], "Invoice")                                # digits after a keyword still work

    def test_the_switch_turns_plurals_off(self):
        self.set_run(Keywords_also_match_plurals="No")
        self.assertIsNone(doctypes.best_match("Sales Invoices for Sales Order")[0])
        self.assertEqual(doctypes.best_match("Sales Invoice for Sales Order")[0], "Invoice")


class LockTests(Isolated, unittest.TestCase):
    def test_second_run_is_refused(self):
        path = self.path("data", "run.lock")
        with RunLock(path, 60):
            with self.assertRaises(LockBusy):
                RunLock(path, 60).__enter__()
        with RunLock(path, 60):                                # released after the first one ended
            pass

    def test_stale_lock_is_taken_over(self):
        path = self.path("data", "run.lock")
        os.makedirs(os.path.dirname(path))
        write(path, "old")
        old = time.time() - 3 * 3600
        os.utime(path, (old, old))
        with RunLock(path, 60):
            pass


class LockOwnerTests(Isolated, unittest.TestCase):
    """A run that was killed (crash, power cut, restart) must not block the next one for an hour."""

    def dead_pid(self):
        p = subprocess.Popen([sys.executable, "-c", "pass"])
        p.wait()
        return p.pid

    def test_a_lock_whose_process_is_gone_is_taken_over_at_once(self):
        path = self.path("data", "run.lock")
        os.makedirs(os.path.dirname(path))
        pid = self.dead_pid()
        write(path, f"pid={pid} started=2026-10-05 15:00:00")                   # fresh: far younger than the age limit
        lock = RunLock(path, 60)
        with lock:
            self.assertEqual(lock.took_over_from, pid)

    def test_a_lock_whose_process_is_running_is_respected(self):
        path = self.path("data", "run.lock")
        os.makedirs(os.path.dirname(path))
        other = subprocess.Popen([sys.executable, "-c", "import time; time.sleep(30)"])
        try:
            write(path, f"pid={other.pid} started=2026-10-05 15:00:00")
            with self.assertRaises(LockBusy):
                RunLock(path, 60).__enter__()
        finally:
            other.kill()
            other.wait()

    def test_a_lock_without_a_process_id_waits_for_the_age_rule(self):
        path = self.path("data", "run.lock")
        os.makedirs(os.path.dirname(path))
        write(path, "something unreadable")
        with self.assertRaises(LockBusy):
            RunLock(path, 60).__enter__()


class DatabaseTests(Isolated, unittest.TestCase):
    def test_duplicates_and_refs(self):
        conn = db.connect(":memory:")
        base = dict(status="Filed", supplier="Acme", doc_type="Invoice", number="1", file_hash="h1", received="2026-10-03 10:00:00")
        r1 = db.insert_doc(conn, base)
        r2 = db.insert_doc(conn, dict(base, number="2", file_hash="h2"))
        self.assertEqual((r1, r2), ("R000001", "R000002"))
        self.assertEqual(db.find_duplicate(conn, "Acme", "Invoice", "9", "h1")["ref"], "R000001")     # same content
        self.assertEqual(db.find_duplicate(conn, "Acme", "Invoice", "2", "zz")["ref"], "R000002")     # same number
        self.assertIsNone(db.find_duplicate(conn, "Acme", "Invoice", "3", "h3"))
        self.assertIsNone(db.find_duplicate(conn, "Other", "Invoice", "1", "h9"))

    def test_email_outcomes_count_attempts(self):
        conn = db.connect(":memory:")
        db.record_email(conn, "box", "E1", "s", "2026-10-03", "error", "why")
        db.record_email(conn, "box", "E1", "s", "2026-10-03", "error", "why")
        self.assertEqual(db.seen_emails(conn)[db.email_key("box", "E1")], ("error", 2))

    def test_backup_keeps_one_per_day(self):
        conn = db.connect(self.path("data", "bot.db"))
        folder = self.path("bk")
        self.assertIsNotNone(db.backup(conn, folder))
        self.assertIsNone(db.backup(conn, folder))
        old = os.path.join(folder, "bot_20200101.db")
        write(old, "x")
        db.backup(conn, folder, keep_days=14)
        self.assertFalse(os.path.exists(old))


if __name__ == "__main__":
    unittest.main()
