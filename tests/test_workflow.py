"""End-to-end tests: the whole bot against a simulated Outlook. Every rule that touches an e-mail or a file is checked here."""
import datetime
import os
import subprocess
import sys
import unittest

from openpyxl import load_workbook

from tests.helpers import COLYTON, DATA, NOW, ONE, ROOT, Isolated, connect_db, gateway, mail, run, write

from invoicebot import settings

OCT = "OCTOBER 2026 EMAIL INVOICES"
RED = "Forwarded / Reply"


def standard_mailbox(home):
    """A mailbox with one e-mail of every kind. Returns (gateway, {name: FakeMail})."""
    gw = gateway()
    jpg = os.path.join(home, "photo.jpg")
    write(jpg, os.urandom(60_000))
    m = {}
    add = lambda key, *a, **k: m.__setitem__(key, gw.add(ONE, mail(*a, **k)))
    add("invoice", "Tax Invoice 778899 Bidfood", ["invoice_bidfood.pdf"], hours_ago=3)
    add("credit", "Credit Note CN-9001", ["credit_note.pdf"], hours_ago=4)
    add("pod", "POD D-77120", ["pod.pdf"], hours_ago=5)
    add("statement", "Statement for September", ["statement.pdf"], hours_ago=6)
    add("summary", "Account Summary for 0001458043", ["account_summary.pdf"], hours_ago=7)
    add("overdue", "Overdue invoice reminder", ["invoice_other.pdf"], hours_ago=8)
    add("stop", "07254094 **Stop Supply Warning**", [], hours_ago=9)
    add("forward", "FW: Invoice NS-4040", ["invoice_company_named.pdf"], hours_ago=10)
    add("plain", "Meeting tomorrow at 10", [], hours_ago=11)
    add("image", "Invoice - photo attached", [jpg], hours_ago=12)
    add("corrupt", "inv", ["corrupt.pdf"], hours_ago=13)
    add("excel", "Invoice XL-5001", ["invoice_excel.xlsx"], hours_ago=14)
    add("rarity", "From Rarity Wholesale", ["statement_rarity.pdf"], hours_ago=15,
        sender_name="Rarity Wholesale", sender_email="accounts@raritywholesale.com.au")
    add("noattach", "Invoice 123 as discussed", [], hours_ago=16)
    add("fw_noattach", "FW: Tax invoice 555", [], hours_ago=17)
    return gw, m


def state(gw):
    return {(box, x.entry_id): (x.unread, x.flag, tuple(x.categories)) for box, ms in gw.boxes.items() for x in ms}


class RoutingTests(Isolated, unittest.TestCase):
    def setUp(self):
        super().setUp()
        self.gw, self.m = standard_mailbox(self.home)
        self.summary = run(self.gw)

    def test_files_are_in_the_right_folders_with_the_right_names(self):
        files = self.files_under("TestOutput")
        one = f"{OCT}/ONE PEARL ENERGY/"
        self.assertIn(one + "Bidfood Australia Ltd_778899.pdf", files)
        self.assertIn(one + "Acme Foods Pty Ltd_CN-9001_CN.pdf", files)
        self.assertIn(one + "Haulage Express Pty Ltd_D-77120_POD.pdf", files)
        self.assertIn(one + "Beta Distributors Pty Ltd_XL-5001.xlsx", files)
        self.assertIn(f"{OCT}/RADHE KRISHNA INVESTMENT/Nova Snacks Pty Ltd_NS-4040.pdf", files)   # the document names that company
        self.assertEqual(len(files), 5)

    def test_parked_and_unreadable_files(self):
        held = self.files_under("Holding")
        self.assertIn("_Statements/August 2026/ONE PEARL ENERGY/Fresh Fruit Co Pty Ltd_NoSite_12-08-2026_STM.pdf", held)                 # statements have their own folder
        self.assertIn("_Statements/August 2026/ONE PEARL ENERGY/Rarity_NoSite_12-08-2026_STM.pdf", held)                       # exception sender, no keyword in subject
        self.assertIn("_Unrecognised/October 2026/Account Summary/Drinks Group Pty Ltd_02-10-2026_ACCSUM.pdf", held)
        self.assertIn("_Unrecognised/October 2026/03-10-2026_corrupt.pdf", held)                          # the arrival date, dd-mm-yyyy
        self.assertFalse([f for f in held if f.startswith("_Unassigned")])

    def test_document_emails_are_read_and_ticked(self):
        for key in ("invoice", "credit", "pod", "statement", "summary", "excel", "rarity", "corrupt"):
            with self.subTest(key=key):
                self.assertFalse(self.m[key].unread)
                self.assertEqual(self.m[key].flag, 1)                      # complete (the tick)
                self.assertEqual(self.m[key].categories, [])

    def test_blocked_emails_are_flagged_and_never_downloaded(self):
        for key in ("overdue", "stop"):
            self.assertTrue(self.m[key].unread)
            self.assertEqual(self.m[key].flag, 2)                          # red flag
        self.assertEqual(self.summary.c["blocked"], 2)
        conn = connect_db(self.home)
        self.assertEqual(conn.execute("SELECT COUNT(*) FROM docs WHERE supplier LIKE 'Orbit%'").fetchone()[0], 0)

    def test_forwarded_document_gets_the_red_category(self):
        self.assertEqual(self.m["forward"].categories, [])
        self.assertEqual(self.m["forward"].flag, 2)
        self.assertTrue(self.m["forward"].unread)                          # 5.7.0: a forward is left UNREAD for a person (no tick)
        self.assertEqual(self.m["fw_noattach"].categories, [])
        self.assertTrue(self.m["fw_noattach"].unread)
        self.assertEqual(self.m["fw_noattach"].flag, 2)

    def test_image_only_email(self):
        self.assertEqual(self.m["image"].categories, ["Image attachment"])
        self.assertTrue(self.m["image"].unread)
        self.assertEqual(self.m["image"].flag, 0)

    def test_untouched_emails(self):
        for key in ("plain", "noattach"):
            self.assertTrue(self.m[key].unread)
            self.assertEqual((self.m[key].flag, self.m[key].categories), (0, []))

    def test_credit_note_amounts_are_negative(self):
        row = connect_db(self.home).execute("SELECT * FROM docs WHERE number='CN-9001'").fetchone()
        self.assertEqual((row["amount"], row["gst"], row["doc_type"], row["status"]), (-44.0, -4.0, "Credit note", "Filed"))

    def test_the_record_holds_everything(self):
        conn = connect_db(self.home)
        by = dict(conn.execute("SELECT status, COUNT(*) FROM docs GROUP BY status").fetchall())
        self.assertEqual(by, {"Filed": 5, "Parked": 3, "Unrecognised": 1})
        self.assertEqual(self.summary.c["documents"], 9)
        for r in conn.execute("SELECT * FROM docs WHERE status='Filed'"):
            self.assertTrue(os.path.exists(r["current_path"]), r["current_path"])
            self.assertTrue(r["ref"].startswith("R"))

    def test_originals_are_kept_in_the_archive_and_scratch_is_empty(self):
        self.assertEqual(self.files_under("Scratch"), [])
        self.assertEqual(len(self.files_under("Archive")), 9)

    def test_excel_files_for_people(self):
        wb = load_workbook(self.path("Reports", "Record_2026-10.xlsx"))
        rows = list(wb["Record"].iter_rows(values_only=True))
        self.assertEqual(rows[0][:4], ("Ref", "Supplier", "Document type", "Number"))
        self.assertEqual(len(rows) - 1, 9)                                  # arrival month = October
        numbers = {r[3] for r in rows[1:]}
        self.assertIn("778899", numbers)
        fix = load_workbook(self.path("Reports", "To_Fix.xlsx"))
        self.assertEqual(fix.sheetnames, ["How to use", "To fix", "Image-only emails", "Overdue-urgent emails", "New suppliers", "New sites", "Sites by company", "Learned sites", "Lists"])
        self.assertEqual(len([r for r in fix["To fix"].iter_rows(min_row=2, values_only=True) if r[0] and r[1] != "Parked"]), 1)   # only the corrupt file needs a person (NoSite statements are listed too, flagged, status Parked)
        self.assertEqual(len(list(fix["Overdue-urgent emails"].iter_rows(min_row=2, values_only=True))), 2)
        self.assertEqual(len(list(fix["Image-only emails"].iter_rows(min_row=2, values_only=True))), 1)

    def test_a_second_run_does_nothing_new_and_respects_manual_changes(self):
        self.m["overdue"].flag = 0                       # a person cleared the red flag
        before = os.listdir(self.path("TestOutput", OCT, "ONE PEARL ENERGY"))
        s2 = run(self.gw)
        self.assertEqual(s2.c["documents"], 0)
        self.assertEqual(s2.c["blocked"], 0)
        self.assertEqual(self.m["overdue"].flag, 0)       # the bot does not flag it again
        self.assertEqual(sorted(os.listdir(self.path("TestOutput", OCT, "ONE PEARL ENERGY"))), sorted(before))



class FlagTests(Isolated, unittest.TestCase):
    def test_existing_user_flag_is_never_overwritten(self):
        self.set_run(**{"Skip e-mails already flagged in Outlook": "No"})   # with the default (Yes) a flagged e-mail is not read at all
        gw = gateway()
        mine = gw.add(ONE, mail("Invoice 1 please", ["invoice_other.pdf"], flag=2))
        run(gw)
        self.assertEqual(mine.flag, 2)                    # still the person's own red flag
        self.assertFalse(mine.unread)                     # but it is marked as read: the invoice was filed


class ShadowTests(Isolated, unittest.TestCase):
    def test_shadow_changes_nothing_and_reports(self):
        gw, m = standard_mailbox(self.home)
        before = state(gw)
        summary = run(gw, shadow=True)
        self.assertEqual(state(gw), before)
        for folder in ("TestOutput", "Holding", "Archive"):
            self.assertFalse(os.path.exists(self.path(folder)), folder)
        self.assertFalse(os.path.exists(self.path("data", "bot.db")))
        self.assertEqual(self.files_under("Scratch"), [])
        reports = [f for f in self.files_under("Reports") if f.startswith("Shadow_Report_")]
        self.assertEqual(len(reports), 1)
        wb = load_workbook(self.path("Reports", reports[0]))
        rows = list(wb["Would do"].iter_rows(min_row=2, values_only=True))
        self.assertGreaterEqual(len(rows), 14)
        joined = " | ".join(str(r) for r in rows)
        self.assertIn("ONE PEARL ENERGY", joined)
        self.assertIn("red flag", joined)
        self.assertEqual(summary.c["Filed"], 5)

    def test_shadow_then_live_gives_the_same_result(self):
        gw, _ = standard_mailbox(self.home)
        run(gw, shadow=True)
        s = run(gw)
        self.assertEqual(s.c["Filed"], 5)


class LiveSetupTests(Isolated, unittest.TestCase):
    def test_a_damaged_settings_file_stops_the_run_instead_of_using_defaults(self):
        gw = gateway()
        gw.add(ONE, mail("Invoice INV-1001", ["invoice_acme.pdf"]))
        write(settings.settings_file(), "this is not an excel file")
        settings.reset()
        s = run(gw)
        self.assertEqual(s.stop_reason, "setup incomplete")
        self.assertEqual(gw.scan_calls, 0)
        self.assertFalse(os.path.exists(self.path("TestOutput")))

    def test_a_typo_in_destination_mode_stops_the_run(self):
        gw = gateway()
        self.set_run(Destination_mode="lve")
        s = run(gw)
        self.assertEqual(s.stop_reason, "setup incomplete")
        self.assertEqual(gw.scan_calls, 0)

    def test_live_mode_without_a_path_stops_and_changes_nothing(self):
        self.set_run(Destination_mode="live")
        gw, _ = standard_mailbox(self.home)
        before = state(gw)
        s = run(gw)
        self.assertEqual(s.stop_reason, "setup incomplete")
        self.assertEqual(gw.scan_calls, 0)
        self.assertEqual(state(gw), before)
        self.assertFalse(os.path.exists(self.path("TestOutput")))

    def test_live_month_and_company_folders_missing_then_created(self):
        root = self.path("sharepoint")
        os.makedirs(root)
        self.set_run(Destination_mode="live", Live_destination_root=root, Live_mode_creates_missing_folders_like_test_mode="No",
                     Create_month_folder_in_live_mode="No", Create_company_folder_in_live_mode="No")
        gw = gateway()
        em = gw.add(ONE, mail("Tax Invoice 778899", ["invoice_bidfood.pdf"]))
        s1 = run(gw)
        self.assertEqual(s1.c["Unassigned"], 1)
        self.assertIn("Bidfood Australia Ltd_778899.pdf", self.files_under("sharepoint", "HOLDINGS", "_Unassigned"))   # 5.15.0: HOLDINGS inside
        self.assertEqual(self.docs_under("sharepoint"), [])                       # no month / company folder created
        self.assertFalse(em.unread)                                               # the e-mail is done: the file is safe in holding
        # staff create next month's folder (and the company folder), the bot releases the file by itself
        os.makedirs(os.path.join(root, OCT, "ONE PEARL ENERGY"))
        s2 = run(gw)
        self.assertEqual(s2.c["released"], 1)
        self.assertEqual(self.docs_under("sharepoint"), [f"{OCT}/ONE PEARL ENERGY/Bidfood Australia Ltd_778899.pdf"])
        self.assertEqual(self.files_under("sharepoint", "HOLDINGS", "_Unassigned"), [])
        row = connect_db(self.home).execute("SELECT status FROM docs").fetchone()
        self.assertEqual(row["status"], "Filed")

    def test_live_mode_writes_into_the_configured_root(self):
        root = self.path("sharepoint")
        os.makedirs(os.path.join(root, OCT, "ONE PEARL ENERGY"))
        self.set_run(Destination_mode="live", Live_destination_root=root)
        gw = gateway()
        gw.add(ONE, mail("Tax Invoice 778899", ["invoice_bidfood.pdf"]))
        run(gw)
        self.assertEqual(self.docs_under("sharepoint"), [f"{OCT}/ONE PEARL ENERGY/Bidfood Australia Ltd_778899.pdf"])
        self.assertIn("Reports/To_Fix.xlsx", self.files_under("sharepoint"))      # 5.15.0: Reports inside the live folder
        self.assertFalse(os.path.exists(self.path("TestOutput")))


class RulesInActionTests(Isolated, unittest.TestCase):
    def test_duplicate_invoice_in_two_mailboxes_is_saved_once(self):
        gw = gateway()
        a = gw.add(ONE, mail("Invoice INV-1001", ["invoice_acme.pdf"], hours_ago=3))
        b = gw.add(COLYTON, mail("FW: Invoice INV-1001", ["invoice_acme.pdf"], hours_ago=2))
        s = run(gw)
        self.assertEqual(s.c["duplicates"], 1)
        self.assertEqual(len(self.files_under("TestOutput")), 1)
        self.assertFalse(a.unread)
        self.assertTrue(b.unread)                                          # 5.7.0: the forward stays unread, with the red category

    def test_month_comes_from_the_arrival_date(self):
        now = datetime.datetime(2026, 10, 1, 1, 0)
        gw = gateway()
        gw.add(ONE, mail("Invoice INV-1001", ["invoice_acme.pdf"], hours_ago=1.5, now=now))        # arrived 30 Sept 23:30
        run(gw, now=now)
        self.assertEqual(self.files_under("TestOutput"), ["SEPTEMBER 2026 EMAIL INVOICES/ONE PEARL ENERGY/Acme Foods Pty Ltd_INV-1001.pdf"])

    def test_month_can_follow_the_invoice_date(self):
        self.set_run(Month_basis="invoice")
        now = datetime.datetime(2026, 10, 1, 1, 0)
        gw = gateway()
        gw.add(ONE, mail("Invoice INV-1001", ["invoice_acme.pdf"], hours_ago=1.5, now=now))
        run(gw, now=now)
        self.assertTrue(self.files_under("TestOutput")[0].startswith("OCTOBER 2026"))             # the invoice is dated 2 Oct

    def test_company_order_setting(self):
        self.set_run(Company_source_order="mailbox,document,site")
        gw = gateway()
        gw.add(ONE, mail("FW: Invoice NS-4040", ["invoice_company_named.pdf"]))
        run(gw)
        self.assertEqual(self.files_under("TestOutput"), [f"{OCT}/ONE PEARL ENERGY/Nova Snacks Pty Ltd_NS-4040.pdf"])

    def test_supplier_folder_override(self):
        self.add_rows("Suppliers", [["Acme Foods", "Acme Foods Pty Ltd", "Acme", "MILK INVOICES"]])
        gw = gateway()
        gw.add(ONE, mail("Invoice INV-1001", ["invoice_acme.pdf"]))
        run(gw)
        self.assertEqual(self.files_under("TestOutput"), [f"{OCT}/MILK INVOICES/Acme Foods Pty Ltd_INV-1001.pdf"])

    def test_page_limit_protects_memory(self):
        gw = gateway()
        gw.add(ONE, mail("Invoice LF-1", ["invoice_20_pages.pdf"]))
        s = run(gw)
        self.assertEqual(s.c["Unrecognised"], 1)
        reason = connect_db(self.home).execute("SELECT reason FROM docs").fetchone()[0]
        self.assertIn("too many pages", reason)

    def test_oversized_attachment_is_not_read(self):
        self.set_run(Max_attachment_MB="0")
        gw = gateway()
        gw.add(ONE, mail("Invoice INV-1001", ["invoice_acme.pdf"]))
        s = run(gw)
        self.assertEqual(s.c["Unrecognised"], 1)
        self.assertIn("larger than", connect_db(self.home).execute("SELECT reason FROM docs").fetchone()[0])

    def test_blocked_word_in_the_body(self):
        self.set_run(Scan_body_for_blocked_words="Yes")
        gw = gateway()
        em = gw.add(ONE, mail("Invoice 5 for September", ["invoice_acme.pdf"], body="This is a final reminder, please pay."))
        s = run(gw)
        self.assertEqual((s.c["blocked"], s.c["documents"], em.flag), (1, 0, 2))

    def test_first_time_supplier_can_be_held(self):
        self.set_run(**{"Hold_first-time_suppliers": "Yes"})
        gw = gateway()
        gw.add(ONE, mail("Invoice INV-1001", ["invoice_acme.pdf"]))
        s = run(gw)
        self.assertEqual(s.c["Unassigned"], 1)
        self.assertIn("first time", connect_db(self.home).execute("SELECT reason FROM docs").fetchone()[0])

    def test_red_category_for_every_fw_re_when_asked(self):
        self.set_run(Red_category_applies_to="all")
        gw = gateway()
        em = gw.add(ONE, mail("RE: lunch on Friday"))
        run(gw)
        self.assertEqual((em.categories, em.flag), ([], 2))

    def test_date_window_limits_what_is_read(self):
        gw = gateway()
        gw.add(ONE, mail("Invoice INV-1001", ["invoice_acme.pdf"], hours_ago=3))
        gw.add(ONE, mail("Invoice 778899", ["invoice_bidfood.pdf"], hours_ago=24 * 6))              # 6 days ago: outside the 2-day catch-up
        s = run(gw)
        self.assertEqual(s.c["documents"], 1)
        s = run(gw, from_date=datetime.date(2026, 9, 25), to_date=datetime.date(2026, 10, 3))
        self.assertEqual(s.c["documents"], 1)                                                       # the older one, now in range
        self.assertEqual(len(self.files_under("TestOutput")), 2)

    def test_path_emails_to_site_handlers(self):
        from invoicebot import sites
        self.set_run(Email_path_mode="draft")
        self.add_rows("Sites", [["Pearl Energy Colyton", "Colyton Fuels Pty Ltd", "", "|".join(sites.make_fragments("88-90 Great Western Hwy, Colyton NSW 2760")),
                                 "Sam", "sam@example.com", "Yes", "", ""]])
        gw = gateway()
        gw.add(ONE, mail("Invoice INV-1001", ["invoice_acme.pdf"]))
        run(gw)
        self.assertEqual(len(gw.drafts), 1)
        self.assertEqual(gw.drafts[0]["to"], "sam@example.com")
        self.assertIn("Acme Foods Pty Ltd_INV-1001.pdf", gw.drafts[0]["body"])
        self.assertEqual(gw.sent, [])


class SafetyTests(Isolated, unittest.TestCase):
    def test_a_running_bot_blocks_a_second_one(self):
        write(self.path("data", "run.lock"), "pid=1")
        gw = gateway()
        gw.add(ONE, mail("Invoice INV-1001", ["invoice_acme.pdf"]))
        s = run(gw)
        self.assertEqual(s.stop_reason, "another run in progress")
        self.assertEqual(gw.scan_calls, 0)

    def test_no_edit_permission_in_outlook_does_not_lose_documents(self):
        gw, m = standard_mailbox(self.home)
        gw.read_only = True
        s = run(gw)
        self.assertGreater(s.c["outlook errors"], 0)
        self.assertEqual(s.c["Filed"], 5)                                           # the documents are still filed
        s2 = run(gw)
        self.assertEqual(s2.c["documents"], 0)                                      # and not processed again

    def test_failed_download_is_retried_then_given_up(self):
        gw = gateway()
        em = gw.add(ONE, mail("Invoice INV-1001", ["invoice_acme.pdf"]))
        gw.save_failures = {(em.entry_id, "invoice_acme.pdf")}
        s1 = run(gw)
        self.assertEqual((s1.c["email errors"], s1.c["documents"]), (1, 0))
        self.assertTrue(em.unread)                                                  # untouched: it will be retried
        s2 = run(gw)
        self.assertEqual(s2.c["email errors"], 1)
        s3 = run(gw)                                                                # third failure: give up
        outcome = connect_db(self.home).execute("SELECT outcome, attempts FROM emails").fetchone()
        self.assertEqual((outcome["outcome"], outcome["attempts"]), ("gave-up", 3))
        s4 = run(gw)
        self.assertEqual(s4.c["email errors"], 0)                                   # no more attempts on a given-up e-mail

    def test_failed_download_works_on_the_retry(self):
        gw = gateway()
        em = gw.add(ONE, mail("Invoice INV-1001", ["invoice_acme.pdf"]))
        gw.save_failures = {(em.entry_id, "invoice_acme.pdf")}
        run(gw)
        gw.save_failures = set()                                                    # the problem went away
        s = run(gw)
        self.assertEqual(s.c["Filed"], 1)
        self.assertFalse(em.unread)

    def test_one_failed_attachment_means_the_email_is_not_ticked(self):
        gw = gateway()
        em = gw.add(ONE, mail("Invoice 2 files", ["invoice_acme.pdf", "invoice_bidfood.pdf"]))
        gw.save_failures = {(em.entry_id, "invoice_bidfood.pdf")}
        s = run(gw)
        self.assertEqual((s.c["Filed"], s.c["email errors"]), (1, 1))
        self.assertTrue(em.unread)                                                  # not ticked: the second file is still to come
        gw.save_failures = set()
        s = run(gw)
        self.assertEqual((s.c["Filed"], s.c["duplicates"]), (1, 1))                 # the first is recognised, the second is filed
        self.assertFalse(em.unread)
        self.assertEqual(len(self.files_under("TestOutput")), 2)

    def test_a_missing_mailbox_is_skipped_not_fatal(self):
        gw = gateway()
        del gw.boxes["Two@pearlenergy.com.au"]
        gw.add(ONE, mail("Invoice INV-1001", ["invoice_acme.pdf"]))
        s = run(gw)
        self.assertEqual(s.c["Filed"], 1)
        self.assertTrue(any("not found" in line for line in s.mailbox_lines))

    def test_incomplete_windows_are_repeated(self):
        self.set_run(Max_emails_per_mailbox_per_run="2")
        gw = gateway()
        for i in range(5):
            gw.add(ONE, mail(f"Tax Invoice {778890 + i}", ["invoice_bidfood.pdf"], hours_ago=1 + i))
        s = run(gw)
        self.assertTrue(any("window kept open" in line for line in s.mailbox_lines))
        conn = connect_db(self.home)
        self.assertEqual(conn.execute("SELECT COUNT(*) FROM mailbox_runs WHERE ok=1 AND mailbox LIKE 'PEARL ENERGY NO: 1%'").fetchone()[0], 0)

    def test_cannot_open_to_fix_does_not_stop_a_run(self):
        write(self.path("Reports", "To_Fix.xlsx"), "this is not an excel file")
        gw = gateway()
        gw.add(ONE, mail("Invoice INV-1001", ["invoice_acme.pdf"]))
        self.assertEqual(run(gw).c["Filed"], 1)


class ToFixTests(Isolated, unittest.TestCase):
    def test_staff_fix_an_unreadable_document(self):
        gw = gateway()
        gw.add(ONE, mail("inv", ["corrupt.pdf"]))
        run(gw)
        conn = connect_db(self.home)
        ref = conn.execute("SELECT ref FROM docs").fetchone()[0]
        self.assertEqual(conn.execute("SELECT status FROM docs").fetchone()[0], "Unrecognised")
        wb = load_workbook(self.path("Reports", "To_Fix.xlsx"))
        ws = wb["To fix"]
        self.assertEqual(ws.cell(2, 1).value, ref)
        ws.cell(2, 13).value, ws.cell(2, 14).value, ws.cell(2, 15).value = "Fixed Supplier Ltd", "FX-1", "Invoice"
        wb.save(self.path("Reports", "To_Fix.xlsx"))
        s = run(gw)
        self.assertEqual(s.c["fixes"], 1)                                           # one document fixed (three fields)
        self.assertEqual(s.c["released"], 1)
        self.assertEqual(self.files_under("TestOutput"), [f"{OCT}/ONE PEARL ENERGY/Fixed Supplier Ltd_FX-1.pdf"])
        self.assertEqual(self.files_under("Holding", "_Unrecognised"), [])
        row = connect_db(self.home).execute("SELECT status, supplier, number FROM docs").fetchone()
        self.assertEqual(tuple(row), ("Filed", "Fixed Supplier Ltd", "FX-1"))
        self.assertEqual(connect_db(self.home).execute("SELECT COUNT(*) FROM corrections").fetchone()[0], 3)

    def test_staff_choose_the_company_folder_of_a_waiting_file(self):
        self.set_run(**{"Hold_first-time_suppliers": "Yes"})
        gw = gateway()
        gw.add(ONE, mail("Invoice INV-1001", ["invoice_acme.pdf"]))
        run(gw)
        self.assertEqual(self.files_under("TestOutput"), [])
        self.set_run(**{"Hold_first-time_suppliers": "No"})                                  # the supplier is now confirmed
        wb = load_workbook(self.path("Reports", "To_Fix.xlsx"))
        wb["To fix"].cell(2, 12).value = "TWO PEARL ENERGY"
        wb.save(self.path("Reports", "To_Fix.xlsx"))
        run(gw)
        self.assertEqual(self.files_under("TestOutput"), [f"{OCT}/TWO PEARL ENERGY/Acme Foods Pty Ltd_INV-1001.pdf"])


class SenderTests(Isolated, unittest.TestCase):
    """One shared mailbox (jiya) receives the invoices of every company, forwarded from the company mailboxes."""

    def setUp(self):
        super().setUp()
        from invoicebot import mailboxes
        mailboxes.apply([{"outlook": "jiya@pearlenergy.com.au", "action": "test"}])

    def test_each_file_goes_to_the_company_that_sent_it(self):
        gw = gateway("jiya@pearlenergy.com.au")
        box = "jiya@pearlenergy.com.au"
        for sender, subject, f in [("P & O Investment", "RE: Invoice 778899", "invoice_bidfood.pdf"),
                                   ("Highland Energy 1", "RE: Invoice INV-1001", "invoice_acme.pdf"),
                                   ("Pearl Energy Colyton", "Invoice OW-3030", "invoice_other.pdf"),
                                   ("Shri Hari Investment", "Goodman Baking - Delivery Docket", "pod.pdf"),
                                   ("Some Supplier Pty Ltd", "Credit Note CN-9001", "credit_note.pdf")]:       # unknown sender: the mailbox decides
            gw.add(box, mail(subject, [f], sender_name=sender))
        s = run(gw, only="jiya")
        self.assertEqual(s.c["Filed"], 5)
        files = self.files_under("TestOutput")
        for folder, name in [("P & O INVESTMENT", "Bidfood Australia Ltd_778899.pdf"), ("HIGHLAND ENERGY 1", "Acme Foods Pty Ltd_INV-1001.pdf"),
                             ("COLYTON FUELS", "Orbit Wholesale Pty Ltd_OW-3030.pdf"), ("SHRI HARI INVESTMENT", "Haulage Express Pty Ltd_D-77120_POD.pdf"),
                             ("TEST JIYA", "Acme Foods Pty Ltd_CN-9001_CN.pdf")]:
            self.assertIn(f"{OCT}/{folder}/{name}", files)
        sources = dict(connect_db(self.home).execute("SELECT number, company_source FROM docs").fetchall())
        self.assertEqual((sources["778899"], sources["CN-9001"]), ("sender", "mailbox"))

    def test_a_company_named_on_the_document_beats_the_sender(self):
        gw = gateway("jiya@pearlenergy.com.au")
        gw.add("jiya@pearlenergy.com.au", mail("FW: Invoice NS-4040", ["invoice_company_named.pdf"], sender_name="Pearl Energy Colyton"))
        run(gw, only="jiya")
        self.assertEqual(self.files_under("TestOutput"), [f"{OCT}/RADHE KRISHNA INVESTMENT/Nova Snacks Pty Ltd_NS-4040.pdf"])


class HeldFileNamingAndReasonsTests(Isolated, unittest.TestCase):
    """An unrecognised document with a supplier and a number is named like any other; the run says WHY documents wait."""

    def future_dated(self):
        from tests import make_data
        path = os.path.join(self.home, "future.pdf")
        make_data.invoice(path, "Shady Supplies Pty Ltd", "SH-55501", net=100.0, date="11/10/2099")      # a date that cannot be right
        return path

    def test_a_failed_invoice_is_still_named_supplier_and_number(self):
        self.set_run(Month_folder_follows="arrival")                  # the old rule: a future date is a problem (the new month window files it in the current month)
        gw = gateway()
        gw.add(ONE, mail("Invoice 100 date wrong", [self.future_dated()]))
        gw.add(ONE, mail("inv", ["corrupt.pdf"]))
        s = run(gw)
        held = sorted(self.files_under("Holding"))
        self.assertEqual(s.c["Unrecognised"], 2)
        self.assertIn("_Unrecognised/October 2026/Shady Supplies Pty Ltd_SH-55501.pdf", held)                                                 # supplier_number
        self.assertIn("_Unrecognised/October 2026/03-10-2026_corrupt.pdf", held)                                                              # nothing was read: the old name, with the date

    def test_the_summary_counts_the_reasons(self):
        self.set_run(Month_folder_follows="arrival")
        gw = gateway()
        for i in range(3):
            gw.add(ONE, mail(f"inv {i}", ["corrupt.pdf"], hours_ago=1 + i))
        gw.add(ONE, mail("Invoice 100 date wrong", [self.future_dated()]))
        s = run(gw)
        text = s.text()
        self.assertIn("Why documents are waiting or unrecognised", text)
        self.assertRegex(text, r"3 x could not open the file")
        self.assertIn("is in the future", text)

    def test_the_first_outlook_problem_is_shown(self):
        class Refusing(type(gateway())):
            def mark_done(self, info):
                raise RuntimeError("the tick could not be set (FlagStatus: boom)")
        gw = Refusing()
        for r in __import__("invoicebot.companies", fromlist=["x"]).all_active():
            gw.add_mailbox(__import__("invoicebot.companies", fromlist=["x"]).mailbox_names(r)[0])
        gw.add(ONE, mail("Tax Invoice 778899", ["invoice_bidfood.pdf"]))
        s = run(gw)
        self.assertEqual(s.c["outlook errors"], 1)
        self.assertIn("first problem: mark_done: the tick could not be set (FlagStatus: boom)", s.text())


class CrashRecoveryTests(Isolated, unittest.TestCase):
    def test_a_run_after_a_crash_carries_on_and_says_so(self):
        import subprocess
        dead = subprocess.Popen([sys.executable, "-c", "pass"])
        dead.wait()
        write(self.path("data", "run.lock"), f"pid={dead.pid} started=2026-10-03 13:30:00")               # the lock a killed run leaves behind
        gw = gateway()
        gw.add(ONE, mail("Tax Invoice 778899", ["invoice_bidfood.pdf"]))
        s = run(gw)
        self.assertEqual(s.c["Filed"], 1)
        self.assertTrue(any("previous run did not finish" in w for w in s.warnings), s.warnings)
        self.assertFalse(os.path.exists(self.path("data", "run.lock")))                                   # and it removes the lock at its end


class AlreadyTickedTests(Isolated, unittest.TestCase):
    """A new install (empty database) must not read and file e-mails that were already ticked in Outlook."""

    def test_a_ticked_email_is_not_read_again_even_with_an_empty_database(self):
        from invoicebot.outlook_io import OL_FLAG_COMPLETE
        gw = gateway()
        gw.add(ONE, mail("Tax Invoice 778899", ["invoice_bidfood.pdf"], unread=False, flag=OL_FLAG_COMPLETE))
        s = run(gw)
        self.assertEqual(s.c["documents"], 0)
        self.assertEqual(s.c["already ticked"], 1)

    def test_the_fallback_category_counts_as_done_too(self):
        gw = gateway()
        gw.add(ONE, mail("Tax Invoice 778899", ["invoice_bidfood.pdf"], unread=False, categories=["Processed by bot"]))
        s = run(gw)
        self.assertEqual((s.c["documents"], s.c["already ticked"]), (0, 1))

    def test_an_unticked_email_is_still_read(self):
        gw = gateway()
        gw.add(ONE, mail("Tax Invoice 778899", ["invoice_bidfood.pdf"]))
        self.assertEqual(run(gw).c["documents"], 1)

    def test_the_switch_can_turn_the_skipping_off(self):
        from invoicebot.outlook_io import OL_FLAG_COMPLETE
        self.set_run(**{"Skip e-mails already ticked in Outlook": "No"})
        gw = gateway()
        gw.add(ONE, mail("Tax Invoice 778899", ["invoice_bidfood.pdf"], unread=False, flag=OL_FLAG_COMPLETE))
        self.assertEqual(run(gw).c["documents"], 1)


class AlreadyFlaggedTests(Isolated, unittest.TestCase):
    def test_a_red_flagged_email_is_not_read(self):
        from invoicebot.outlook_io import OL_FLAG_MARKED
        gw = gateway()
        gw.add(ONE, mail("Tax Invoice 778899", ["invoice_bidfood.pdf"], flag=OL_FLAG_MARKED))
        s = run(gw)
        self.assertEqual((s.c["documents"], s.c["already ticked"]), (0, 1))

    def test_the_flag_switch_can_turn_it_off(self):
        from invoicebot.outlook_io import OL_FLAG_MARKED
        self.set_run(**{"Skip e-mails already flagged in Outlook": "No"})
        gw = gateway()
        gw.add(ONE, mail("Tax Invoice 778899", ["invoice_bidfood.pdf"], flag=OL_FLAG_MARKED))
        self.assertEqual(run(gw).c["documents"], 1)


class ParallelSafetyTests(Isolated, unittest.TestCase):
    """The parallel readers can never freeze a run (a real hang was found and fixed in version 5.4.1)."""

    def jobs(self, n=4):
        path = os.path.join(DATA, "invoice_bidfood.pdf")
        return [{"job_id": i, "path": path, "ctx": {"subject": "Tax Invoice 778899", "orig_name": "invoice_bidfood.pdf"}, "info": None} for i in range(n)]

    def pipeline(self, pool):
        from invoicebot.pipeline import Pipeline
        p = Pipeline(gateway(), log=lambda *a, **k: None, workers=2, now=NOW)
        p._setup()
        p.conf["low_ram"] = 0
        p.pool, p.parallel_timeout = pool, 0.3
        return p

    def test_a_stuck_reader_is_given_up_and_the_batch_is_read_here_instead(self):
        import concurrent.futures

        class StuckPool:
            killed = False
            def submit(self, fn, job):
                return concurrent.futures.Future()                                 # never completes: a reader that hangs
            def shutdown(self, wait=True, cancel_futures=False):
                StuckPool.killed = True

        p = self.pipeline(StuckPool())
        results = p._analyse_raw(self.jobs(3))
        self.assertEqual([r["invoice_number"] for r in results], ["778899"] * 3)      # every document was still read
        self.assertTrue(StuckPool.killed)
        self.assertIsNone(p.pool)                                                    # a new pool is made for the next batch
        self.assertTrue(any("parallel reading had a problem" in w for w in p.summary.warnings))

    def test_a_crashed_reader_process_is_handled(self):
        from concurrent.futures.process import BrokenProcessPool

        class CrashedPool:
            def submit(self, fn, job):
                raise BrokenProcessPool("a reader process died")
            def shutdown(self, wait=True, cancel_futures=False):
                pass

        p = self.pipeline(CrashedPool())
        self.assertEqual([r["invoice_number"] for r in p._analyse_raw(self.jobs(2))], ["778899"] * 2)

    def test_a_file_that_hangs_is_skipped_not_waited_for_forever(self):
        import concurrent.futures
        from invoicebot.pipeline import Pipeline

        class StuckPool:
            def submit(self, fn, job):
                return concurrent.futures.Future()                                 # a reader that never answers
            def shutdown(self, wait=True, cancel_futures=False):
                pass

        p = Pipeline(gateway(), log=lambda *a, **k: None, now=NOW)                 # no explicit worker count: the guarded way
        p._setup()
        p.pool, p.parallel_timeout = StuckPool(), 0.3
        res = p._read_one(self.jobs(1)[0])
        self.assertTrue(res["fatal"])
        self.assertTrue(any("took too long" in x for x in res["problems"]), res["problems"])
        self.assertIsNone(p.pool)

    def test_the_readers_are_recycled_by_the_program_after_n_documents(self):
        self.set_run(Recycle_worker_after_documents="10")
        made = []

        class OkPool:
            def __init__(self):
                made.append(self)
                self.closed = False
            def submit(self, fn, job):
                import concurrent.futures
                f = concurrent.futures.Future()
                f.set_result(fn(job))
                return f
            def shutdown(self, wait=True, cancel_futures=False):
                self.closed = True

        p = self.pipeline(None)
        p._pool = lambda: p.pool if p.pool is not None else (setattr(p, "pool", OkPool()) or setattr(p, "pool_tasks", 0) or p.pool)
        for _ in range(4):
            p._analyse_raw(self.jobs(8))                                              # 8 per batch; 10 documents per reader x 2 readers = 20
        self.assertEqual(len(made), 2)                                               # replaced once, after the third batch (24 >= 20); batch 4 used the new one
        self.assertTrue(made[0].closed)

    def test_real_processes_survive_being_recycled_in_the_middle_of_a_run(self):
        """The run that used to hang: more documents than a reader may read, so the readers are replaced while work continues."""
        import shutil
        from unittest import mock
        from invoicebot import pipeline
        self.set_run(Recycle_worker_after_documents="10")
        gw = gateway()
        for i in range(40):
            copy = os.path.join(self.home, f"copy{i}.pdf")
            shutil.copyfile(os.path.join(DATA, "invoice_bidfood.pdf"), copy)
            gw.add(ONE, mail(f"Tax Invoice 7788{i:02d}", [copy], hours_ago=1 + i / 100))
        with mock.patch.object(pipeline, "PARALLEL_TIMEOUT_SECONDS", 25):               # if the old bug came back, this fails instead of freezing
            s = run(gw, workers=2)
        self.assertEqual(s.c["Filed"] + s.c["duplicates"], 40)                           # the same file 40 times: one filed, 39 recognised as duplicates
        self.assertFalse([w for w in s.warnings if "parallel" in w])


class EfficiencyTests(Isolated, unittest.TestCase):
    """The Excel files are written only when something they show has changed."""

    def reports(self):
        return {f: os.path.getmtime(self.path("Reports", f)) for f in os.listdir(self.path("Reports"))}

    def age_the_reports(self):
        old = os.path.getmtime(self.path("Reports", "To_Fix.xlsx")) - 3600
        for f in self.reports():
            os.utime(self.path("Reports", f), (old, old))
        return self.reports()

    def first_run(self):
        gw = gateway()
        gw.add(ONE, mail("Tax Invoice 778899", ["invoice_bidfood.pdf"], hours_ago=3))
        gw.add(ONE, mail("inv", ["corrupt.pdf"], hours_ago=4))
        s = run(gw)
        return gw, s

    def test_the_report_has_no_summary_sheet(self):
        self.first_run()
        wb = load_workbook(self.path("Reports", "Record_2026-10.xlsx"))
        self.assertEqual(wb.sheetnames, ["Record", "Needs attention"])
        self.assertEqual(wb.active.title, "Record")
        wb.close()

    def test_an_idle_run_writes_no_excel_file(self):
        gw, s1 = self.first_run()
        self.assertIn("Excel files updated: Record_2026-10, To_Fix", s1.notes)
        before = self.age_the_reports()
        s2 = run(gw)
        self.assertEqual(s2.c["documents"], 0)
        self.assertEqual(self.reports(), before)                                    # not one file was touched
        self.assertIn("Excel files updated: none (nothing new)", s2.notes)

    def test_a_new_document_rewrites_what_it_changes(self):
        gw, _ = self.first_run()
        before = self.age_the_reports()
        gw.add(ONE, mail("Credit Note CN-9001", ["credit_note.pdf"], hours_ago=0.5))
        s = run(gw)
        self.assertEqual(s.c["Filed"], 1)
        after = self.reports()
        self.assertGreater(after["Record_2026-10.xlsx"], before["Record_2026-10.xlsx"])
        self.assertGreater(after["To_Fix.xlsx"], before["To_Fix.xlsx"])
        wb = load_workbook(self.path("Reports", "Record_2026-10.xlsx"))
        self.assertEqual(len(list(wb["Record"].iter_rows(min_row=2, values_only=True))), 3)
        wb.close()

    def test_a_blocked_email_updates_to_fix_but_not_the_record(self):
        gw, _ = self.first_run()
        before = self.age_the_reports()
        gw.add(ONE, mail("Overdue invoice reminder", [], hours_ago=0.5))
        run(gw)
        after = self.reports()
        self.assertEqual(after["Record_2026-10.xlsx"], before["Record_2026-10.xlsx"])          # no document changed
        self.assertGreater(after["To_Fix.xlsx"], before["To_Fix.xlsx"])                         # the overdue list did
        wb = load_workbook(self.path("Reports", "To_Fix.xlsx"))
        self.assertEqual(len([r for r in wb["Overdue-urgent emails"].iter_rows(min_row=2, values_only=True) if r[0]]), 1)
        wb.close()

    def test_a_deleted_file_is_created_again(self):
        gw, _ = self.first_run()
        os.remove(self.path("Reports", "Record_2026-10.xlsx"))
        os.remove(self.path("Reports", "To_Fix.xlsx"))
        s = run(gw)
        self.assertEqual(s.c["documents"], 0)
        self.assertIn("Excel files updated: Record_2026-10, To_Fix", s.notes)
        self.assertTrue(os.path.exists(self.path("Reports", "To_Fix.xlsx")))

    def test_a_fix_that_does_not_release_the_file_still_updates_the_record(self):
        """A staff fix that leaves the file unrecognised must still reach the monthly Record (it was stale once)."""
        gw = gateway()
        gw.add(ONE, mail("inv", ["corrupt.pdf"]))
        run(gw)
        wb = load_workbook(self.path("Reports", "To_Fix.xlsx"))
        wb["To fix"].cell(2, 13).value = "Typed Supplier Ltd"                          # a supplier only: still no document type
        wb.save(self.path("Reports", "To_Fix.xlsx"))
        wb.close()
        s = run(gw)
        self.assertEqual((s.c["released"], s.c["fixes"]), (0, 1))
        rec = load_workbook(self.path("Reports", "Record_2026-10.xlsx"))
        self.assertEqual([r[1] for r in rec["Record"].iter_rows(min_row=2, values_only=True)], ["Typed Supplier Ltd"])
        rec.close()

    def test_a_staff_fix_updates_the_files(self):
        gw, _ = self.first_run()
        wb = load_workbook(self.path("Reports", "To_Fix.xlsx"))
        ws = wb["To fix"]
        ws.cell(2, 13).value, ws.cell(2, 14).value, ws.cell(2, 15).value = "Fixed Supplier Ltd", "FX-1", "Invoice"
        wb.save(self.path("Reports", "To_Fix.xlsx"))
        wb.close()
        before = self.age_the_reports()
        s = run(gw)
        self.assertEqual(s.c["released"], 1)
        self.assertGreater(self.reports()["Record_2026-10.xlsx"], before["Record_2026-10.xlsx"])


class StatementTests(Isolated, unittest.TestCase):
    """Three holding folders side by side: _Unassigned, _Unrecognised and _Statements."""

    def test_statements_have_their_own_folder_and_invoices_are_not_mixed_in(self):
        gw = gateway()
        for subject, f in [("Tax Invoice 778899", "invoice_bidfood.pdf"), ("Credit Note CN-9001", "credit_note.pdf"),
                           ("POD D-77120", "pod.pdf"), ("Statement for September", "statement.pdf")]:
            gw.add(ONE, mail(subject, [f]))
        s = run(gw)
        self.assertEqual({f.rsplit("/", 1)[0] for f in self.files_under("TestOutput")}, {f"{OCT}/ONE PEARL ENERGY"})   # 3 documents, one folder
        self.assertEqual(self.files_under("Holding"), ["_Statements/August 2026/ONE PEARL ENERGY/Fresh Fruit Co Pty Ltd_NoSite_12-08-2026_STM.pdf"])
        self.assertEqual((s.c["Filed"], s.c["Parked"]), (3, 1))

    def test_the_three_folders_exist_from_the_first_run(self):
        run(gateway())                                                              # an empty run
        self.assertEqual(sorted(os.listdir(self.path("Holding"))), ["_Account Activity Listings", "_Direct Debit Notices", "_Order Confirmations", "_Statements", "_Unassigned", "_Unrecognised"])

    def test_a_rehearsal_creates_no_folder(self):
        run(gateway(), shadow=True)
        self.assertFalse(os.path.exists(self.path("Holding")))

    def test_the_folder_name_is_a_setting(self):
        from openpyxl import load_workbook
        wb = load_workbook(settings.settings_file())
        ws = wb["DocTypes"]
        for r in range(2, ws.max_row + 1):
            if ws.cell(r, 1).value == "Statement":
                ws.cell(r, 3).value = "Statements"
        wb.save(settings.settings_file())
        wb.close()
        settings.reset()
        gw = gateway()
        gw.add(ONE, mail("Statement for September", ["statement.pdf"]))
        run(gw)
        self.assertEqual(self.files_under("Holding"), ["Statements/August 2026/ONE PEARL ENERGY/Fresh Fruit Co Pty Ltd_NoSite_12-08-2026_STM.pdf"])
        self.assertTrue(os.path.isdir(self.path("Holding", "Statements")))

    def test_the_date_format_in_names_is_a_setting(self):
        self.set_run(File_name_date_format="yyyymmdd")
        gw = gateway()
        gw.add(ONE, mail("Statement for September", ["statement.pdf"]))
        gw.add(ONE, mail("inv", ["corrupt.pdf"]))
        run(gw)
        held = self.files_under("Holding")
        self.assertIn("_Statements/August 2026/ONE PEARL ENERGY/Fresh Fruit Co Pty Ltd_NoSite_20260812_STM.pdf", held)
        self.assertIn("_Unrecognised/October 2026/20261003_corrupt.pdf", held)

    def test_two_different_statements_of_one_supplier_are_both_kept(self):
        gw = gateway()
        gw.add(ONE, mail("Statement for September", ["statement.pdf"], hours_ago=3))
        gw.add(ONE, mail("Statement for August", ["statement_rarity.pdf"], hours_ago=4, sender_name="Rarity Wholesale",
                         sender_email="accounts@raritywholesale.com.au"))
        s = run(gw)
        self.assertEqual((s.c["Parked"], s.c["duplicates"]), (2, 0))

    def test_the_same_statement_twice_is_one_file(self):
        gw = gateway()
        gw.add(ONE, mail("Statement for September", ["statement.pdf"], hours_ago=3))
        gw.add(ONE, mail("Statement for September again", ["statement.pdf"], hours_ago=2))
        s = run(gw)
        self.assertEqual((s.c["Parked"], s.c["duplicates"]), (1, 1))

    def test_a_statement_can_still_be_filed_like_an_invoice_if_the_settings_say_so(self):
        from openpyxl import load_workbook
        wb = load_workbook(settings.settings_file())
        ws = wb["DocTypes"]
        for r in range(2, ws.max_row + 1):
            if ws.cell(r, 1).value == "Statement":
                ws.cell(r, 2).value, ws.cell(r, 3).value = "Company folder", None
        wb.save(settings.settings_file())
        wb.close()
        settings.reset()
        gw = gateway()
        gw.add(ONE, mail("Statement for September", ["statement.pdf"]))
        s = run(gw)
        self.assertEqual((s.c["Filed"], s.c["Parked"]), (1, 0))
        self.assertEqual(self.files_under("TestOutput"), ["AUGUST 2026 EMAIL INVOICES/ONE PEARL ENERGY/Fresh Fruit Co Pty Ltd_NoSite_12-08-2026_STM.pdf"])      # dated 12 Aug: within the 3-month window, its own month folder


class SubjectFilterTests(Isolated, unittest.TestCase):
    """A test inside a mailbox that also holds REAL e-mails must touch only the test e-mails."""

    def build(self):
        gw = gateway()
        real = [gw.add(ONE, mail("Tax Invoice 111", ["invoice_acme.pdf"], hours_ago=5)),
                gw.add(ONE, mail("Overdue notice from supplier", ["invoice_other.pdf"], hours_ago=6)),
                gw.add(ONE, mail("FW: Invoice 222", ["invoice_company_named.pdf"], hours_ago=7))]
        test = [gw.add(ONE, mail("BOTTEST Tax Invoice 778899", ["invoice_bidfood.pdf"], hours_ago=2)),
                gw.add(ONE, mail("BOTTEST Overdue invoice reminder", ["invoice_other.pdf"], hours_ago=3))]
        return gw, real, test

    def test_only_matching_emails_are_touched(self):
        gw, real, test = self.build()
        s = run(gw, subject_filter="bottest")
        self.assertEqual((s.c["emails scanned"], s.c["Filed"], s.c["blocked"]), (2, 1, 1))
        for m in real:                                            # the real e-mails are exactly as they were
            self.assertEqual((m.unread, m.flag, m.categories), (True, 0, []))
        self.assertEqual((test[0].unread, test[0].flag), (False, 1))
        self.assertEqual((test[1].unread, test[1].flag), (True, 2))
        self.assertTrue(any("TEST FILTER" in n for n in s.notes))

    def test_a_filtered_test_does_not_move_the_catch_up_window(self):
        gw, real, test = self.build()
        run(gw, subject_filter="BOTTEST")
        self.assertEqual(connect_db(self.home).execute("SELECT COUNT(*) FROM mailbox_runs").fetchone()[0], 0)
        s = run(gw)                                               # the normal run still sees (and handles) the real e-mails
        self.assertEqual((s.c["Filed"], s.c["blocked"]), (2, 1))
        self.assertFalse(real[0].unread)
        self.assertEqual(real[1].flag, 2)
        self.assertEqual((real[2].categories, real[2].flag), ([], 2))
        self.assertEqual(s.c["documents"], 2)                     # the two BOTTEST e-mails were not processed twice

    def test_rehearsal_with_the_filter_changes_nothing(self):
        gw, real, test = self.build()
        before = state(gw)
        s = run(gw, shadow=True, subject_filter="bottest")
        self.assertEqual(state(gw), before)
        self.assertEqual(s.c["emails scanned"], 2)


class ParallelAndCliTests(Isolated, unittest.TestCase):
    def test_parallel_reading_gives_the_same_result(self):
        gw, _ = standard_mailbox(self.home)
        s = run(gw, workers=2)
        self.assertEqual((s.c["Filed"], s.c["Parked"], s.c["Unrecognised"]), (5, 3, 1))
        self.assertEqual(len(self.files_under("TestOutput")), 5)

    def test_demo_command(self):
        self.addCleanup(__import__("shutil").rmtree, os.path.join(ROOT, "demo_run"), True)
        out = subprocess.run([sys.executable, os.path.join(ROOT, "run_bot.py"), "--demo", "--workers", "1"],
                             capture_output=True, text=True, timeout=180, cwd=ROOT)
        self.assertEqual(out.returncode, 0, out.stderr[-500:])
        self.assertIn("RUN SUMMARY", out.stdout)
        self.assertIn("WHAT HAPPENED TO EACH SAMPLE E-MAIL", out.stdout)

    def test_bad_dates_are_rejected_by_the_command_line(self):
        out = subprocess.run([sys.executable, os.path.join(ROOT, "run_bot.py"), "--from", "31/12/2026"],
                             capture_output=True, text=True, timeout=60, cwd=ROOT)
        self.assertNotEqual(out.returncode, 0)
        self.assertIn("YYYY-MM-DD", out.stderr)


if __name__ == "__main__":
    unittest.main()
