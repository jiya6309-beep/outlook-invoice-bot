"""5.15.0: the real LIVE folder (OFFICE INDIA\\EMAIL INVOICES) with HOLDINGS and Reports inside it, LIVE makes missing folders like
TEST, the Sold-to address before the bill-to (Bega), the amount payable wins (Asahi), 'Never read' really means never, and a new
mailbox choice 'Read every e-mail'."""
import os
import shutil
import tempfile
import unittest

from openpyxl import load_workbook

from tests.docs_v515 import asahi_invoice, bega_invoice
from tests.helpers import DATA, ONE, Isolated, connect_db, gateway, mail, run
from tests.test_site_layout import site_row

from invoicebot import autofind, companies, labelreader, mailboxes, pdftext, settings
from invoicebot.router import ConfigError, Router

GOPINATH = "Gopinath Investment Pty Ltd"
KELSO = site_row("Ampol Kelso", "19 Sydney Road, Kelso NSW 2795", company=GOPINATH)
JIMBOOMBA = site_row("Mobil Jimboomba", "665 Cusack Lane, Jimboomba QLD 4280", company=GOPINATH)
WYONG = site_row("Pearl Energy North Wyong", "1 Pacific Highway, North Wyong NSW 2259", company="Central Coast Energy Pty Ltd")
JIYA = "Jiya@pearlenergy.com.au"


class Docs(Isolated, unittest.TestCase):
    def setUp(self):
        super().setUp()
        self.tmp = tempfile.mkdtemp(prefix="v515_")
        self.set_run(Destination_pattern="{MONTH_FOLDER}\\{COMPANY}\\{SITE}")
        self.add_rows("Sites", [KELSO, JIMBOOMBA, WYONG])

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)
        super().tearDown()

    def make(self, builder, name, **kw):
        p = os.path.join(self.tmp, name)
        builder(p, **kw)
        return p

    def docs(self):
        return [dict(r) for r in connect_db(self.home).execute("SELECT * FROM docs ORDER BY rowid")]


# ------------------------------------------------------------------------------------------------ 1. LIVE folders
class LiveFolderTests(Isolated, unittest.TestCase):
    def live(self, **extra):
        self.root = self.path("OFFICE INDIA", "EMAIL INVOICES")
        os.makedirs(self.root)
        self.set_run(Destination_mode="live", Live_destination_root=self.root, Destination_pattern="{MONTH_FOLDER}\\{COMPANY}\\{SITE}",
                     **extra)

    def test_the_default_live_folder_is_the_real_email_invoices_folder(self):
        self.assertEqual(settings.run_value("Live destination root"),
                         r"C:\SHAREPOINT\pearlenergy.com.au\Patel Group - Documents\OFFICE INDIA\EMAIL INVOICES")
        self.assertTrue(settings.run_bool("Live mode creates missing folders like test mode", False))
        self.assertEqual((settings.run_value("Holding root"), settings.run_value("Reports root")), ("", ""))

    def test_holdings_and_reports_follow_the_live_folder(self):
        self.live()
        self.assertEqual(settings.holding_root(), os.path.join(self.root, "HOLDINGS"))
        self.assertEqual(settings.reports_root(), os.path.join(self.root, "Reports"))
        other = self.path("elsewhere")
        self.set_run(Holding_root=other, Reports_root=other + "R")                   # a typed path still wins
        self.assertEqual((settings.holding_root(), settings.reports_root()), (other, other + "R"))
        self.set_run(Destination_mode="test")                                       # TEST never writes there
        self.assertEqual(settings.holding_root(), self.path("Holding"))

    def test_a_live_run_builds_month_company_site_holdings_and_reports(self):
        self.live()
        self.add_rows("Sites", [site_row("Mobil Colyton", "88-90 Great Western Hwy, Colyton NSW 2760")])
        gw = gateway()
        gw.add(ONE, mail("Tax Invoice INV-1001", ["invoice_acme.pdf"]))
        gw.add(ONE, mail("Statement September", ["statement.pdf"], hours_ago=3))
        s = run(gw)
        files = self.files_under("OFFICE INDIA", "EMAIL INVOICES")
        self.assertIn("OCTOBER 2026 EMAIL INVOICES/ONE PEARL ENERGY/MOBIL COLYTON/Acme Foods Pty Ltd_INV-1001.pdf", files)
        self.assertIn("Reports/To_Fix.xlsx", files)
        self.assertTrue(any(f.startswith("HOLDINGS/") for f in files), files)       # the statement went to HOLDINGS
        self.assertFalse(os.path.exists(self.path("Holding")) and os.listdir(self.path("Holding")))
        self.assertEqual(s.c["Filed"], 1)

    def test_an_existing_folder_is_used_never_made_twice(self):
        self.live()
        os.makedirs(os.path.join(self.root, "HOLDINGS", "_Unassigned"))
        os.makedirs(os.path.join(self.root, "OCTOBER 2026 EMAIL INVOICES", "ONE PEARL ENERGY", "MOBIL  COLYTON"))   # two spaces
        self.add_rows("Sites", [site_row("Mobil Colyton", "88-90 Great Western Hwy, Colyton NSW 2760")])
        gw = gateway()
        gw.add(ONE, mail("Tax Invoice INV-1001", ["invoice_acme.pdf"]))
        run(gw)
        company = os.path.join(self.root, "OCTOBER 2026 EMAIL INVOICES", "ONE PEARL ENERGY")
        self.assertEqual(os.listdir(company), ["MOBIL  COLYTON"])
        self.assertEqual(sorted(os.listdir(self.root)), ["HOLDINGS", "OCTOBER 2026 EMAIL INVOICES", "Reports"])

    def test_a_missing_email_invoices_folder_stops_and_writes_nothing(self):
        missing = self.path("OFFICE INDIA", "EMAIL INVOICES")
        os.makedirs(self.path("OFFICE INDIA"))
        self.set_run(Destination_mode="live", Live_destination_root=missing)
        with self.assertRaises(ConfigError) as cm:
            Router().check_config()
        self.assertIn("SharePoint folder was not found", str(cm.exception))
        gw = gateway()
        em = gw.add(ONE, mail("Tax Invoice INV-1001", ["invoice_acme.pdf"]))
        run(gw)
        self.assertEqual(os.listdir(self.path("OFFICE INDIA")), [])                  # no fake local EMAIL INVOICES
        self.assertTrue(em.unread)

    def test_the_folder_is_found_in_c_sharepoint(self):
        sp = self.path("C_SHAREPOINT")
        want = os.path.join(sp, "pearlenergy.com.au", "Patel Group - Documents", "OFFICE INDIA", "EMAIL INVOICES")
        os.makedirs(want)
        old = autofind.EXTRA_ROOTS
        autofind.EXTRA_ROOTS = [sp]
        try:
            self.assertEqual(autofind.find(), [want])
            self.set_run(Destination_mode="live")                                    # the typed default does not exist here
            gw = gateway()
            gw.add(ONE, mail("Tax Invoice INV-1001", ["invoice_acme.pdf"]))
            s = run(gw)
            self.assertEqual(s.c["Filed"] + s.c["Unassigned"], 1)
            self.assertTrue(os.path.isdir(os.path.join(want, "Reports")))
        finally:
            autofind.EXTRA_ROOTS = old


# ------------------------------------------------------------------------------------------------ 2. the settings file
class UpgradeTests(Isolated, unittest.TestCase):
    def test_the_payments_teams_file_is_upgraded_once(self):
        shutil.copyfile(os.path.join(DATA, "settings_payments_team_5_13.xlsx"), settings.settings_file())
        settings.reset()
        added = settings.upgrade_workbook()
        settings.reset()
        self.assertEqual(settings.run_value("Settings version"), settings.SETTINGS_VERSION)
        self.assertEqual(settings.run_value("Live destination root"), settings.LIVE_ROOT)
        self.assertEqual((settings.run_value("Holding root"), settings.run_value("Reports root")), ("", ""))
        self.assertEqual(settings.run_value("Auto-find folder name"), "EMAIL INVOICES")
        self.assertTrue(settings.run_bool("Live mode creates missing folders like test mode", False))
        star = next(r for r in settings.rows("Labels") if r["Supplier name contains"] == "*")
        self.assertNotIn("advice no", star["Invoice date label"].lower())
        self.assertTrue(star["Invoice number label"].lower().endswith("advice no"))
        self.assertTrue(star["Total (incl. GST) label"].lower().startswith("total payable|total amount payable|amount payable"))
        self.assertTrue(star["Total (incl. GST) label"].lower().endswith("|total"))
        self.assertIn("total incl taxes", star["Total (incl. GST) label"].lower())
        self.assertTrue(any(r["Supplier name contains"] == "Asahi" for r in settings.rows("Labels")))
        bega = next(r for r in settings.rows("Labels") if r["Supplier name contains"].lower() == "bega")
        self.assertEqual(bega["Site / deliver-to label"].lower(), "sold to")                 # kept as the payments team wrote it
        kelso = next(r for r in settings.rows("Sites") if r["Site name"] == "Ampol Kelso")
        self.assertIn("19 sydney", kelso["Address fragments"])
        self.assertIn("17 great western", kelso["Address fragments"])
        self.assertFalse(any(r["Company name"].startswith("TEST - ") for r in settings.rows("Companies")))
        self.assertEqual(companies.special_for(JIYA)["role"], "read all")
        singleton = next(r for r in settings.rows("Sites") if r["Site name"] == "Mobil Singleton")
        self.assertIn("Mobil Dunolly", singleton["Aliases"])                                 # not touched
        self.assertTrue(added)
        self.assertEqual(settings.upgrade_workbook(), [])                                    # only once


# ------------------------------------------------------------------------------------------------ 3. Sold to (Bega)
class SoldToTests(Docs):
    def test_bega_goes_to_the_sold_to_site_not_the_payer(self):
        gw = gateway()
        gw.add(ONE, mail("Bega invoice 577967880", [self.make(bega_invoice, "bega.pdf")]))
        run(gw)
        d = self.docs()[0]
        self.assertEqual((d["site"], d["company_folder"]), ("Ampol Kelso", "GOPINATH INVESTMENT"))
        self.assertIn("sold-to", d["site_source"])
        self.assertEqual(d["amount"], 102.80)
        self.assertIn("OCTOBER 2026 EMAIL INVOICES/GOPINATH INVESTMENT/AMPOL KELSO/", "/".join(self.files_under("TestOutput")))

    def test_a_ship_to_still_wins_over_the_sold_to(self):
        from invoicebot import sites
        m = sites.match("", blocks=["[SOLD] Ampol Kelso | 17 Great Western Hwy | KELSO NSW 2795",
                                    "[SHIP] Mobil Jimboomba | 665 Cusack Lane | JIMBOOMBA QLD 4280"], supplier="Bega")
        self.assertEqual(m["site"], "Mobil Jimboomba")
        m = sites.match("", blocks=["Mobil Jimboomba | PO Box 7501", "[SOLD] Ampol Kelso | 17 Great Western Hwy"], supplier="Bega")
        self.assertEqual((m["site"], m["source"]), ("Ampol Kelso", sites.SOLD_SOURCE))
        doc = pdftext.load(self.make(bega_invoice, "b2.pdf"))
        self.assertTrue(any(b.startswith("[SOLD] ") for b in labelreader._read_blocks(doc, settings.labels_for("", "Bega")["site"])))

    def test_an_unknown_sold_to_never_falls_back_to_a_guess_between_two_sites(self):
        p = self.make(bega_invoice, "b3.pdf", sold_lines=["Account: 1", "Somewhere Else", "1 Nowhere St", "TOWN NSW 2000"])
        gw = gateway()
        gw.add(ONE, mail("Bega invoice", [p]))
        s = run(gw)
        d = self.docs()[0]
        self.assertEqual((d["site"], s.c["Unassigned"]), ("UNASSIGNED", 1))                  # the payer (Mobil Jimboomba) is never the site
        self.assertIn("SOLD-TO NOT IN SITE LIST", d["reason"] + d["site_source"])
        self.assertIn("Somewhere Else", d["ship_to"])                                       # shown in To_Fix, learned when staff choose


# ------------------------------------------------------------------------------------------------ 4. the amount payable
class PayableTests(Docs):
    def amounts(self, path, supplier="Asahi Beverages"):
        r = labelreader.read(pdftext.load(path), settings.labels_for(supplier, supplier))
        return r["amount"], r["notes"]

    def test_asahi_total_payable_is_the_amount(self):
        gw = gateway()
        gw.add(ONE, mail("Asahi invoice", [self.make(asahi_invoice, "asahi.pdf")]))
        run(gw)
        d = self.docs()[0]
        self.assertEqual(d["amount"], 1336.44)                                              # not 1,125.56 (net total ex.CDL)
        self.assertEqual(d["gst"], 116.80)

    def test_asahi_without_readable_headings_the_maths_decides(self):
        amount, notes = self.amounts(self.make(asahi_invoice, "a2.pdf", readable_headings=False))
        self.assertEqual(amount, 1336.44)
        self.assertFalse(any("please check" in n for n in notes))

    def test_the_row_rule(self):
        self.assertEqual(labelreader._row_total([1125.56, 94.08, 1219.64, 116.80, 1336.44]), (1336.44, True))
        self.assertEqual(labelreader._row_total([100.00, 10.00, 110.00]), (110.00, True))
        self.assertEqual(labelreader._row_total([-100.00, -10.00, -110.00]), (-110.00, True))      # credit notes
        self.assertEqual(labelreader._row_total([50.00, 80.00]), (80.00, False))                   # nothing adds up: largest

    def test_payable_words_come_first_and_a_zero_balance_never_wins(self):
        words = settings.labels_for("Anyone", "Anyone")["total"]
        self.assertEqual(words[:3], ["total payable", "total amount payable", "amount payable"])
        self.assertLess(words.index("total inc gst"), words.index("balance due"))
        self.assertLess(words.index("total incl taxes"), words.index("invoice total"))
        self.assertEqual(words[-1], "total")
        self.add_rows("Labels", [["bega", "sold to", "invoice no", "", "total gst included", "", "total due"]])
        self.assertEqual(settings.labels_for("Bega", "Bega")["total"][0], "total due")       # a supplier's own row still first


# ------------------------------------------------------------------------------------------------ 5. Never read means never
class NeverReadTests(Isolated, unittest.TestCase):
    def outlook(self):
        gw = gateway()
        self.scanned = []
        real = gw.scan

        def scan(store, *a, **kw):
            self.scanned.append(store)
            return real(store, *a, **kw)
        gw.scan = scan
        return gw

    def test_one_company_read_and_twelve_never_reads_only_that_one(self):
        gw = self.outlook()
        boxes = [b for b in gw.list_mailboxes()]
        items = [{"outlook": b, "action": "company", "company": "PEARL ENERGY NO: 1 PTY LTD", "active": True} if b == ONE
                 else {"outlook": b, "action": "special", "role": "never"} for b in boxes]
        mailboxes.apply(items)
        settings.reset()
        for b in boxes:
            gw.add(b, mail("Tax Invoice 778899", ["invoice_bidfood.pdf"]))
        s = run(gw)
        self.assertEqual(set(self.scanned), {ONE})
        self.assertEqual(s.c["documents"], 1)
        self.assertTrue(any("Two@pearlenergy.com.au is set to Never read" in l or "TWO" in l.upper() for l in s.mailbox_lines))

    def test_the_other_name_of_a_never_mailbox_cannot_reach_it(self):
        gw = self.outlook()
        self.add_rows("SpecialMailboxes", [["Two@pearlenergy.com.au", "Never", "", "Yes", ""]])
        # the old 5.14.0 result: the company row still has its second name 'Two Pearl Energy'
        wb = load_workbook(settings.settings_file())
        ws = wb["Companies"]
        for r in range(2, ws.max_row + 1):
            if "Two@" in str(ws.cell(r, 4).value):
                ws.cell(r, 4).value = "Two Pearl Energy"
        wb.save(settings.settings_file())
        settings.reset()
        gw.add("Two@pearlenergy.com.au", mail("Tax Invoice 778899", ["invoice_bidfood.pdf"]))
        run(gw)
        self.assertNotIn("Two@pearlenergy.com.au", self.scanned)

    def test_a_company_without_any_mailbox_name_is_skipped_not_a_crash(self):
        wb = load_workbook(settings.settings_file())
        ws = wb["Companies"]
        for r in range(2, ws.max_row + 1):
            if "Three@" in str(ws.cell(r, 4).value):
                ws.cell(r, 4).value = None
        wb.save(settings.settings_file())
        settings.reset()
        gw = self.outlook()
        gw.add(ONE, mail("Tax Invoice 778899", ["invoice_bidfood.pdf"]))
        s = run(gw)
        self.assertEqual(s.c["documents"], 1)

    def test_setting_a_test_mailbox_back_to_never_removes_its_test_row(self):
        mailboxes.apply([{"outlook": JIYA, "action": "test"}])
        settings.reset()
        self.assertTrue(any(r["Company name"] == "TEST - " + JIYA for r in settings.rows("Companies")))
        mailboxes.apply([{"outlook": JIYA, "action": "special", "role": "never"}])
        settings.reset()
        self.assertFalse(any(r["Company name"].startswith("TEST - ") for r in settings.rows("Companies")))


# ------------------------------------------------------------------------------------------------ 6. Read every e-mail
class ReadAllTests(Docs):
    def test_every_e_mail_is_read_and_filed_like_a_company_mailbox(self):
        msgs = mailboxes.apply([{"outlook": JIYA, "action": "special", "role": "read all"}])
        settings.reset()
        self.assertIn("read every e-mail", msgs[0])
        gw = gateway(JIYA)
        gw.add(JIYA, mail("Bega invoice", [self.make(bega_invoice, "bega.pdf")]))
        gw.add(JIYA, mail("Asahi invoice", [self.make(asahi_invoice, "asahi.pdf")], hours_ago=3), folder="Suppliers")   # a sub-folder
        s = run(gw)
        self.assertEqual(s.c["Filed"], 2)
        files = "/".join(self.files_under("TestOutput"))
        self.assertIn("GOPINATH INVESTMENT/AMPOL KELSO/", files)
        self.assertIn("/PEARL ENERGY NORTH WYONG/", files)
        self.assertNotIn("TEST", files)

    def test_the_mailbox_is_never_a_company_clue(self):
        mailboxes.apply([{"outlook": JIYA, "action": "special", "role": "read all"}])
        settings.reset()
        row, source, _ = companies.resolve("", "", JIYA)
        self.assertIsNone(row)

    def test_any_number_of_mailboxes_can_read_everything(self):
        msgs = mailboxes.apply([{"outlook": JIYA, "action": "special", "role": "read all"},
                                {"outlook": "Other@pearlenergy.com.au", "action": "special", "role": "read all"}])
        settings.reset()
        self.assertEqual(len(msgs), 2)
        self.assertEqual({r["role"] for r in companies.special_rows() if r["role"] == "read all"}, {"read all"})
        self.assertEqual(sum(1 for r in companies.special_rows() if r["role"] == "read all"), 2)


if __name__ == "__main__":
    unittest.main()
