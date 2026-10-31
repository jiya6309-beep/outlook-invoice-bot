"""5.11.0: Fleet Card and M STATEMENTS mailboxes, special mailboxes, statements by company, Direct Debit Notices,
Account Activity Listings, ship-to learning / SharePoint folders, LIVE folder rules, the Record run marker, the A/C Excel,
the To_Fix additions and the one-time settings upgrade."""
import datetime
import os
import tempfile
import unittest

from openpyxl import Workbook, load_workbook

from tests.helpers import CAREFUL_LIVE, DATA, ONE, Isolated, connect_db, gateway, mail, run
from tests.test_site_layout import site_row
from tests import make_data

from invoicebot import companies, db, exports, settings, sites, suppliers

FLEET_PDF = os.path.join(DATA, "fleetcard_ampol_farleigh_01_sep_26.pdf")
PRESTON_PDF = "unilever_billto_denham_court_shipto_preston.pdf"
FUEL = "P I Fuel Invoice"
MSTAT = "mstatement@pearlenergy.com.au"
FLEETBOX = MSTAT                       # 5.12: Fleet Card statements arrive in mstatement@ (P I Fuel Invoice = fuel invoices)

FARLEIGH = site_row("Ampol Farleigh", "971 Bruce Hwy, Farleigh QLD 4741", company="Pearl Energy No 2 Pty Ltd")
DENHAM = site_row("Ampol Denham court", "194 Campbelltown Road, Denham Court NSW 2565", company="Gopinath Investment Pty Ltd")
COLYTON = site_row("Mobil Colyton", "88-90 Great Western Hwy, Colyton NSW 2760", "MOBIL COLYTON", company="Colyton Fuels Pty Ltd")
FILLER = [site_row(f"Ampol Town{i}", company="Gopinath Investment Pty Ltd") for i in range(3)]


class Base(Isolated, unittest.TestCase):
    def setUp(self):
        super().setUp()
        self.tmp = tempfile.mkdtemp(prefix="v511_")
        self.fleet_root = self.path("TestOutput", "FLEETCARD STATEMENTS")     # TEST mode: never SharePoint
        os.makedirs(self.fleet_root, exist_ok=True)
        self.add_rows("Sites", [FARLEIGH, DENHAM, COLYTON] + FILLER)

    def set_special(self, mailbox, role, dest):
        wb = load_workbook(settings.settings_file())
        ws = wb["SpecialMailboxes"]
        for r in range(2, ws.max_row + 1):
            if mailbox.lower() in str(ws.cell(r, 1).value).lower():
                ws.cell(r, 3).value = dest
        wb.save(settings.settings_file())
        wb.close()
        settings.reset()

    def gw(self):
        g = gateway()
        for m in (FUEL, MSTAT, "Security", "office@pearlenergy.com.au"):
            g.add_mailbox(m)
        return g

    def pdf(self, name, title, lines, supplier="Zed Supplies Pty Ltd"):
        path = os.path.join(self.tmp, name)
        make_data.simple(path, title, lines, supplier=supplier)
        return path

    def docs(self):
        return [dict(r) for r in connect_db(self.home).execute("SELECT * FROM docs ORDER BY id")]


class FleetCardTests(Base):
    def test_a_fleet_card_statement_goes_to_its_site_folder_named_by_report_issue_date(self):
        g = self.gw()
        em = g.add(FLEETBOX, mail("Fleetcard Merchant Payment Advice", [FLEET_PDF]))
        s = run(g)
        self.assertEqual(s.c["Filed"], 1)
        self.assertEqual(self.files_under("TestOutput", "FLEETCARD STATEMENTS"), ["AMPOL FARLEIGH/01-09-2026_STM.pdf"])
        d = self.docs()[0]
        self.assertEqual((d["doc_type"], d["supplier"], d["site"]), ("Fleet Card Statement", "Business Fuel Cards", "Ampol Farleigh"))
        self.assertFalse(em.unread)                                    # read + tick as usual
        learned = db.learned_all(connect_db(self.home))
        self.assertEqual([(x["key"], x["site"]) for x in learned], [("merchant:248858", "Ampol Farleigh")])

    def test_an_existing_site_folder_is_used_and_a_second_copy_is_not_made(self):
        os.makedirs(os.path.join(self.fleet_root, "Ampol  Farleigh"))       # other spelling, as staff typed it
        g = self.gw()
        g.add(FLEETBOX, mail("Fleetcard Merchant Payment Advice", [FLEET_PDF]))
        run(g)
        self.assertEqual(os.listdir(self.fleet_root), ["Ampol  Farleigh"])

    def test_in_any_other_mailbox_it_still_goes_to_the_fleet_folder(self):
        g = self.gw()                                                     # 5.19.0 (was: 'Other' outside the Fleet card mailbox)
        g.add(ONE, mail("Fleetcard Merchant Payment Advice", [FLEET_PDF]))
        run(g)
        self.assertEqual(self.docs()[0]["doc_type"], "Fleet Card Statement")
        self.assertNotIn("EMAIL INVOICES", self.docs()[0]["dest_dir"])

    def test_unknown_site_waits_then_the_learned_merchant_number_files_it(self):
        conn = connect_db(self.home)
        self.add_rows("Sites", [])                                     # make sure the sheet is loaded
        wb = load_workbook(settings.settings_file())                   # remove Ampol Farleigh: the site cannot be read from the name
        ws = wb["Sites"]
        for r in range(ws.max_row, 1, -1):
            if ws.cell(r, 1).value == "Ampol Farleigh":
                ws.delete_rows(r)
        ws.append(site_row("Pearl Energy Farleigh North", company="Pearl Energy No 2 Pty Ltd"))
        wb.save(settings.settings_file())
        settings.reset()
        g = self.gw()
        g.add(FLEETBOX, mail("Fleetcard Merchant Payment Advice", [FLEET_PDF]))
        s = run(g)
        self.assertEqual(s.c["Unassigned"], 1)
        self.assertIn("01-09-2026_STM.pdf", self.files_under("Holding", "_Unassigned"))
        self.assertEqual(self.docs()[0]["ship_to"], "merchant:248858")
        db.learn(conn, "ship-to", settings.supplier_key("Business Fuel Cards"), "merchant:248858", "Pearl Energy Farleigh North")
        conn.commit()
        s2 = run(self.gw())
        self.assertEqual(s2.c["released"], 1)
        self.assertEqual(self.files_under("TestOutput", "FLEETCARD STATEMENTS"), ["PEARL ENERGY FARLEIGH NORTH/01-09-2026_STM.pdf"])


class SpecialMailboxTests(Base):
    def test_fuel_invoices_in_the_inbox_and_its_subfolders(self):
        g = self.gw()
        g.add(FUEL, mail("Your invoice", ["fuel_INV-115316.pdf"]))                         # no keyword, in the Inbox
        g.add(FUEL, mail("Fuel", ["fuel_INV-114870.pdf"]), folder="Colyton")             # in a sub-folder
        g.add(FUEL, mail("Fuel", ["fuel_INV-115270.pdf"]), folder="P1\\Old")             # deeper
        s = run(g)
        self.assertEqual(s.c["Filed"], 3)
        self.assertEqual(sorted(self.files_under("TestOutput", "FUEL INVOICES")),
                         ["CENTRAL COAST ENERGY 1/INV-115270_PEARL ENERGY NORTH WYONG.pdf",
                          "HIGHLAND ENERGY 1/INV-115316_PEARL ENERGY TARAGO.pdf",
                          "TWEED HEADS ENERGY NO.1/INV-114870_PEARL ENERGY TWEED HEADS.pdf"])
        rows = list(load_workbook(self.path("Reports", [f for f in os.listdir(self.path("Reports")) if f.startswith("PI_Invoice_list_")][0])).active.iter_rows(values_only=True))
        self.assertEqual(rows[0][:5], ("Download Date", "Contact", "Invoice Number", "Reference", "Invoice Date"))
        by_no = {r[2]: r for r in rows[1:]}
        self.assertEqual(by_no["INV-115270"][1], "Central Coast Energy 1 Pty Ltd - PEARL ENERGY NORTH WYONG")
        self.assertEqual(by_no["INV-115270"][3], "C1120RFI00143558-OZ1104-SPLIT-MORISSET")
        self.assertEqual(by_no["INV-115270"][5:8], (None, None, None))                     # XERO / ABA / QUICK FUEL: staff fill them

    def test_the_fuel_list_keeps_staff_ticks_and_never_lists_twice(self):
        g = self.gw()
        g.add(FUEL, mail("Fuel", ["fuel_INV-115316.pdf"]))
        run(g)
        name = [f for f in os.listdir(self.path("Reports")) if f.startswith("PI_Invoice_list_")][0]
        wb = load_workbook(self.path("Reports", name))
        wb.active.cell(2, 6).value = "YES"                                                 # staff tick XERO Bill
        wb.save(self.path("Reports", name))
        g2 = self.gw()
        g2.add(FUEL, mail("Fuel", ["fuel_INV-115316.pdf"], hours_ago=0.5))                # the same invoice again
        g2.add(FUEL, mail("Fuel", ["fuel_INV-114870.pdf"], hours_ago=0.5))
        run(g2)
        rows = list(load_workbook(self.path("Reports", name)).active.iter_rows(values_only=True))
        self.assertEqual([r[2] for r in rows[1:]], ["INV-115316", "INV-114870"])
        self.assertEqual(rows[1][5], "YES")

    def test_fleet_card_record_has_the_agreed_columns(self):
        g = self.gw()
        g.add(FLEETBOX, mail("Fleetcard Merchant Payment Advice", [FLEET_PDF]))
        run(g)
        name = [f for f in os.listdir(self.path("Reports")) if f.startswith("FleetCard_Record_")][0]
        rows = list(load_workbook(self.path("Reports", name))["Record"].iter_rows(values_only=True))
        self.assertEqual(rows[0], ("Merchant No", "Report Issue Date", "Total payment", "Site", "Company", "Status", "File name"))
        self.assertEqual((rows[1][0], rows[1][2], rows[1][3], rows[1][6]), ("248858", 15611.06, "Ampol Farleigh", "01-09-2026_STM.pdf"))

    def test_on_hold_and_never_mailboxes_are_not_read(self):
        g = self.gw()
        e1 = g.add("Security", mail("Invoice 1", ["invoice_bidfood.pdf"]))
        e2 = g.add("office@pearlenergy.com.au", mail("Invoice 2", ["invoice_bidfood.pdf"]))
        s = run(g)
        self.assertEqual(s.c["documents"], 0)                 # Security is read only by the security run, office@ never
        self.assertTrue(e1.unread and e2.unread)

    def test_a_companies_row_with_office_is_still_never_read(self):
        wb = load_workbook(settings.settings_file())
        ws = wb["Companies"]
        ws.cell(2, 4).value = "office@pearlenergy.com.au"
        wb.save(settings.settings_file())
        settings.reset()
        self.assertNotIn("office@pearlenergy.com.au", [companies.mailbox_names(r)[0] for r in companies.reading_list()])

    def test_a_special_mailbox_is_never_a_company_clue(self):
        row, source, _w = companies.resolve("", "", MSTAT)
        self.assertIsNone(row)

    def test_the_mailbox_finder_marks_special_mailboxes(self):
        from invoicebot import mailboxes
        out = {x["outlook"]: x for x in mailboxes.suggest(["P I Fuel Invoice", "Security", "office@pearlenergy.com.au", "One@pearlenergy.com.au"])}
        self.assertEqual(out["P I Fuel Invoice"]["special"], "fuel invoices")
        self.assertEqual(out["Security"]["special"], "security")
        self.assertEqual(out["office@pearlenergy.com.au"]["special"], "never")
        self.assertNotIn("special", out["One@pearlenergy.com.au"])


class HoldingLayoutTests(Base):
    def test_statements_go_to_month_then_company(self):
        p = self.pdf("stm.pdf", "STATEMENT", ["Statement Date: 12/08/2026", "Mobil Colyton", "88-90 Great Western Hwy Colyton NSW 2760"])
        g = self.gw()
        g.add(ONE, mail("Statement", [p]))
        run(g)
        self.assertEqual(self.files_under("Holding", "_Statements"),
                         ["August 2026/COLYTON FUELS/Zed Supplies Pty Ltd_Mobil Colyton_12-08-2026_STM.pdf"])

    def test_direct_debit_notice_by_company_and_never_named_dear_valued(self):
        p = self.pdf("ddn.pdf", "DIRECT DEBIT NOTICE", ["15.10.2026", "PEARL ENERGY NO.1 FUT PAYER", "Dear Valued Customer",
                                                       "Account Number: 1206521"], supplier="TIP TOP Bakeries")
        g = self.gw()
        g.add(ONE, mail("Direct Debit Notice", [p]))
        run(g)
        files = self.files_under("Holding", "_Direct Debit Notices")
        self.assertEqual(len(files), 1)
        self.assertTrue(files[0].startswith("ONE PEARL ENERGY/Tip Top Bakeries_"), files)
        self.assertTrue(files[0].endswith("_DDN.pdf"))

    def test_account_activity_listing_by_site(self):
        p = self.pdf("aal.pdf", "Account Activity Listing", ["Customer account 33427741", "Document date 12/10/2026",
                                                            "MOBIL COLYTON", "88-90 Great Western Hwy Colyton NSW 2760"],
                     supplier="Campbells Wholesale Pty Ltd")
        g = self.gw()
        g.add(ONE, mail("Account Activity Listing", [p]))
        run(g)
        files = self.files_under("Holding", "_Account Activity Listings")
        self.assertEqual(len(files), 1)
        self.assertTrue(files[0].startswith("MOBIL COLYTON/") and files[0].endswith("_12-10-2026_AAL.pdf"), files)

    def test_a_greeting_or_label_is_never_a_supplier(self):
        for bad in ("Dear Valued", "Billing To", "Bill To Pearl Energy", "Deliver To Mobil Ballina"):
            self.assertFalse(suppliers.plausible(bad), bad)
        for good in ("Delivery Express Pty Ltd", "Customer Foods Pty Ltd", "Business Fuel Cards Pty Ltd"):
            self.assertTrue(suppliers.plausible(good), good)


class ShipToTests(Base):
    def setUp(self):
        super().setUp()
        self.root = self.path("MAIL INVOICES")
        os.makedirs(self.root)
        self.set_run(Destination_pattern="{MONTH_FOLDER}\\{COMPANY}\\{SITE}", Destination_mode="live", Live_destination_root=self.root,
                     **CAREFUL_LIVE)

    def deliver(self, hours_ago=2):
        g = self.gw()
        g.add(ONE, mail("Unilever invoice", [PRESTON_PDF], hours_ago=hours_ago))
        return run(g)

    def test_live_waits_when_the_site_folder_is_missing_and_creates_month_and_company_only(self):
        self.add_rows("Sites", [site_row("Pearl Energy Preston", "1B Bell Street, Preston VIC 3072", company="Pearl Energy No 5 Pty Ltd")])
        s = self.deliver()
        self.assertEqual(s.c["Unassigned"], 1)
        five = os.path.join(self.root, "OCTOBER 2026 EMAIL INVOICES", "FIVE PEARL ENERGY")
        self.assertTrue(os.path.isdir(five))                                       # month + company created
        self.assertEqual(os.listdir(five), [])                                    # the site folder is NOT
        self.assertIn("site folder does not exist", self.docs()[0]["reason"])

    def test_preston_is_filed_once_staff_create_its_folder(self):
        s = self.deliver()                                                        # Preston is not in the Sites sheet
        self.assertEqual(s.c["Unassigned"], 1)
        target = os.path.join(self.root, "OCTOBER 2026 EMAIL INVOICES", "FIVE PEARL ENERGY", "PEARL ENERGY PRESTON")
        os.makedirs(target)
        s2 = run(self.gw())
        self.assertEqual(s2.c["released"], 1)
        self.assertEqual(os.listdir(target), ["Unilever_3506013060.pdf"])
        d = self.docs()[0]
        self.assertEqual((d["status"], d["company_folder"]), ("Filed", "FIVE PEARL ENERGY"))
        self.assertIn("SharePoint folder", d["site_source"])

    def test_a_site_chosen_in_to_fix_is_learned_for_the_next_invoice(self):
        os.makedirs(os.path.join(self.root, "OCTOBER 2026 EMAIL INVOICES", "GOPINATH INVESTMENT", "AMPOL TOWN0"))
        self.deliver()
        path = os.path.join(self.root, "Reports", "To_Fix.xlsx")          # 5.15.0: Reports inside the live folder
        wb = load_workbook(path)
        ws = wb["To fix"]
        head = [c.value for c in ws[1]]
        self.assertIn("PEARL ENERGY PRESTON", str(ws.cell(2, head.index("Deliver-to as printed") + 1).value))
        ws.cell(2, head.index("Site folder (TYPE HERE)") + 1).value = "AMPOL TOWN0"
        wb.save(path)
        wb.close()
        s = run(self.gw())
        self.assertEqual(s.c["sites learned"], 1)
        conn = connect_db(self.home)
        self.assertIn("num:64228817", [x["key"] for x in db.learned_all(conn)])
        # the same deliver-to again (a new invoice number): filed by itself
        conn.execute("DELETE FROM docs")
        conn.execute("DELETE FROM emails")
        conn.commit()
        s2 = self.deliver(hours_ago=0.5)                                          # inside the next catch-up window
        self.assertEqual(s2.c["Filed"], 1)
        self.assertIn("learned from a staff fix", self.docs()[-1]["site_source"])

    def test_a_learned_site_can_be_forgotten_in_to_fix(self):
        conn = connect_db(self.home)
        db.learn(conn, "ship-to", "unilever", "num:64228817", "Ampol Town0", "GOPINATH INVESTMENT")
        conn.commit()
        run(self.gw())                                                            # writes To_Fix with the Learned sites sheet
        path = os.path.join(self.root, "Reports", "To_Fix.xlsx")          # 5.15.0: Reports inside the live folder
        wb = load_workbook(path)
        ws = wb["Learned sites"]
        self.assertEqual(ws.cell(2, 5).value, "Ampol Town0")
        ws.cell(2, 9).value = "YES"
        wb.save(path)
        wb.close()
        run(self.gw())
        self.assertEqual(db.learned_all(connect_db(self.home)), [])


class ExcelTests(Base):
    def run_three(self):
        g = self.gw()
        g.add(ONE, mail("Tax Invoice 778899", ["invoice_bidfood.pdf"], hours_ago=5))
        g.add(ONE, mail("Credit Note CN-9001", ["credit_note.pdf"], hours_ago=4))
        g.add(ONE, mail("POD D-77120", ["pod.pdf"], hours_ago=3))
        return run(g)

    def test_record_has_the_download_time_and_marks_the_last_e_mail_of_the_run(self):
        self.run_three()
        ws = load_workbook(self.path("Reports", "Record_2026-10.xlsx"))["Record"]
        rows = list(ws.iter_rows(values_only=True))
        head = rows[0]
        self.assertIn("Downloaded by bot", head)
        marks = [r for r in rows[1:] if r[head.index("Run marker")]]
        self.assertEqual(len(marks), 1)
        self.assertEqual(marks[0][head.index("Document type")], "POD")             # the newest e-mail
        self.assertIn("LAST E-MAIL OF RUN", marks[0][head.index("Run marker")])
        self.assertRegex(rows[1][head.index("Downloaded by bot")], r"^\d\d-\d\d-\d{4} \d\d:\d\d$")

    def test_ac_excel_has_the_company_format(self):
        self.run_three()
        path, n = exports.write_ac(connect_db(self.home), "2026-01-01", "2026-12-31")
        self.assertTrue(os.path.basename(path).startswith("AC_2026-01-01_2026-12-31"))
        rows = list(load_workbook(path)["AC"].iter_rows(values_only=True))
        self.assertEqual(rows[0], tuple(exports.AC_HEADERS))
        by_number = {(r[4] or r[6]): r for r in rows[1:]}
        pod = by_number["D-77120"]
        self.assertIsNotNone(pod[3])
        self.assertIsNone(pod[5])
        self.assertIsNone(pod[7])                                                  # no amount for a docket
        cn = by_number["CN-9001"]
        self.assertLess(cn[7], 0)                                                  # credit notes negative
        self.assertIn("HYPERLINK", rows[1][8])
        self.assertEqual(n, 3)

    def test_the_command_line_makes_the_ac_excel(self):
        from invoicebot.panel import Controller
        args = Controller(self.home).args_for("ac", {"from": "2026-10-01", "to": "2026-10-31"})
        self.assertIn("--ac", args)


class ReportsRootTests(Isolated, unittest.TestCase):
    def test_reports_go_to_the_reports_root(self):
        target = self.path("SP", "MAILS", "Reports")
        os.makedirs(os.path.dirname(target))
        os.makedirs(self.path("SP", "MAIL INVOICES"))
        self.set_run(Reports_root=target, Destination_mode="live", Live_destination_root=self.path("SP", "MAIL INVOICES"))
        self.assertEqual(exports.reports_dir(), target)
        g = gateway()
        g.add(ONE, mail("Tax Invoice 778899", ["invoice_bidfood.pdf"]))
        run(g)
        self.assertIn("To_Fix.xlsx", os.listdir(target))

    def test_an_unsynced_sharepoint_folder_stops_the_run(self):
        os.makedirs(self.path("SP", "MAIL INVOICES"))
        self.set_run(Holding_root=self.path("NOT SYNCED", "OFFICE INDIA", "MAILS"), Destination_mode="live",
                     Live_destination_root=self.path("SP", "MAIL INVOICES"))
        s = run(gateway())
        self.assertEqual(s.stop_reason, "setup incomplete")
        self.assertFalse(os.path.exists(self.path("NOT SYNCED")))

    def test_test_mode_never_writes_holding_or_reports_into_sharepoint(self):
        sp = self.path("SP", "OFFICE INDIA", "MAILS")
        os.makedirs(sp)
        self.set_run(Holding_root=sp, Reports_root=os.path.join(sp, "Reports"))       # LIVE paths filled in, but TEST mode
        g = gateway()
        g.add(ONE, mail("Tax Invoice 778899", ["invoice_bidfood.pdf"]))
        run(g)
        self.assertEqual(os.listdir(sp), [])                                           # nothing went to "SharePoint"
        self.assertIn("To_Fix.xlsx", os.listdir(self.path("Reports")))
        self.assertTrue(os.path.isdir(self.path("Holding")))

    def test_a_onedrive_conflict_copy_of_to_fix_is_reported(self):
        os.makedirs(self.path("Reports"), exist_ok=True)
        open(self.path("Reports", "To_Fix-OFFICEPC.xlsx"), "w").close()
        s = run(gateway())
        self.assertTrue(any("conflict copy" in w for w in s.warnings))


class UpgradeTests(Isolated, unittest.TestCase):
    def test_an_older_settings_file_is_upgraded_once(self):
        path = settings.settings_file()
        wb = load_workbook(path)
        ws = wb["Run"]
        for r in range(ws.max_row, 1, -1):
            key = str(ws.cell(r, 1).value or "")
            if key == "Settings version":
                ws.delete_rows(r)
            elif key == "Live mode creates missing folders like test mode":
                ws.cell(r, 2).value = "Yes"                                        # as 5.10.0 shipped it
        kw = wb["Keywords"]
        for r in range(2, kw.max_row + 1):
            if kw.cell(r, 1).value == "direct debit notice":
                kw.cell(r, 2).value = "Other"
        del wb["SpecialMailboxes"]
        sites_ws = wb["Sites"]
        sites_ws.append(["BP Orange North", "Pearl Energy No 3 Pty Ltd", "", "5 hanharan", "", "", "Yes", "", "", ""])
        wb.save(path)
        settings.reset()
        added = settings.upgrade_workbook()
        self.assertTrue(any("Live mode creates" in a for a in added))
        settings.reset()
        self.assertTrue(settings.run_bool("Live mode creates missing folders like test mode", False))   # 5.11.0 No, then 5.15.0 Yes
        self.assertTrue(settings.run_bool("Create company folder in live mode", False))
        self.assertEqual(settings.run_value("Settings version"), settings.SETTINGS_VERSION)
        self.assertIn("SpecialMailboxes", load_workbook(path).sheetnames)
        kws = {r["Keyword"]: r["Document type"] for r in settings.rows("Keywords")}
        self.assertEqual(kws["direct debit notice"], "Direct Debit Notice")
        self.assertIn("5 hanrahan", [r["Address fragments"] for r in settings.rows("Sites")])
        self.assertEqual(settings.upgrade_workbook(), [])                         # once only

    def test_a_person_s_later_choice_is_not_overwritten(self):
        settings.upgrade_workbook()
        self.set_run(Create_site_folder_in_live_mode="Yes")
        settings.upgrade_workbook()
        settings.reset()
        self.assertTrue(settings.run_bool("Create site folder in live mode", False))


class SharedMailboxTests(unittest.TestCase):
    def test_a_shared_mailbox_is_opened_by_address_when_outlook_does_not_list_it(self):
        from invoicebot.outlook_io import ClassicOutlook

        class Rec:
            def Resolve(self):
                return True

        class NS:
            Folders = []

            def CreateRecipient(self, a):
                self.asked = a
                return Rec()

            def GetSharedDefaultFolder(self, rec, n):
                return "INBOX-OF-PIFUEL" if n == 6 else None

        o = ClassicOutlook.__new__(ClassicOutlook)
        o.ns, o._inbox_cache = NS(), {}
        key = o.resolve_store(["P I Fuel Invoice", "pifuelinvoice@pearlenergy.com.au"])
        self.assertEqual(key, "shared:pifuelinvoice@pearlenergy.com.au")
        self.assertEqual(o._inbox(key), "INBOX-OF-PIFUEL")


if __name__ == "__main__":
    unittest.main()
