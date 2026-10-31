"""5.7.0, step 1: forwards/replies are downloaded and left unread; dockets; supplier ids; legal endings; duplicates; spreadsheets beside invoices."""
import os
import unittest

from openpyxl import Workbook

from tests import make_data
from tests.helpers import ONE, Isolated, connect_db, gateway, mail, run

from invoicebot import doctypes, parsing, pdftext, settings, suppliers


class Step1Tests(Isolated, unittest.TestCase):
    def pdf(self, name, title, lines, supplier="Orbit Wholesale Pty Ltd"):
        path = os.path.join(self.home, name)
        make_data.simple(path, title, lines, supplier)
        return path

    def inv(self, name, number, supplier="Orbit Wholesale Pty Ltd", extra=()):
        return self.pdf(name, "TAX INVOICE", [f"Invoice No: {number}", "Invoice Date: 02/10/2026", *extra, "Total GST 10.00", "Total inc GST 110.00"], supplier)

    def rows(self):
        return [dict(r) for r in connect_db(self.home).execute("SELECT * FROM docs ORDER BY id")]

    # ---------------------------------------------------------------- forwards and replies
    def test_a_forward_without_a_document_word_is_downloaded_filed_and_left_unread(self):
        gw = gateway()
        m = gw.add(ONE, mail("FW: your order", [self.inv("a.pdf", "9101")]))
        run(gw)
        r = self.rows()
        self.assertEqual([x["status"] for x in r], ["Filed"])
        self.assertTrue(m.unread)                                          # UNREAD
        self.assertNotEqual(m.flag, 1)                                     # no tick (1 = complete)
        self.assertEqual(m.flag, 2)                                         # red flag
        self.assertEqual(m.categories, [])

    def test_a_reply_with_an_unrecognised_document_goes_to_unrecognised_and_stays_unread(self):
        gw = gateway()
        p = self.pdf("b.pdf", "PRICE LIST", ["hello"])
        m = gw.add(ONE, mail("RE: question", [p]))
        run(gw)
        self.assertEqual([x["status"] for x in self.rows()], ["Unrecognised"])
        self.assertTrue(m.unread)
        self.assertEqual(m.flag, 2)
        self.assertEqual(m.categories, [])

    def test_a_normal_email_is_still_read_and_ticked(self):
        gw = gateway()
        m = gw.add(ONE, mail("Tax invoice 9102", [self.inv("c.pdf", "9102")]))
        run(gw)
        self.assertFalse(m.unread)
        self.assertEqual(m.flag, 1)

    def test_a_blocked_word_still_blocks_a_forward(self):
        gw = gateway()
        m = gw.add(ONE, mail("FW: overdue invoice", [self.inv("d.pdf", "9103")]))
        run(gw)
        self.assertEqual(self.rows(), [])

    def test_the_old_behaviour_can_be_switched_back(self):
        self.set_run(**{"Download attachments of FW/RE e-mails": "No"})
        gw = gateway()
        gw.add(ONE, mail("FW: your order", [self.inv("e.pdf", "9104")]))
        run(gw)
        self.assertEqual(self.rows(), [])

    # ---------------------------------------------------------------- dockets
    def test_a_docket_is_filed_like_an_invoice_without_amounts(self):
        p = self.pdf("f.pdf", "DELIVERY DOCKET", ["Docket No: DK5501", "Date: 02/10/2026", "Qty 5 bread"])
        gw = gateway()
        gw.add(ONE, mail("Docket DK5501", [p]))
        run(gw)
        (r,) = self.rows()
        self.assertEqual((r["status"], r["doc_type"]), ("Filed", "POD"))
        self.assertEqual((r["amount"], r["gst"], r["ex_amount"]), (None, None, None))
        self.assertFalse(r["remarks"])                                     # no GST NOT FOUND
        self.assertTrue(r["dest_name"].endswith("_POD.pdf"))

    def test_the_page_overrides_a_misleading_subject(self):
        p = self.pdf("g.pdf", "DEL DOCKET", ["DEL DOCKET : FZ99120", "THIS IS ONLY A DELIVERY DOCKET, YOU WILL BE INVOICED", "DEL DATE : 01/10/2026"])
        gw = gateway()
        gw.add(ONE, mail("Tax invoice for September", [p]))
        run(gw)
        (r,) = self.rows()
        self.assertEqual(r["doc_type"], "POD")
        self.assertEqual(r["number"], "FZ99120")

    # ---------------------------------------------------------------- supplier from ids, legal endings
    def test_supplier_from_the_abn_alone(self):
        self.add_rows("SupplierIDs", [("Zed Test Bakery", "98765432109", "", "")])
        p = self.inv("h.pdf", "7001", supplier="")
        gw = gateway()
        gw.add(ONE, mail("Tax invoice 7001", [p]))
        run(gw)
        (r,) = self.rows()
        self.assertEqual((r["supplier"], r["status"]), ("Zed Test Bakery", "Filed"))

    def test_supplier_from_the_email_domain_alone(self):
        self.add_rows("SupplierIDs", [("Mail Bakery", "", "mailbakery.com.au", "")])
        p = self.inv("i.pdf", "7002", supplier="", extra=["EM: orders@mailbakery.com.au"])
        gw = gateway()
        gw.add(ONE, mail("Tax invoice 7002", [p]))
        run(gw)
        (r,) = self.rows()
        self.assertEqual(r["supplier"], "Mail Bakery")

    def test_a_name_cut_before_its_ending_is_still_the_supplier(self):
        text = "Tax Invoice\nBill to\nMCALLISTER & SONS PTY.\nPearl Energy No 5 Pty Ltd\nLIMITED\nT/A Gunna-Doo Bakery\nINVOICE NO 5"
        name, src = suppliers.identify(text)
        self.assertEqual(name, "McAllister & Sons Pty")

    def test_pty_and_pvt_are_the_same_ending(self):
        self.assertEqual(settings.supplier_key("Envisage Holdings Pty Ltd"), settings.supplier_key("ENVISAGE HOLDINGS PVT. LIMITED"))
        self.assertEqual(suppliers.identify("Tax Invoice\nSharma Traders Pvt Ltd\nInvoice No 1\n")[0], "Sharma Traders Pvt Ltd")

    def test_the_payee_in_the_bank_box_is_the_supplier(self):
        text = "Tax Invoice\nBill to Envisage Holdings Pty Ltd\nMobil Dubbo\nView online\nAccount Name: Envisage Holdings PTY LTD\nBSB: 082 842"
        self.assertEqual(suppliers.identify(text)[0], "Envisage Holdings Pty Ltd")

    def test_a_service_company_name_does_not_make_an_invoice(self):
        typ, _ = doctypes.decide("", "x.pdf", "Bonfect Snackfood Services\nPrice list", None, "", "Bonfect Snackfood Services")
        self.assertIsNone(typ)

    # ---------------------------------------------------------------- reading
    def test_sept_is_a_month(self):
        import datetime
        self.assertEqual(parsing.parse_date("26 Sept 2026"), datetime.date(2026, 9, 26))
        self.assertEqual(parsing.parse_date("​26 Sept 2026"), datetime.date(2026, 9, 26))

    def test_letters_delivered_one_by_one_are_joined(self):
        w = lambda t, a, b: {"text": t, "x0": a, "x1": b, "top": 10, "bottom": 20}
        words = [w("T", 0, 7), w("otal", 6.5, 20), w("Inc", 23, 36), w("GST", 38, 52), w("1718.45", 80, 115)]
        line = pdftext.build_lines(words)[0]
        self.assertEqual(line.text, "Total Inc GST 1718.45")

    # ---------------------------------------------------------------- duplicates and spreadsheets
    def test_a_reprint_is_skipped_and_noted(self):
        gw = gateway()
        gw.add(ONE, mail("Tax invoice 8001", [self.inv("j.pdf", "8001")], hours_ago=5))
        gw.add(ONE, mail("Tax invoice 8001 reprint", [self.inv("k.pdf", "8001", extra=["* Reprint 2 *"])], hours_ago=2))
        s = run(gw)
        r = self.rows()
        self.assertEqual(len(r), 1)
        self.assertIn("duplicate", r[0]["remarks"] or "")
        self.assertEqual(len([f for f in self.files_under("TestOutput") if f.endswith(".pdf")]), 1)

    def test_a_spreadsheet_is_filed_beside_its_invoice(self):
        x = os.path.join(self.home, "breakdown.xlsx")
        wb = Workbook()
        wb.active.append(["item", "qty"]); wb.active.append(["pies", 390])
        wb.save(x)
        gw = gateway()
        gw.add(ONE, mail("Tax invoice 8002", [x, self.inv("l.pdf", "8002")]))
        run(gw)
        files = self.files_under("TestOutput")
        pdfs = [f for f in files if f.endswith(".pdf")]
        xl = [f for f in files if f.endswith(".xlsx")]
        self.assertEqual((len(pdfs), len(xl)), (1, 1))
        self.assertEqual(os.path.dirname(pdfs[0]), os.path.dirname(xl[0]))
        self.assertIn("_breakdown", xl[0])


if __name__ == "__main__":
    unittest.main()


class UpgradeTests(Isolated, unittest.TestCase):
    def test_an_older_settings_file_gets_the_new_sheets_and_settings_and_keeps_its_own_edits(self):
        from openpyxl import load_workbook
        path = settings.settings_file()
        wb = load_workbook(path)
        for name in ("SupplierIDs", "LegalEndings", "SiteAliases"):
            del wb[name]
        run_ws = wb["Run"]
        for r in range(run_ws.max_row, 1, -1):
            if run_ws.cell(r, 1).value in ("Month window (months)", "Download attachments of FW/RE e-mails"):
                run_ws.delete_rows(r)
        wb["Keywords"].append(["my own word", "Invoice", "Yes"])
        wb.save(path)
        wb.close()
        settings.reset()
        added = settings.upgrade_workbook()
        self.assertTrue(any("SupplierIDs" in a for a in added))
        self.assertTrue(any("Month window" in a for a in added))
        wb = load_workbook(path)
        self.assertIn("SupplierIDs", wb.sheetnames)
        self.assertEqual(wb["SupplierIDs"].cell(2, 1).value, "JP Bakery")
        self.assertIn("my own word", [r[0] for r in wb["Keywords"].iter_rows(min_row=2, values_only=True)])
        wb.close()
        self.assertEqual(settings.upgrade_workbook(), [])                    # nothing more to add the second time


class SupplierFromEmailDomainTests(unittest.TestCase):
    def test_name_is_found_from_the_suppliers_own_email_domain(self):
        from invoicebot import suppliers
        text = ("Tax Invoice\nBill to Ship to Gold Coast Cookies\nPearl Energy Morriset 4/37 Machinery Dr\n"
                "one@pearlenergy.com.au hello@goldcoastcookies.com.au\nABN 66641150254\n")
        self.assertEqual(suppliers.name_from_email_domain(text), "Gold Coast Cookies")
        self.assertIsNone(suppliers.name_from_email_domain("one@pearlenergy.com.au bob@gmail.com"))


class TipTopTests(Isolated, unittest.TestCase):
    def test_tip_top_is_the_supplier_wherever_it_is_named_or_its_abn_is_printed(self):
        from invoicebot import settings
        text = "A Division of George Weston Foods Limited\nABN 45 008 429 632\nDIRECT DEBIT NOTICE\nTIP TOP CREDIT TEAM"
        self.assertEqual(settings.supplier_id_row_for(text)["Supplier (record name)"], "Tip Top Bakeries")
        self.assertEqual(settings.supplier_id_row_for("Regards TIP TOP Bakeries 1300")["Supplier (record name)"], "Tip Top Bakeries")


class SideBySideColumnsTests(Isolated, unittest.TestCase):
    def test_supplier_printed_beside_the_bill_to_ship_to_labels_is_found(self):
        text = ("Tax Invoice\nBill to Ship to Dorries Anglers Choice Pty Ltd\nAmpol Nanango Ampol Nanango PO Box 507\n"
                "28 Henry St 28 Henry St Gin Gin QLD 4671\nABN: 22683633865\nthree@pearlenergy.com.au\nAmount due Invoice number\n")
        self.assertEqual(suppliers.heuristic_names(text)[0], "Dorries Anglers Choice Pty Ltd")

    def test_a_customer_named_after_bill_to_is_never_taken_as_the_supplier(self):
        text = "Tax Invoice\nBill to Pearl Energy Pty Ltd\n12 Some Street\nAmount due Invoice number\n"
        self.assertNotEqual(suppliers.heuristic_names(text)[0], "Pearl Energy Pty Ltd")


class StackedLabelTests(Isolated, unittest.TestCase):
    TEXT = ("TAX INVOICE\nInvoice # Invoice To AO Global Pty Ltd\n17112 HONEY sHAH 3 / 10A Durkin Place\nOrder # Pearl Energy Londonderry PEAKHURST NSW 2210\n"
            "17112 Cnr Londonderry Rd & AUSTRALIA\nInvoice Date Carrington Rd 0450 399 166\n06 Oct 2026 LONDONDERRY NSW 2753 ABN: 42 089 269 778\n"
            "Terms AUSTRALIA\nDue On Invoice Date Phone: 0433685129\nPayment Due Email:\n06 Oct 2026 pearl.londonderry@pearlenerg\n")

    def test_the_number_below_its_label_is_read(self):
        from invoicebot import legacy
        self.assertEqual(legacy.find_invoice_number(self.TEXT), "17112")

    def test_the_supplier_beside_invoice_hash_and_invoice_to_is_found(self):
        self.assertEqual(suppliers.heuristic_names(self.TEXT)[0], "AO Global Pty Ltd")

    def test_payment_due_is_never_a_supplier_name(self):
        self.assertIsNone(suppliers.heuristic_names("Tax Invoice\nPayment Due\n06 Oct 2026\nInvoice No 55\n")[1])
