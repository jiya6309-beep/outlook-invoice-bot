"""5.16.0: MobiGo's Order # is its invoice number; order documents (MobiGo order printouts, Fresho / Rarity 'THIS IS NOT AN INVOICE',
PFD Sales Orders) are Order Confirmations, never invoices and never _Unrecognised; Peters' INV DATE 220726 = 22/07/2026."""
import os
import shutil
import tempfile
import unittest

from tests import docs_v516 as D
from tests.helpers import DATA, ONE, Isolated, connect_db, gateway, mail, run
from tests.test_site_layout import site_row

from invoicebot import doctypes, labelreader, pdftext, settings

SITES = [site_row("Pearl Energy Woodridge", "101 Compton Road, Woodridge QLD 4114"),
         site_row("BP Quirindi", "4 Duke Street, Quirindi NSW 2343", company="Pearl Energy No 3 Pty Ltd"),
         site_row("Pearl Energy Anna Bay", "4136 Nelson Bay Road, Anna Bay NSW 2316"),
         site_row("Pearl Energy Deagon", "11 Deport Road, Deagon QLD 4017"),
         site_row("Mobil Ballina", "485 River Street, Ballina NSW 2478", company="Pearl Energy No 2 Pty Ltd")]


class Base(Isolated, unittest.TestCase):
    def setUp(self):
        super().setUp()
        self.tmp = tempfile.mkdtemp(prefix="v516_")
        self.set_run(Destination_pattern="{MONTH_FOLDER}\\{COMPANY}\\{SITE}")
        self.add_rows("Sites", SITES)

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)
        super().tearDown()

    def deliver(self, builder, subject="Document attached", **mail_kw):
        p = os.path.join(self.tmp, builder.__name__ + ".pdf")
        builder(p)
        g = gateway()
        g.add(ONE, mail(subject, [p], **mail_kw))
        s = run(g)
        return s, dict(connect_db(self.home).execute("SELECT * FROM docs").fetchone())


class MobiGoTests(Base):
    def test_the_tax_invoice_uses_the_order_number(self):
        s, d = self.deliver(D.mobigo_tax_invoice, "Invoice")
        self.assertEqual((d["doc_type"], d["number"], d["status"]), ("Invoice", "S506325", "Filed"))
        self.assertEqual(d["doc_date"], "2026-07-23")
        self.assertTrue(d["dest_name"].endswith("_S506325.pdf"))
        self.assertEqual(d["site"], "Pearl Energy Woodridge")

    def test_the_order_printout_is_an_order_confirmation(self):
        s, d = self.deliver(D.mobigo_order, "Order S00372")
        self.assertEqual((d["doc_type"], d["status"]), ("Order Confirmation", "Parked"))
        self.assertEqual((d["site"], d["doc_date"]), ("BP Quirindi", "2026-10-22"))
        self.assertIn("_Order Confirmations", d["dest_dir"])
        self.assertEqual(s.c["Unrecognised"], 0)


class OrderTests(Base):
    def test_fresho_not_an_invoice_is_an_order_confirmation(self):
        s, d = self.deliver(D.fresho_order, "Invoice for your order")                     # even with 'Invoice' in the subject
        self.assertEqual((d["doc_type"], d["site"], d["doc_date"]), ("Order Confirmation", "Pearl Energy Anna Bay", "2026-07-23"))
        self.assertTrue(d["dest_name"].endswith("_Pearl Energy Anna Bay_23-07-2026_OC.pdf"), d["dest_name"])
        self.assertEqual(s.c["Unrecognised"], 0)

    def test_pfd_sales_order_is_an_order_confirmation(self):
        s, d = self.deliver(D.pfd_sales_order, "Sales Order VU959200")
        self.assertEqual((d["doc_type"], d["site"], d["doc_date"]), ("Order Confirmation", "Pearl Energy Deagon", "2026-07-23"))
        self.assertEqual(s.c["Unrecognised"], 0)

    def test_an_invoice_that_mentions_its_sales_order_stays_an_invoice(self):
        header = "Acme Foods\nTAX INVOICE\nSales Order No 4455"
        self.assertEqual(doctypes.decide("", "", header, None, header + "\nInvoice No 1001")[0], "Invoice")
        self.assertEqual(doctypes.best_match("Sales Invoices for Sales Order")[0], "Invoice")         # the real subject

    def test_a_statement_that_says_not_a_tax_invoice_stays_a_statement(self):
        text = "Acme Foods\nSTATEMENT\nThis is not a tax invoice"
        self.assertEqual(doctypes.decide("", "", text, None, text)[0], "Statement")

    def test_a_tax_invoice_with_a_not_a_tax_invoice_remittance_slip_stays_an_invoice(self):
        page = "Acme Foods\nTAX INVOICE\nInvoice No 1001\n...\nREMITTANCE ADVICE - this is not a tax invoice"
        self.assertEqual(doctypes.decide("", "", "Acme Foods\nTAX INVOICE\nInvoice No 1001", None, page)[0], "Invoice")

    def test_an_order_line_on_a_page_that_says_invoice_is_not_decided_by_the_order_rule(self):
        page = "Acme\nOrder No: 12345\nInvoice\nTotal 10.00"
        self.assertNotEqual(doctypes.decide("Invoice 5", "", "Acme", None, page)[0], "Order Confirmation")


class PetersTests(Base):
    def test_inv_date_ddmmyy_is_read(self):
        s, d = self.deliver(D.peters_invoice, "Peters invoice 11620355")
        self.assertEqual((d["doc_type"], d["number"], d["doc_date"], d["status"]), ("Invoice", "11620355", "2026-07-22", "Filed"))
        self.assertEqual((d["site"], d["amount"]), ("Mobil Ballina", 330.80))

    def test_a_six_digit_number_is_a_date_only_under_a_ddmmyy_label(self):
        self.assertEqual(labelreader.parse_ddmmyy("220726").isoformat(), "2026-07-22")
        self.assertIsNone(labelreader.parse_ddmmyy("320726"))
        self.assertNotIn("inv date (ddmmyy)", settings.labels_for("Someone Else", "Someone Else")["date"])
        p = os.path.join(self.tmp, "peters.pdf")
        D.peters_invoice(p)
        r = labelreader.read(pdftext.load(p), settings.labels_for("Someone Else", "Someone Else"))
        self.assertNotEqual(r["invoice_date"] and r["invoice_date"].isoformat(), "2026-08-31")    # never the due date 310826


class UpgradeTests(Isolated, unittest.TestCase):
    def test_older_files_get_the_mobigo_and_peters_rows_once(self):
        shutil.copyfile(os.path.join(DATA, "settings_payments_team_5_13.xlsx"), settings.settings_file())
        settings.reset()
        added = settings.upgrade_workbook()
        settings.reset()
        self.assertTrue(any("MobiGo" in a for a in added) and any("Peters" in a for a in added))
        self.assertEqual(settings.labels_for("MobiGo", "Mobigo")["invoice_no"][0], "order #")
        self.assertEqual(settings.labels_for("Peters", "Peters")["date"][0], "inv date (ddmmyy)")
        self.assertEqual(settings.upgrade_workbook(), [])

    def test_an_existing_peters_row_is_completed_not_replaced(self):
        self.add_rows("Labels", [["Peters", "deliver to", "invoice no", "date", "", "", "total due"]])
        from openpyxl import load_workbook
        wb = load_workbook(settings.settings_file())
        ws = wb["Labels"]
        for r in range(ws.max_row, 1, -1):                            # a 5.15.0 file: without the new default rows
            if ws.cell(r, 1).value in ("MobiGo",) or (ws.cell(r, 1).value == "Peters" and "ddmmyy" in str(ws.cell(r, 4).value)):
                ws.delete_rows(r)
        run_ws = wb["Run"]
        for r in range(2, run_ws.max_row + 1):
            if run_ws.cell(r, 1).value == "Settings version":
                run_ws.cell(r, 2).value = "5.15.0"
        wb.save(settings.settings_file())
        settings.reset()
        settings.upgrade_workbook()
        settings.reset()
        peters = [r for r in settings.rows("Labels") if r["Supplier name contains"] == "Peters"]
        self.assertEqual(len(peters), 1)
        self.assertEqual(peters[0]["Invoice date label"], "inv date (ddmmyy)|date")
        self.assertEqual(peters[0]["Total (incl. GST) label"], "total due")


if __name__ == "__main__":
    unittest.main()
