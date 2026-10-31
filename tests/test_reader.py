"""Reader tests: invoice layouts built in code, checked against what the reader must conclude."""
import datetime
import os
import unittest

from tests.helpers import Isolated  # noqa: F401  (also puts the project on sys.path)

from invoicebot import doctypes, extractor, labelreader, settings, sites, suppliers
from invoicebot.pdftext import Document, build_lines

D = datetime.date


def read(rows, supplier=""):
    return labelreader.read(Document.synthetic(rows), settings.labels_for(supplier, supplier))


def money_ok(r, amount, gst):
    return r["amount"] == amount and r["gst"] == gst and not r["problems"]


class PasswordPdfTests(unittest.TestCase):
    def test_a_password_protected_pdf_says_so(self):
        import tempfile
        from reportlab.pdfgen import canvas
        d = tempfile.mkdtemp()
        p = os.path.join(d, "locked.pdf")
        c = canvas.Canvas(p, encrypt="secret")
        c.drawString(60, 700, "Invoice No: LOCK-1 total 55.00")
        c.showPage()
        c.save()
        res = extractor.analyse(p, {"subject": "Tax Invoice LOCK-1", "orig_name": "locked.pdf"})
        self.assertTrue(res["fatal"])
        self.assertIn("password-protected", res["problems"][0])


class ReaderTests(Isolated, unittest.TestCase):

    def test_01_values_on_the_same_line(self):
        r = read([[(20, "ACME FOODS PTY LTD")], [(20, "Invoice No: INV-1001")], [(20, "Invoice Date: 02/10/2026")],
                  [(300, "Total excl GST"), (450, "$100.00")], [(300, "Total GST"), (450, "$10.00")],
                  [(300, "Total inc GST"), (450, "$110.00")]])
        self.assertEqual(r["invoice_number"], "INV-1001")
        self.assertEqual(r["invoice_date"], D(2026, 10, 2))
        self.assertTrue(money_ok(r, 110.0, 10.0))

    def test_02_table_layout_headings_above_values(self):
        r = read([[(20, "Invoice No"), (200, "Invoice Date"), (380, "Customer")],
                  [(20, "90210"), (200, "02/10/2026"), (380, "Harristown")],
                  [(300, "Subtotal"), (400, "GST"), (500, "Total")],
                  [(300, "100.00"), (400, "10.00"), (500, "110.00")]])
        self.assertEqual(r["invoice_number"], "90210")
        self.assertEqual(r["invoice_date"], D(2026, 10, 2))
        self.assertTrue(money_ok(r, 110.0, 10.0))

    def test_03_gst_total_and_sub_total_are_not_the_total(self):
        r = read([[(20, "Invoice No 5566")], [(20, "Date 02/10/2026")], [(300, "Sub Total"), (450, "100.00")],
                  [(300, "GST Total"), (450, "10.00")], [(300, "Total"), (450, "110.00")]])
        self.assertTrue(money_ok(r, 110.0, 10.0))

    def test_04_a_missing_amount_stays_empty_and_nothing_is_calculated(self):
        r = read([[(20, "Invoice No 7788")], [(20, "Date 02/10/2026")],
                  [(300, "Total excl GST"), (450, "100.00")], [(300, "Total inc GST"), (450, "100.00")]])
        self.assertEqual((r["amount"], r["gst"], r["ex_amount"]), (100.0, None, 100.0))          # GST is NOT worked out
        self.assertFalse(r["problems"])
        r = read([[(20, "Invoice No 7788")], [(20, "Date 02/10/2026")],
                  [(300, "Total excl GST"), (450, "100.00")], [(300, "Total inc GST"), (450, "110.00")]])
        self.assertEqual((r["amount"], r["gst"], r["ex_amount"]), (110.0, None, 100.0))          # not 10.00: the bot calculates nothing
        self.assertFalse(r["problems"])
        self.assertFalse(any("worked out" in n for n in r["notes"]))

    def test_05_figures_that_do_not_add_up_are_kept_as_read(self):
        base = [[(20, "Invoice No 7788")], [(20, "Date 02/10/2026")]]
        r = read(base + [[(300, "Total excl GST"), (450, "100.00")], [(300, "Total GST"), (450, "10.00")],
                         [(300, "Total inc GST"), (450, "150.00")]])
        self.assertEqual((r["amount"], r["gst"], r["ex_amount"]), (150.0, 10.0, 100.0))
        self.assertFalse(r["problems"])                                                           # no warning, nothing blanked
        r = read(base + [[(300, "Total GST"), (450, "500.00")], [(300, "Total inc GST"), (450, "110.00")]])
        self.assertEqual((r["amount"], r["gst"]), (110.0, 500.0))
        self.assertFalse(r["problems"])
        r = read(base + [[(300, "Total GST"), (450, "10.00")], [(300, "Total inc GST"), (450, "900.00")]])
        self.assertEqual((r["amount"], r["gst"]), (900.0, 10.0))
        self.assertFalse(r["problems"])

    def test_06_taxable_amount_with_gst_free_items(self):
        r = read([[(20, "Invoice No 7788")], [(20, "Date 02/10/2026")], [(300, "Taxable Amount"), (450, "100.00")],
                  [(300, "GST"), (450, "10.00")], [(300, "Total inc GST"), (450, "135.00")]])
        self.assertFalse(r["problems"])
        self.assertEqual((r["amount"], r["gst"]), (135.0, 10.0))

    def test_13_item_table_is_not_a_summary(self):
        r = read([[(14, "Product Code"), (90, "Description"), (330, "Unit Price"), (430, "Net Value"), (500, "Total incl Taxes")],
                  [(14, "3766"), (90, "DF Full Cream Milk"), (430, "29.97"), (500, "29.97")],
                  [(14, "7768"), (90, "Dare Espresso"), (430, "13.20"), (500, "14.52")]])
        self.assertEqual((r["amount"], r["gst"], r["ex_amount"]), (None, None, None))             # a line of the table is not the total
        self.assertFalse(any("total" in p or "GST" in p for p in r["problems"]))                  # and a missing total no longer blocks

    def test_07_date_formats_and_the_due_date_trap(self):
        for text, expected in [("2nd October 2026", D(2026, 10, 2)), ("2026-10-02", D(2026, 10, 2)),
                               ("02-Oct-26", D(2026, 10, 2)), ("02.10.2026", D(2026, 10, 2))]:
            with self.subTest(text=text):
                r = read([[(20, "Invoice No 7788")], [(20, "Invoice Date: " + text)]])
                self.assertEqual(r["invoice_date"], expected)
        r = read([[(20, "Invoice No 7788")], [(20, "Due Date: 30/10/2026")], [(20, "Date: 02/10/2026")]])
        self.assertEqual(r["invoice_date"], D(2026, 10, 2))

    def test_08_amount_formats(self):
        r = read([[(20, "Invoice No 7788")], [(20, "Date 02/10/2026")], [(300, "Total excl GST"), (450, "(100.00)")],
                  [(300, "Total GST"), (450, "(10.00)")], [(300, "Total inc GST"), (450, "(110.00)")]])
        self.assertTrue(money_ok(r, -110.0, -10.0))
        r = read([[(20, "Invoice No 7788")], [(20, "Date 02/10/2026")], [(300, "GST (10%)"), (450, "A$10.00")],
                  [(300, "Total (AUD)"), (450, "A$110.00")]])
        self.assertEqual((r["amount"], r["gst"]), (110.0, 10.0))

    def test_09_supplier_labels_come_first(self):
        r = read([[(20, "Document No 778899")], [(20, "Transaction Date 02/10/2026")],
                  [(300, "Excl GST Value"), (450, "100.00")], [(300, "GST Value"), (450, "10.00")],
                  [(300, "Total Value"), (450, "110.00")]], supplier="Bidfood Australia Ltd")
        self.assertEqual(r["invoice_number"], "778899")
        self.assertTrue(money_ok(r, 110.0, 10.0))
        r = read([[(20, "INV:8024871400")], [(20, "Date: 06/09/2026")], [(300, "Merchandise Amount"), (450, "50.00")],
                  [(300, "GST"), (450, "5.00")], [(300, "Invoice Total"), (450, "55.00")]], supplier="Lidcombe Ice")
        self.assertEqual(r["invoice_number"], "8024871400")
        self.assertTrue(money_ok(r, 55.0, 5.0))

    def test_10_site_from_the_deliver_to_block(self):
        self.add_rows("Sites", [["Pearl Energy Harristown", "Pearl Energy No 1 Pty Ltd", "", "529 alderley|alderley street harristown",
                                 "", "", "Yes", "", ""],
                                ["Pearl Energy Toowoomba", "Pearl Energy No 1 Pty Ltd", "", "154 main", "", "", "Yes", "", ""]])
        rows = [[(20, "Invoice To:"), (300, "Deliver To:")],
                [(20, "Colyton Fuels Pty Ltd"), (300, "Pearl Energy Harristown")],
                [(20, "154 Main St"), (300, "529 Alderley Street")]]
        r = read(rows + [[(20, "Invoice No 7788")], [(20, "Date 02/10/2026")]])
        m = sites.match(Document.synthetic(rows).text, blocks=r["blocks"])
        self.assertEqual(m["site"], "Pearl Energy Harristown")
        self.assertIn("deliver-to", m["source"])

    def test_11_email_subject_cross_check(self):
        self.assertEqual(extractor.subject_warnings("FW: Invoice 2155 from BHAVANI", "2155", ""), [])
        self.assertEqual(len(extractor.subject_warnings("FW: Your Swap'n'Go Invoice 1827604001", "1830023247",
                                                        "x 1827604001 y 1830023247")), 1)

    def test_12_mixed_font_sizes_are_one_line(self):
        words = [{"text": "Total", "x0": 300, "x1": 330, "top": 714.0, "bottom": 734.0},
                 {"text": "Inc", "x0": 335, "x1": 350, "top": 714.0, "bottom": 734.0},
                 {"text": "GST", "x0": 355, "x1": 385, "top": 714.0, "bottom": 734.0},
                 {"text": "484.30", "x0": 470, "x1": 520, "top": 722.6, "bottom": 731.6},
                 {"text": "Next", "x0": 300, "x1": 330, "top": 740.0, "bottom": 749.0}]
        lines = build_lines(words)
        self.assertEqual(len(lines), 2)
        self.assertIn("484.30", lines[0].text)

    def test_14_total_credit_and_gst_included(self):
        r = read([[(20, "Invoice No. 927238697")], [(20, "Invoice Date 01/10/2026")],
                  [(300, "TOTAL CREDIT:"), (450, "$ 44.49 AUD")], [(300, "Total GST included:"), (450, "$ 1.32 AUD")]])
        self.assertEqual((r["amount"], r["gst"]), (44.49, 1.32))

    def test_15_totals_row_under_wrapped_headings(self):
        head = [[(20, "Description"), (150, "Net Total"), (230, "Net Total ex."), (320, "Total GST"), (410, "Total")],
                [(150, "ex.CDL"), (230, "GST"), (410, "Payable")],
                [(20, "Pepsi Max"), (150, "49.71"), (230, "53.07"), (320, "5.31"), (410, "58.38")]]
        totals = [[(20, "Total AUD"), (150, "199.71"), (230, "212.07"), (320, "21.21"), (410, "233.28")]]
        r = read(head + totals + [[(20, "Invoice No 9019088353")], [(20, "Date 02/10/2026")]])
        self.assertEqual((r["amount"], r["gst"]), (233.28, 21.21))
        self.assertFalse(r["problems"])
        r = read(head + [[(20, "Invoice No 9019088353")], [(20, "Date 02/10/2026")]])
        self.assertIsNone(r["amount"])

    def test_16_unlabelled_date_and_due_date(self):
        r = read([[(20, "9019088353")], [(20, "TAX INVOICE")], [(20, "02.10.2026")], [(300, "Payment Due Date: 28.11.2026")],
                  [(300, "Total inc GST"), (450, "110.00")], [(300, "Total GST"), (450, "10.00")]])
        self.assertEqual(r["invoice_date"], D(2026, 10, 2))
        self.assertTrue(r["notes"])
        r = read([[(20, "Invoice No 1234")], [(300, "Payment Due Date: 28.11.2026")]])
        self.assertIsNone(r["invoice_date"])

    def test_17_invoice_number_from_subject_only_when_in_the_pdf(self):
        self.assertEqual(extractor.hint_invoice_number("FW: Asahi Customer Invoice 9019088353 6b4c5abe_x.PDF", "... 9019088353 ..."), "9019088353")
        self.assertIsNone(extractor.hint_invoice_number("FW: Bega Inv or Adj 0927238697, Acc 241668", "Payer Account: 241668 927238697"))
        self.assertIsNone(extractor.hint_invoice_number("FW: Invoice 5551234", "nothing about it"))
        # the file-name 'Account_9014462' must still be recognised as an account number (old bug)
        self.assertIsNone(extractor.hint_invoice_number("Lactalis_Account_9014462_Stop_Supply.PDF", "Account 9014462 stop supply"))

    def test_18_supplier_names_and_credit_notes(self):
        text = "ADJUSTMENT BDD Australia Pty Ltd ABN: 68 083 019 390\nNOTE - CREDIT\nInvoice No. 927238697"
        self.assertEqual(suppliers.heuristic_name(text), "BDD Australia Pty Ltd")
        self.assertEqual(settings.canonical_supplier("BDD Australia Pty Ltd")[0], "Bega Cheese Limited")
        typ, _ = doctypes.decide(header="ADJUSTMENT\nBDD Australia Pty Ltd ABN: 68\nPO Box 23084\nNOTE - CREDIT\nInvoice No. 927238697")
        self.assertEqual(typ, "Credit note")
        typ, _ = doctypes.decide(header="TAX INVOICE\nAcme Pty Ltd\nPayment terms: credit note policy")
        self.assertEqual(typ, "Invoice")

    def test_19_real_world_supplier_traps(self):
        # a weekly statement title is never a supplier; address leftovers are not a supplier; totals lines are not a name
        self.assertIsNone(suppliers.heuristic_name("WEEKLY STATEMENT\nTotal inc GST 99.00\nInvoice No 1"))
        self.assertNotIn(suppliers.heuristic_name("PEARL ENERGY MAYFIELD WEST\n15 MURRAY DWYER CIRCUIT\nMAYFIELD NSW 2304\nAUSTRALIA\nTAX INVOICE"), ("Australia",))
        self.assertEqual(suppliers.match_logo_text("RARITY WHOLESAI", suppliers.known_names())[0], "Rarity Wholesale")


if __name__ == "__main__":
    unittest.main()
