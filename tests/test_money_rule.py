"""The three amounts (total incl. GST, GST, amount excl. GST) are OPTIONAL: read as printed, never calculated, never a reason to hold a document.
Remarks says GST NOT FOUND only when all three are missing. Statements and other types are never touched by this rule."""
import os
import unittest
from unittest import mock

from openpyxl import load_workbook

from tests import make_data
from tests.helpers import ONE, Isolated, connect_db, gateway, mail, run

from invoicebot import doctypes, naming, suppliers


class MoneyRuleTests(Isolated, unittest.TestCase):
    def pdf(self, name, title, lines, supplier="Orbit Wholesale Pty Ltd"):
        path = os.path.join(self.home, name)
        make_data.simple(path, title, lines, supplier)
        return path

    def run_one(self, subject, path, hours_ago=3):
        gw = gateway()
        gw.add(ONE, mail(subject, [path], hours_ago=hours_ago))
        s = run(gw)
        row = connect_db(self.home).execute("SELECT * FROM docs ORDER BY id DESC LIMIT 1").fetchone()
        return s, row

    def test_only_the_total_is_found_and_the_rest_stays_empty(self):
        s, r = self.run_one("Tax Invoice 5501", self.pdf("a.pdf", "TAX INVOICE", ["Invoice No: 5501", "Invoice Date: 02/10/2026", "Total inc GST 110.00"]))
        self.assertEqual(r["status"], "Filed")
        self.assertEqual((r["amount"], r["gst"], r["ex_amount"]), (110.0, None, None))             # empty, not 0.00, and NOT worked out
        self.assertEqual(r["remarks"] or "", "")                                                   # one figure was found: no remark

    def test_gst_missing_but_the_other_two_found_is_not_calculated(self):
        s, r = self.run_one("Tax Invoice 5502", self.pdf("b.pdf", "TAX INVOICE", ["Invoice No: 5502", "Invoice Date: 02/10/2026", "Total excl GST 100.00", "Total inc GST 110.00"]))
        self.assertEqual(r["status"], "Filed")
        self.assertEqual((r["amount"], r["gst"], r["ex_amount"]), (110.0, None, 100.0))            # GST stays empty although it could be 10.00
        self.assertEqual(r["remarks"] or "", "")

    def test_all_three_missing_files_the_document_with_the_remark(self):
        s, r = self.run_one("Tax Invoice 5503", self.pdf("c.pdf", "TAX INVOICE", ["Invoice No: 5503", "Invoice Date: 02/10/2026", "Thank you for your business"]))
        self.assertEqual(r["status"], "Filed")                                                     # in its folder, not Unrecognised
        self.assertEqual((r["amount"], r["gst"], r["ex_amount"]), (None, None, None))
        self.assertEqual(r["remarks"], "GST NOT FOUND")
        self.assertTrue(self.files_under("TestOutput"))

    def test_figures_that_disagree_are_kept_without_any_remark(self):
        s, r = self.run_one("Tax Invoice 5504", self.pdf("d.pdf", "TAX INVOICE", ["Invoice No: 5504", "Invoice Date: 02/10/2026", "Total excl GST 100.00", "Total GST 10.00", "Total inc GST 150.00"]))
        self.assertEqual(r["status"], "Filed")
        self.assertEqual((r["amount"], r["gst"], r["ex_amount"]), (150.0, 10.0, 100.0))            # exactly as printed
        self.assertEqual(r["remarks"] or "", "")
        self.assertNotIn("add up", (r["warnings"] or "") + (r["reason"] or ""))

    def test_a_gst_free_invoice_shows_zero_not_blank(self):
        s, r = self.run_one("Tax Invoice 5505", self.pdf("e.pdf", "TAX INVOICE", ["Invoice No: 5505", "Invoice Date: 02/10/2026", "All items GST Free", "Total inc GST 55.00"]))
        self.assertEqual((r["amount"], r["gst"]), (55.0, 0.0))
        self.assertEqual(r["remarks"] or "", "")

    def test_gst_free_is_not_assumed_when_the_document_does_not_say_so(self):
        s, r = self.run_one("Tax Invoice 5506", self.pdf("f.pdf", "TAX INVOICE", ["Invoice No: 5506", "Invoice Date: 02/10/2026", "Total inc GST 55.00"]))
        self.assertEqual((r["amount"], r["gst"]), (55.0, None))

    def test_a_credit_note_with_only_a_total_is_negative(self):
        s, r = self.run_one("Credit Note CN-77", self.pdf("g.pdf", "CREDIT NOTE", ["Credit Note No: CN-77", "Date: 02/10/2026", "Total inc GST 44.00"]))
        self.assertEqual((r["doc_type"], r["amount"], r["gst"]), ("Credit note", -44.0, None))

    def test_statements_are_never_touched_by_this_rule(self):
        s, r = self.run_one("Statement for September", self.pdf("h.pdf", "WEEKLY STATEMENT", ["Statement Date: 02/10/2026", "All GST Free", "Total 500.00"]))
        self.assertEqual(r["doc_type"], "Statement")
        self.assertEqual((r["amount"], r["gst"], r["ex_amount"]), (None, None, None))             # no GST 0.00 added to a statement
        self.assertNotIn("GST", r["remarks"] or "")                                                # and no GST remark (only the NoSite note)

    def test_a_pod_is_never_touched_by_this_rule(self):
        s, r = self.run_one("Delivery Docket D-1", self.pdf("i.pdf", "PROOF OF DELIVERY", ["Delivery No: D-1", "Date: 02/10/2026", "No GST"]))
        self.assertEqual((r["doc_type"], r["gst"], r["remarks"] or ""), ("POD", None, ""))

    def subject_only(self):
        """Force the case 'the page itself does not say it is an invoice; only the e-mail subject does'."""
        return mock.patch.object(doctypes, "decide", return_value=("Invoice", "subject 'invoice'"))

    def test_the_guard_holds_a_document_with_no_amounts_whose_type_came_only_from_the_subject(self):
        with self.subject_only():
            s, r = self.run_one("Invoice for your order", self.pdf("j.pdf", "PRICE LIST", ["Invoice No: 9001", "Invoice Date: 02/10/2026", "Thank you"]))
        self.assertEqual(r["status"], "Unrecognised")
        self.assertIn("does not say it is an invoice", r["reason"])

    def test_the_guard_does_not_hold_a_document_that_has_at_least_one_amount(self):
        with self.subject_only():
            s, r = self.run_one("Invoice for your order", self.pdf("j.pdf", "PRICE LIST", ["Invoice No: 9001", "Invoice Date: 02/10/2026", "Total inc GST 20.00"]))
        self.assertEqual((r["status"], r["amount"]), ("Filed", 20.0))

    def test_the_guard_can_be_switched_off(self):
        self.set_run(Hold_documents_without_amounts_when_the_type_came_only_from_the_subject="No")
        with self.subject_only():
            s, r = self.run_one("Invoice for your order", self.pdf("j.pdf", "PRICE LIST", ["Invoice No: 9001", "Invoice Date: 02/10/2026", "Thank you"]))
        self.assertEqual((r["status"], r["remarks"]), ("Filed", "GST NOT FOUND"))

    def test_the_new_columns_are_in_the_record(self):
        self.run_one("Tax Invoice 5507", self.pdf("k.pdf", "TAX INVOICE", ["Invoice No: 5507", "Invoice Date: 02/10/2026", "Total excl GST 100.00", "Total inc GST 110.00"]))
        self.run_one("Tax Invoice 5508", self.pdf("l.pdf", "TAX INVOICE", ["Invoice No: 5508", "Invoice Date: 02/10/2026", "Nothing printed"]), hours_ago=0.3)
        wb = load_workbook(self.path("Reports", "Record_2026-10.xlsx"))
        ws = wb["Record"]
        head = [c.value for c in ws[1]]
        for name in ("Amount incl GST", "GST", "Amount excl GST", "Remarks"):
            self.assertIn(name, head)
        rows = {r[head.index("Number")]: r for r in ws.iter_rows(min_row=2, values_only=True)}
        first, second = rows["5507"], rows["5508"]
        self.assertEqual((first[head.index("Amount incl GST")], first[head.index("GST")], first[head.index("Amount excl GST")]), (110.0, None, 100.0))
        self.assertIn(first[head.index("Remarks")], (None, ""))
        self.assertEqual(second[head.index("Remarks")], "GST NOT FOUND")
        wb.close()


class OddCharacterTests(unittest.TestCase):
    def test_odd_characters_in_numbers_and_names_are_tidied(self):
        self.assertEqual(naming.tidy("INV\uf02d3121"), "INV-3121")                                  # a Symbol-font hyphen
        self.assertEqual(naming.tidy("INV\u20113121"), "INV-3121")                                  # a non-breaking hyphen
        self.assertEqual(naming.tidy("INV\u22123121"), "INV-3121")                                  # a minus sign
        self.assertEqual(naming.tidy("INV\u00ad3121\u200b"), "INV3121")                             # invisible characters disappear
        self.assertEqual(naming.tidy("A\u00a0B"), "A B")
        self.assertEqual(naming.tidy("Café Pty Ltd"), "Café Pty Ltd")                               # real letters are kept
        self.assertIsNone(naming.tidy(None))
        name = naming.build("{supplier}_{number}", "Limited", "", "INV\uf02d3121", "", "", "Invoice", "x.pdf", ".pdf")
        self.assertEqual(name, "Limited_INV-3121.pdf")


class SupplierSanityTests(unittest.TestCase):
    def test_sentences_links_and_bare_legal_words_are_not_supplier_names(self):
        for bad in ("Limited", "Pty Ltd", "View online", "View and pay online", "Legal title of these goods remains in Bonfect Pty Ltd until paid",
                    "Please remit payment to the account below", "x", "1234 5678"):
            self.assertFalse(suppliers.plausible(bad), bad)
        for good in ("Bega Cheese Limited", "Coca-Cola Europacific Partners Australia Pty Ltd", "TIP TOP Bakeries", "Perkal Pty Ltd", "P N & J H Chauncy"):
            self.assertTrue(suppliers.plausible(good), good)


if __name__ == "__main__":
    unittest.main()
