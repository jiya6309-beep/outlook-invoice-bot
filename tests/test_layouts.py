"""Layouts that failed on real e-mails, kept as a regression set: a fix for one supplier must not break another.
Each PDF is a rebuilt copy of the real layout (the supplier's name, ABN and the page text order are the point)."""
import os
import unittest

from tests import make_data
from tests.helpers import ONE, Isolated, connect_db, gateway, mail, run

HERE = os.path.join(os.path.dirname(__file__), "layouts")


class LayoutTests(Isolated, unittest.TestCase):
    def doc(self, subject, path):
        gw = gateway()
        em = gw.add(ONE, mail(subject, [path]))
        s = run(gw)
        rows = [dict(r) for r in connect_db(self.home).execute("SELECT * FROM docs ORDER BY id")]
        return s, (rows[-1] if rows else None), em

    def test_dorries_bill_to_ship_to_beside_the_supplier(self):
        s, r, _ = self.doc("Invoice INV-2431", os.path.join(HERE, "dorries_bill_to_ship_to_beside_supplier.pdf"))
        self.assertEqual((r["supplier"], r["number"]), ("Dorries Anglers Choice Pty Ltd", "INV-2431"))

    def test_kc_supplies_unlabelled_number_under_the_title_and_tax_as_gst(self):
        s, r, _ = self.doc("Tax Invoice 55331", os.path.join(HERE, "kc_unlabelled_number_under_title.pdf"))
        self.assertEqual((r["supplier"], r["number"], r["amount"], r["gst"]), ("KC Supplies", "55331", 256.35, 23.31))

    def test_ao_global_number_below_its_label(self):
        s, r, _ = self.doc("Tax Invoice 17112", os.path.join(HERE, "ao_global_number_below_label.pdf"))
        self.assertEqual((r["supplier"], r["number"], r["amount"], r["gst"]), ("AO Global", "17112", 34.46, 3.13))

    def test_tip_top_is_the_supplier_wherever_it_is_named(self):
        s, r, _ = self.doc("Invoice 8024987413", os.path.join(HERE, "tiptop_direct_debit_notice.pdf"))
        self.assertEqual(r["supplier"], "Tip Top Bakeries")

    def test_credit_note_with_only_a_crm_id_uses_it(self):
        p = os.path.join(self.home, "cn1.pdf")
        make_data.simple(p, "CREDIT NOTE", ["CRM ID 88123", "Date: 02/10/2026", "Total inc GST 44.00", "GST 4.00"], supplier="Zed Supplies Pty Ltd")
        s, r, _ = self.doc("Credit Note 1", p)
        self.assertEqual((r["doc_type"], r["number"]), ("Credit note", "88123"))

    def test_an_invoice_number_wins_over_a_crm_id(self):
        p = os.path.join(self.home, "cn2.pdf")
        make_data.simple(p, "CREDIT NOTE", ["CRM ID 88123", "Credit Note No: CN-777", "Date: 02/10/2026", "Total inc GST 44.00", "GST 4.00"], supplier="Zed Supplies Pty Ltd")
        s, r, _ = self.doc("Credit Note 2", p)
        self.assertEqual(r["number"], "CN-777")

    def test_order_confirmation_found_by_file_name_when_the_subject_has_no_keyword(self):
        p = os.path.join(self.home, "Order-SO106649.pdf")
        make_data.simple(p, "ORDER CONFIRMATION", ["Sales Order No SO106649", "ORDER DATE 6/10/2026", "Total 500.00"], supplier="Statewide Oil Pty Ltd")
        s, r, em = self.doc("Your order SO106649", p)
        self.assertEqual((r["doc_type"], r["status"]), ("Order Confirmation", "Parked"))
        self.assertEqual(r["doc_date"], "2026-10-06")
        self.assertIn("_06-10-2026_OC.pdf", r["current_path"])

    def test_a_purchase_order_pdf_opened_only_for_its_name_is_ignored_without_a_trace(self):
        p = os.path.join(self.home, "Purchase-Order-123.pdf")
        make_data.simple(p, "PURCHASE ORDER", ["PO No 123", "Date: 02/10/2026", "Total 500.00"], supplier="Zed Supplies Pty Ltd")
        s, r, em = self.doc("hello", p)
        self.assertIsNone(r)                                         # no record as a document
        self.assertEqual(s.c["documents"], 0)
        self.assertTrue(em.unread)                                   # the e-mail is left alone
        self.assertEqual(em.flag, 0)

    def test_the_order_name_check_can_be_switched_off(self):
        self.set_run(Check_attachment_names_for_order_confirmations="No")
        p = os.path.join(self.home, "Order-SO1.pdf")
        make_data.simple(p, "ORDER CONFIRMATION", ["Order No SO1", "Date: 02/10/2026"], supplier="Statewide Oil Pty Ltd")
        s, r, _ = self.doc("hello", p)
        self.assertIsNone(r)


if __name__ == "__main__":
    unittest.main()
