"""The company's supplier list (ABN + e-mail domain) is built in: every supplier is found from its ABN alone, from its e-mail domain alone,
and a supplier with several sites (several ABNs) keeps ONE name."""
import re
import unittest

from invoicebot import settings
from tests.helpers import Isolated


class SupplierListTests(Isolated, unittest.TestCase):
    def test_every_supplier_is_found_from_its_abn_alone(self):
        bad = []
        for name, abn, _dom, _acct in settings.SUPPLIER_IDS:
            if not abn:
                continue
            spaced = f"{abn[:2]} {abn[2:5]} {abn[5:8]} {abn[8:]}"
            row = settings.supplier_id_row_for(f"TAX INVOICE\nABN {spaced}\nTotal 10.00")
            if not row or row["Supplier (record name)"] != name:
                bad.append((name, abn, row and row["Supplier (record name)"]))
        self.assertEqual(bad, [])

    def test_every_supplier_with_a_domain_is_found_from_it_alone(self):
        bad = []
        for name, _abn, dom, _acct in settings.SUPPLIER_IDS:
            if not dom:
                continue
            row = settings.supplier_id_row_for(f"Remittance to accounts@{dom}\nTotal 10.00")
            if not row or row["Supplier (record name)"] != name:
                bad.append((name, dom, row and row["Supplier (record name)"]))
        self.assertEqual(bad, [])

    def test_a_supplier_with_several_sites_has_one_name(self):
        for name, abns in (("Peleguy", ("81606668713", "98832539665")), ("Peters", ("11154314913", "22000802040")), ("Nippy's", ("26007841492", "39645758336"))):
            for abn in abns:
                spaced = f"{abn[:2]} {abn[2:5]} {abn[5:8]} {abn[8:]}"
                self.assertEqual(settings.supplier_id_row_for(f"ABN {spaced}")["Supplier (record name)"], name)

    def test_free_mail_and_missing_domains_are_never_used(self):
        free = {"gmail.com", "outlook.com", "yahoo.com", "hotmail.com"}
        for name, abn, dom, _acct in settings.SUPPLIER_IDS:
            self.assertNotIn(dom, free, name)
            self.assertTrue(abn or dom, name)
            self.assertTrue(not abn or re.fullmatch(r"\d{11}", abn), name)


if __name__ == "__main__":
    unittest.main()


class FooterNameTests(unittest.TestCase):
    """A supplier that is not on the list is still caught by its legal ending, also when its name is printed low on the page."""
    TOP = "TAX INVOICE\nInvoice No 5521\nDate 02/10/2026\n" + "\n".join(f"Item {i} 10.00" for i in range(20))

    def names(self, extra):
        from invoicebot import suppliers
        return suppliers.heuristic_names(self.TOP + "\n" + extra)

    def test_a_name_next_to_the_abn_in_the_footer_is_taken(self):
        self.assertEqual(self.names("Zed Fresh Foods Pty Ltd   ABN 12 345 678 901")[0], "Zed Fresh Foods Pty Ltd")

    def test_a_remit_to_name_is_taken(self):
        self.assertEqual(self.names("Remit to: Alpha Beta Distributors Pty Ltd")[0], "Alpha Beta Distributors Pty Ltd")

    def test_a_delivery_or_plain_pty_ltd_line_is_not_taken(self):
        self.assertIsNone(self.names("Deliver to: Site Holdings Pty Ltd")[0])
        self.assertIsNone(self.names("Some Random Company Pty Ltd")[0])
        self.assertIsNone(self.names("Pearl Energy No 1 Pty Ltd   ABN 11 222 333 444")[0])


class OtherSitesAndOldNamesTests(Isolated, unittest.TestCase):
    def read(self, supplier, extra=(), subject="Tax Invoice 7788"):
        import os
        from invoicebot import extractor
        from tests import make_data
        p = os.path.join(self.home, "x.pdf")
        make_data.simple(p, "TAX INVOICE", ["Invoice No: 7788", "Invoice Date: 02/10/2026", *extra,
                                            "Total excl GST 100.00", "Total GST 10.00", "Total inc GST 110.00"], supplier=supplier)
        return extractor.analyse(p, {"subject": subject})["supplier"]

    def test_another_site_of_a_multi_site_supplier_gets_the_same_name(self):
        self.assertEqual(self.read("Peters Ice Cream Pty Ltd"), "Peters")              # a third ABN, no e-mail on the page
        self.assertEqual(self.read("Peleguy Pty Ltd"), "Peleguy")
        self.assertEqual(self.read("Nippy's Pty Ltd"), "Nippy's")

    def test_the_supplier_list_name_beats_an_older_suppliers_sheet_name(self):
        self.assertEqual(self.read("Bega Cheese Limited", ["ABN 68 083 019 390"]), "Bega")
        self.assertEqual(self.read("Lactalis Australia Pty Ltd", ["ABN 56 072 928 879"]), "Lactalis")
        self.assertEqual(self.read("Asahi Beverages", ["ABN 51 004 243 994"]), "Asahi")

    def test_an_older_workbook_gets_the_three_suppliers_rows(self):
        from openpyxl import load_workbook
        wb = load_workbook(settings.settings_file())
        ws = wb["Suppliers"]
        for r in range(ws.max_row, 1, -1):
            if str(ws.cell(r, 1).value or "").lower() in ("peters", "peleguy", "nippy"):
                ws.delete_rows(r)
        wb.save(settings.settings_file())
        settings.reset()
        added = settings.upgrade_workbook()
        self.assertEqual(sorted(a for a in added if a.startswith("Suppliers row")), ["Suppliers row 'Nippy'", "Suppliers row 'Peleguy'", "Suppliers row 'Peters'"])
