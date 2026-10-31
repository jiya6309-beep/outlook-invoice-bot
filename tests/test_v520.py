"""5.20.0 (the _Unassigned review, kept simple): site names suppliers print are added to the Sites sheet once (Mobil Port Macquarie
Hastings, Mobil Mossvale, Pearl Energy Mittagong, Pearl Energy Beresfield South, 1-10 Amy Cl Wyong); a Fleet Card statement is
placed by the merchant NAME at the top (its address is the owner's postal address); an account statement with 'Overdue / Current /
Total Amount Due' is a statement even when its title is a picture (Neverfail)."""
import os
import shutil
import tempfile
import unittest

from tests import docs_v515, docs_v520 as D
from tests.helpers import DATA, ONE, Isolated, connect_db, gateway, mail, run

from invoicebot import doctypes, settings, sites

MSTATEMENT = "mstatement@pearlenergy.com.au"
TEAM_FILE = os.path.join(DATA, "settings_payments_team_5_13.xlsx")


class ReviewTests(Isolated, unittest.TestCase):
    def setUp(self):
        super().setUp()
        shutil.copyfile(TEAM_FILE, settings.settings_file())         # the payments team's real Sites sheet, upgraded to 5.20.0
        settings.reset()
        settings.upgrade_workbook()
        settings.reset()
        self.tmp = tempfile.mkdtemp(prefix="v520_")

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)
        super().tearDown()

    def run_docs(self, items):
        g = gateway(MSTATEMENT)
        for n, (builder, subject, box, kw) in enumerate(items):
            p = os.path.join(self.tmp, f"doc{n}.pdf")
            builder(p, **kw)
            g.add(box, mail(subject, [p], hours_ago=2 + n))
        run(g)
        return {r["orig_name"]: dict(r) for r in connect_db(self.home).execute("SELECT * FROM docs")}

    def test_the_settings_now_know_these_names(self):
        sw = settings.words("MOBIL PORT MACQUARIE HASTINGS | 128 HASTINGS RIVER DRIVE | PORT MACQUARIE NSW 2444")
        self.assertEqual(sites._name_hits(sites.active_sites(), sw), {"Mobil Hasting River"})
        m = sites.match("", blocks=["[SHIP] MOBIL PORT MACQUARIE HASTINGS | 128 HASTINGS RIVER DRIVE | PORT MACQUARIE | NSW 2444"],
                        supplier="Asahi")
        self.assertEqual(m["site"], "Mobil Hasting River")                                   # the Asahi credit note
        self.assertEqual(settings.upgrade_workbook(), [])                                    # added once

    def test_the_reviewed_documents_are_filed(self):
        bega_sold = ["Account: 208062", "Trusts", "Mobil Mossvale", "The Trustee For Pearl Investment", "Fa", "225 Argyle St",
                     "MOSS VALE NSW 2577"]
        docs = self.run_docs([
            (D.pfd_mittagong, "PFD Tax Invoice", ONE, {}),
            (D.slick_chicks, "Slick Chicks invoice", ONE, {}),
            (D.neverfail_statement, "Neverfail statement", ONE, {}),
            (docs_v515.bega_invoice, "Bega invoice", ONE, {"sold_lines": bega_sold, "invoice_lines": ["Payer Account: 235097"] + bega_sold[1:]}),
        ])
        got = {k: (d["doc_type"], d["site"], d["company_folder"]) for k, d in docs.items()}
        self.assertEqual(got["doc0.pdf"], ("Invoice", "Ampol Mittagong", "P & O INVESTMENT"))
        self.assertEqual(got["doc1.pdf"], ("Invoice", "Pearl Energy Beresfield", "ONE PEARL ENERGY"))
        self.assertEqual(got["doc2.pdf"], ("Statement", "Pearl Energy North Wyong", "CENTRAL COAST ENERGY 1"))
        self.assertIn("_Statements", docs["doc2.pdf"]["dest_dir"])
        self.assertEqual(got["doc3.pdf"], ("Invoice", "Mobil Moss Vale", "ONE PEARL ENERGY"))

    def test_fleet_statements_follow_the_merchant_name(self):
        items = [(D.fleet_statement, "Merchant Payment Advice", MSTATEMENT, {"name": n, "merchant": m, "address": a}) for n, m, a in (
            ("MOBIL MOE", "292898", ["37-39 KIRK STREET", "MOE", "VIC 3825"]),
            ("PEARL ENERGY BURLEIGH HEADS", "292606", ["299 SCOTTSDALE DR", "ROBINA", "QLD 4226"]),
            ("MOBIL JENNINGS", "291463", ["144 MEADE STREET", "GLEN INNES", "NSW 2370"]),
            ("Pearl Energy Tomago", "290662", ["354 Maitland Road", "Hexham", "NSW 2322"]))]
        docs = self.run_docs(items)
        got = sorted((d["site"], d["status"]) for d in docs.values())
        self.assertEqual(got, [("Mobil Jennings", "Filed"), ("Mobil Moe", "Filed"), ("Pearl Energy Barleigh Heads", "Filed"),
                               ("Pearl Energy Tomago", "Filed")])
        self.assertTrue(all("FLEETCARD STATEMENTS" in d["dest_dir"] for d in docs.values()))


class StatementPageTests(unittest.TestCase):
    def test_aging_header_makes_a_statement_but_never_beats_tax_invoice(self):
        page = "Bill To\nPearl Energy Wyong\nOverdue  Current  Total Amount Due\nAccount Delivery Address Invoice Date Invoice Reference"
        self.assertEqual(doctypes.STATEMENT_PAGE.search(page) is not None, True)
        page2 = "Acme\nTAX INVOICE\nInvoice No 55\nOverdue  Current  Total Amount Due"
        self.assertIsNone(doctypes.TAX_INVOICE.search("Bill To Pearl Energy Wyong"))
        self.assertIsNotNone(doctypes.TAX_INVOICE.search(page2))


if __name__ == "__main__":
    unittest.main()
