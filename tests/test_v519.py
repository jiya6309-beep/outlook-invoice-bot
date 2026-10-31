"""5.19.0: Fleet Card statements and Pearl Investment fuel invoices are filed in their own folders WHICHEVER mailbox they arrive in
(some suppliers still send them to company mailboxes; Jiya@ reads every e-mail). The same one from two mailboxes is saved once.
An ordinary supplier invoice is never taken for a fuel invoice."""
import os
import shutil
import unittest

from tests.helpers import DATA, ONE, Isolated, connect_db, gateway, mail, run
from tests.test_site_layout import site_row

from invoicebot import db, fuel, mailboxes, settings

JIYA = "Jiya@pearlenergy.com.au"
MSTATEMENT = "mstatement@pearlenergy.com.au"
FLEET = "fleetcard_ampol_farleigh_01_sep_26.pdf"
FUEL = "fuel_INV-114870.pdf"


class AnyMailboxTests(Isolated, unittest.TestCase):
    def setUp(self):
        super().setUp()
        self.add_rows("Sites", [site_row("Ampol Farleigh", "971 Bruce Highway, Farleigh QLD 4741", company="Pearl Energy No 2 Pty Ltd")])

    def docs(self):
        return [dict(r) for r in connect_db(self.home).execute("SELECT * FROM docs ORDER BY id")]

    def test_a_fleet_statement_in_a_company_mailbox_goes_to_the_fleet_folder(self):
        g = gateway()
        g.add(ONE, mail("Merchant Payment Advice", [FLEET]))
        run(g)
        d = self.docs()[0]
        self.assertEqual((d["doc_type"], d["status"], d["dest_name"]), ("Fleet Card Statement", "Filed", "01-09-2026_STM.pdf"))
        self.assertTrue(d["dest_dir"].endswith(os.path.join("FLEETCARD STATEMENTS", "AMPOL FARLEIGH")), d["dest_dir"])
        self.assertNotIn("EMAIL INVOICES", d["dest_dir"])

    def test_jiya_reads_fleet_and_fuel_into_their_own_folders(self):
        mailboxes.apply([{"outlook": JIYA, "action": "special", "role": "read all"}])
        settings.reset()
        g = gateway(JIYA)
        g.add(JIYA, mail("Document", [FLEET]))
        g.add(JIYA, mail("Document", [FUEL], hours_ago=3))
        run(g)
        by_type = {d["doc_type"]: d for d in self.docs()}
        self.assertEqual(by_type["Fleet Card Statement"]["status"], "Filed")
        self.assertIn("FLEETCARD STATEMENTS", by_type["Fleet Card Statement"]["dest_dir"])
        fuel_doc = by_type["Fuel Invoice"]
        self.assertEqual((fuel_doc["status"], fuel_doc["number"]), ("Filed", "INV-114870"))
        self.assertIn("FUEL INVOICES", fuel_doc["dest_dir"])

    def test_the_same_statement_from_two_mailboxes_is_saved_once(self):
        g = gateway(MSTATEMENT)
        g.add(MSTATEMENT, mail("Merchant Payment Advice", [FLEET], hours_ago=5))
        g.add(ONE, mail("Merchant Payment Advice", [FLEET], hours_ago=2))
        s = run(g)
        self.assertEqual(s.c["duplicates"], 1)
        self.assertEqual(sum(1 for d in self.docs() if d["doc_type"] == "Fleet Card Statement"), 1)

    def test_a_fleet_duplicate_is_found_by_merchant_and_date_even_when_the_file_differs(self):
        conn = connect_db(self.home)
        conn.execute("INSERT INTO docs(ref, doc_type, ship_to, doc_date, status) VALUES('X1', 'Fleet Card Statement', 'merchant:299390', "
                     "'2026-10-20', 'Filed')")
        conn.commit()
        self.assertEqual(db.find_fleet_duplicate(conn, "merchant:299390", "2026-10-20")["ref"], "X1")
        self.assertIsNone(db.find_fleet_duplicate(conn, "merchant:299390", "2026-10-27"))

    def test_an_ordinary_invoice_is_never_a_fuel_invoice(self):
        self.assertIsNone(fuel.parse_if_fuel(os.path.join(DATA, "invoice_acme.pdf")))
        self.assertIsNotNone(fuel.parse_if_fuel(os.path.join(DATA, FUEL)))
        g = gateway()
        g.add(ONE, mail("Tax Invoice INV-1001", ["invoice_acme.pdf"]))
        run(g)
        self.assertEqual(self.docs()[0]["doc_type"], "Invoice")

    def test_live_paths_are_unchanged(self):
        root = self.path("EMAIL INVOICES")
        os.makedirs(root)
        self.set_run(Destination_mode="live", Live_destination_root=root)
        self.assertEqual(settings.holding_root(), os.path.join(root, "HOLDINGS"))
        self.assertEqual(settings.reports_root(), os.path.join(root, "Reports"))


if __name__ == "__main__":
    unittest.main()
