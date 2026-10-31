"""The SHIP-TO / DELIVER-TO address wins over the bill-to (invoice-to) address - for the site AND the company."""
import os
import unittest

from tests.helpers import DATA, ONE, Isolated, gateway, mail, run
from tests.test_site_layout import site_row

from invoicebot import settings, sites

TAG = sites.SHIP_TAG
MOSS = site_row("Mobil Moss Vale", "225 Argyle Street, Moss Vale NSW 2577")
ROBINA = site_row("Pearl Energy Robina", "299 Scottsdale Drive, Robina QLD 4226")
MITT = site_row("Mobil Mittagong", "224 Old Hume Hwy, Mittagong NSW 2575", company="Radhe Krishna Investment Pty Ltd")
NYN = site_row("Bp Nyngan", "124 Nymagee St, Nyngan NSW 2825", company="Radhe Krishna Investment Pty Ltd")
POE = site_row("Ampol Bowral", "10 Bowral Road, Bowral NSW 2576", company="P & O Investment Pty Ltd")
TEXT = "Invoice to MOBIL MOSS VALE 225 ARGYLE STREET\nDeliver to PEARL ENERGY ROBINA 299 SCOTTSDALE DRIVE"
BILL = "15407772 | MOBIL MOSS VALE | PEARL ENERGY NO.1 | 225 ARGYLE STREET"
SHIP = "18010240 | PEARL ENERGY ROBINA | PEARL ENERGY NO.1 | 299 SCOTTSDALE DRIVE"


class ShipToMatchTests(Isolated, unittest.TestCase):
    def setUp(self):
        super().setUp()
        self.add_rows("Sites", [MOSS, ROBINA])

    def test_ship_to_wins_when_both_addresses_are_on_the_page(self):
        m = sites.match(TEXT, blocks=[TAG + SHIP, BILL])
        self.assertEqual(m["site"], "Pearl Energy Robina")
        self.assertTrue(m["source"].startswith(sites.SHIP_SOURCE))

    def test_the_order_of_the_blocks_does_not_matter(self):
        self.assertEqual(sites.match(TEXT, blocks=[BILL, TAG + SHIP])["site"], "Pearl Energy Robina")

    def test_only_a_bill_to_block_is_used_as_before(self):
        self.assertEqual(sites.match("Invoice to MOBIL MOSS VALE 225 ARGYLE STREET", blocks=[BILL])["site"], "Mobil Moss Vale")

    def test_a_ship_to_block_that_is_not_an_address_is_ignored(self):
        m = sites.match("Invoice to MOBIL MOSS VALE 225 ARGYLE STREET\nDeliver to Some Other Place", blocks=[TAG + "Some Other Place", BILL])
        self.assertEqual(m["site"], "Mobil Moss Vale")                       # no street, state or postcode: not a ship-to ADDRESS

    def test_an_unknown_ship_to_address_never_falls_back_to_the_bill_to(self):
        m = sites.match(TEXT, blocks=[TAG + "1B BELL STREET | PRESTON VIC 3072", BILL])
        self.assertEqual(m["site"], settings.UNASSIGNED)
        self.assertTrue(sites.ship_hold(m["source"]))

    def test_the_old_fallback_can_be_switched_back_on(self):
        self.set_run(**{"Use only the ship-to address": "No"})
        m = sites.match(TEXT, blocks=[TAG + "1B BELL STREET | PRESTON VIC 3072", BILL])
        self.assertEqual(m["site"], "Mobil Moss Vale")

    def test_the_switch_restores_the_old_rule(self):
        self.set_run(**{"Ship-to address wins over bill-to": "No"})
        m = sites.match(TEXT, blocks=[TAG + SHIP, BILL])
        self.assertEqual(m["site"], settings.UNASSIGNED)                      # two sites on the page: a person decides, as before

    def test_two_sites_inside_the_ship_to_block_itself_are_never_guessed(self):
        both = TAG + "MOBIL MOSS VALE 225 ARGYLE STREET | PEARL ENERGY ROBINA 299 SCOTTSDALE DRIVE"
        self.assertEqual(sites.match(TEXT, blocks=[both, BILL])["site"], settings.UNASSIGNED)


class ShipToDocumentTests(Isolated, unittest.TestCase):
    """Real PDFs in the Unilever layout (two address columns side by side)."""

    def setUp(self):
        super().setUp()
        self.add_rows("Sites", [MOSS, ROBINA, MITT, NYN, POE])

    def read(self, name):
        from invoicebot import extractor
        res, _ = extractor.extract_document(os.path.join(DATA, name), subject="Unilever invoice")
        return res

    def test_robina_not_moss_vale(self):
        r = self.read("unilever_billto_moss_vale_shipto_robina.pdf")
        self.assertEqual(r["site"], "Pearl Energy Robina")
        self.assertTrue(r["site_source"].startswith(sites.SHIP_SOURCE))

    def test_nyngan_not_mittagong(self):
        r = self.read("unilever_billto_mittagong_shipto_nyngan.pdf")
        self.assertEqual(r["site"], "Bp Nyngan")

    def test_the_address_text_shown_in_the_records_has_no_tag(self):
        self.assertNotIn("[SHIP]", self.read("unilever_billto_mittagong_shipto_nyngan.pdf")["address_text"])

    def test_the_ship_to_company_wins_over_the_bill_to_company(self):
        """Billed to a Pearl Energy No 1 site (Moss Vale), delivered to Bp Nyngan, which belongs to Radhe Krishna: the invoice follows the delivery."""
        gw = gateway()
        gw.add(ONE, mail("Unilever invoice 3506008099", ["unilever_billto_moss_vale_shipto_nyngan.pdf"]))
        s = run(gw)
        self.assertEqual((s.c["documents"], s.c["Unassigned"]), (1, 0))
        filed = []
        for root, _, files in os.walk(self.path("TestOutput")):
            filed += [os.path.join(root, f) for f in files]
        self.assertEqual(len(filed), 1)
        self.assertIn("RADHE KRISHNA INVESTMENT", filed[0])
        from tests.helpers import connect_db
        self.assertEqual(connect_db(self.home).execute("SELECT site FROM docs").fetchone()["site"], "Bp Nyngan")


if __name__ == "__main__":
    unittest.main()
