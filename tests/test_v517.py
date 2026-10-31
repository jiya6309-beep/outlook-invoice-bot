"""5.17.0: a ship-to / sold-to address that names the COMPANY and a TOWN finds the company's site in that town, even when the
brand and the street are printed differently (Bega: 'BP Greta, ttf Pearl Energy No.3, 122 New England Hwy, GRETA' -> Pearl Energy
Greta, listed as 122 High Street). Never without the company: Greta also has Ampol Greta (another company)."""
import os
import shutil
import tempfile
import unittest

from tests.docs_v515 import bega_invoice
from tests.helpers import ONE, Isolated, connect_db, gateway, mail, run
from tests.test_site_layout import site_row

from invoicebot import doctypes, sites

GRETA = [site_row("Pearl Energy Greta", "122 High Street, Greta NSW 2334", company="Pearl Energy No 3 Pty Ltd"),
         site_row("Ampol Greta", "2021 New England Highway, Greta NSW 2334", company="Shri Hari Investment Pty Ltd"),
         site_row("Pearl Energy Kurri", "1 Main Rd, Kurri Kurri NSW 2327", company="Pearl Energy No 3 Pty Ltd")]
SOLD = ["Account: 195634", "BP Greta", "ttf Pearl Energy No.3 FUT", "Att: Gurjeet Singh", "122 New England Hwy", "GRETA NSW 2334"]


class CompanyTownTests(Isolated, unittest.TestCase):
    def setUp(self):
        super().setUp()
        self.tmp = tempfile.mkdtemp(prefix="v517_")
        self.set_run(Destination_pattern="{MONTH_FOLDER}\\{COMPANY}\\{SITE}")
        self.add_rows("Sites", GRETA)

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)
        super().tearDown()

    def test_bega_adjustment_note_goes_to_pearl_energy_greta(self):
        p = os.path.join(self.tmp, "bega_cn.pdf")
        bega_invoice(p, sold_lines=SOLD, invoice_lines=["Payer Account: 235093"] + SOLD[1:])
        g = gateway()
        g.add(ONE, mail("Bega credit note", [p]))
        run(g)
        d = dict(connect_db(self.home).execute("SELECT * FROM docs").fetchone())
        self.assertEqual((d["site"], d["company_folder"], d["status"]), ("Pearl Energy Greta", "THREE PEARL ENERGY", "Filed"))
        self.assertIn("company + town", d["site_source"])

    def test_the_title_adjustment_note_credit_is_a_credit_note(self):
        header = "ADJUSTMENT BDD Australia Pty Ltd\nNOTE - CREDIT ABN: 68 083 019 390\nInvoice No. 927377662"
        self.assertEqual(doctypes.decide("", "", header, None, header)[0], "Credit note")

    def test_a_town_without_the_company_never_decides(self):
        m = sites.match("", blocks=["[SOLD] BP Greta | 122 New England Hwy | GRETA NSW 2334"], supplier="Bega")
        self.assertEqual(m["site"], "UNASSIGNED")                            # Pearl Energy Greta or Ampol Greta: a person decides

    def test_the_company_must_have_exactly_one_site_in_that_town(self):
        two = [{"name": "A Greta", "place": "greta"}, {"name": "B Greta", "place": "greta"}]
        self.assertEqual(sites._town_hit(two, "BP Greta | GRETA NSW 2334"), set())
        self.assertEqual(sites._town_hit(two[:1], "BP Greta | GRETA NSW 2334"), {"A Greta"})

    def test_a_ship_to_with_company_and_town_is_found_too(self):
        m = sites.match("", blocks=["[SHIP] BP Greta | ttf Pearl Energy No.3 FUT | 122 New England Hwy | GRETA NSW 2334"], supplier="Bega")
        self.assertEqual(m["site"], "Pearl Energy Greta")
        self.assertTrue(m["source"].startswith(sites.SHIP_SOURCE))


if __name__ == "__main__":
    unittest.main()
