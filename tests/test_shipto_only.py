"""5.11.0: ONLY the ship-to / deliver-to address decides where an invoice goes - never the bill-to.
Built on the two Unilever invoices the payments team reported (filed under the bill-to site by 5.10.0)."""
import os
import unittest

from tests.helpers import ONE, Isolated, connect_db, gateway, mail, run
from tests.test_site_layout import site_row

from invoicebot import settings, sites

TAG = sites.SHIP_TAG
PRESTON_PDF = "unilever_billto_denham_court_shipto_preston.pdf"
ORANGE_PDF = "unilever_billto_harlaxton_shipto_bp_north_orange.pdf"

DENHAM = site_row("Ampol Denham court", "194 Campbelltown Road, Denham Court NSW 2565", company="Gopinath Investment Pty Ltd")
HARLAXTON = site_row("Pearl Energy Harlaxton", "6-10 Ruthven Street, Harlaxton QLD 4350", company="Pearl Energy No 3 Pty Ltd")
ORANGE_PE = site_row("Pearl Energy Orange", "68 Bathurst Road, Orange NSW 2800", company="Pearl Energy No 2 Pty Ltd")
BP_ORANGE = ["BP Orange North", "Pearl Energy No 3 Pty Ltd", "", "5 hanharan|hanharan place orange", "", "", "Yes", "", "", ""]   # the typo in the real list
PRESTON = site_row("Pearl Energy Preston", "1B Bell Street, Preston VIC 3072", company="Pearl Energy No 5 Pty Ltd")
FILLER = [site_row(f"Ampol Town{i}", company="Gopinath Investment Pty Ltd") for i in range(3)] + \
         [site_row(f"BP Town{i}", company="Pearl Energy No 3 Pty Ltd") for i in range(3)]       # 3+ sites make a brand, as in the real list


class Base(Isolated, unittest.TestCase):
    def setUp(self):
        super().setUp()
        self.set_run(Destination_pattern="{MONTH_FOLDER}\\{COMPANY}\\{SITE}")
        self.add_rows("Sites", [DENHAM, HARLAXTON, ORANGE_PE, BP_ORANGE] + FILLER)

    def deliver(self, pdf, subject="Unilever invoice"):
        gw = gateway()
        gw.add(ONE, mail(subject, [pdf]))
        return run(gw)

    def doc(self):
        return dict(connect_db(self.home).execute("SELECT * FROM docs").fetchone())

    def filed(self):
        return [f for f in self.files_under("TestOutput") if f.endswith(".pdf")]


class ReportedInvoicesTests(Base):
    def test_bp_north_orange_goes_to_bp_orange_north_not_harlaxton(self):
        """Ship-to 'BP NORTH ORANGE, 5 HANRAHAN PLACE'; the list says 'BP Orange North' and 'hanharan' (a typo)."""
        s = self.deliver(ORANGE_PDF)
        self.assertEqual(s.c["Filed"], 1)
        self.assertEqual(len(self.filed()), 1)
        self.assertIn("THREE PEARL ENERGY/BP ORANGE NORTH/Unilever_3506011682.pdf", self.filed()[0])
        self.assertEqual(self.doc()["site"], "BP Orange North")

    def test_preston_never_goes_to_the_bill_to_site_denham_court(self):
        s = self.deliver(PRESTON_PDF)
        self.assertEqual((s.c["Filed"], s.c["Unassigned"]), (0, 1))
        self.assertEqual(self.filed(), [])                                            # NOT in GOPINATH INVESTMENT / AMPOL DENHAM COURT
        d = self.doc()
        self.assertIn("SHIP-TO NOT IN SITE LIST", d["reason"])
        self.assertIn("PEARL ENERGY PRESTON", d["reason"])
        self.assertIn("PEARL ENERGY PRESTON", d["ship_to"])
        self.assertIn("Unilever_3506013060.pdf", self.files_under("Holding", "_Unassigned"))

    def test_preston_is_filed_by_itself_once_the_site_is_added(self):
        self.deliver(PRESTON_PDF)
        self.add_rows("Sites", [PRESTON])
        s = run(gateway())
        self.assertEqual(s.c["released"], 1)
        self.assertEqual(len(self.filed()), 1)
        self.assertIn("FIVE PEARL ENERGY/PEARL ENERGY PRESTON/Unilever_3506013060.pdf", self.filed()[0])   # the SHIP-TO site's company
        d = self.doc()
        self.assertEqual((d["status"], d["site"]), ("Filed", "Pearl Energy Preston"))
        self.assertIn("site added later", d["site_source"])
        self.assertEqual(self.files_under("Holding", "_Unassigned"), [])

    def test_a_waiting_ship_to_is_retried_past_the_attempt_limit(self):
        self.set_run(Max_attempts_for_waiting_files="1")
        self.deliver(PRESTON_PDF)
        run(gateway())
        run(gateway())                                                                 # far past the limit of 1
        self.add_rows("Sites", [PRESTON])
        self.assertEqual(run(gateway()).c["released"], 1)

    def test_a_person_can_still_choose_the_site_in_to_fix(self):
        from openpyxl import load_workbook
        self.deliver(PRESTON_PDF)
        wb = load_workbook(self.path("Reports", "To_Fix.xlsx"))
        ws = wb["To fix"]
        head = [c.value for c in ws[1]]
        ws.cell(2, head.index("Site folder (TYPE HERE)") + 1).value = "PEARL ENERGY ORANGE"
        wb.save(self.path("Reports", "To_Fix.xlsx"))
        wb.close()
        s = run(gateway())
        self.assertEqual(s.c["released"], 1)
        self.assertEqual(self.doc()["site"], "Pearl Energy Orange")

    def test_the_old_rule_can_be_switched_back(self):
        self.set_run(**{"Use only the ship-to address": "No"})
        self.deliver(PRESTON_PDF)
        self.assertIn("AMPOL DENHAM COURT", self.filed()[0])                         # 5.10.0 behaviour, on request only


class MatchingTests(Base):
    BILL = "15537500 | PEARL ENERGY HARLAXTON | 6 TO 10 RUTHVEN STREET"

    def match(self, ship, bill=None, supplier="Unilever"):
        return sites.match("x", blocks=[TAG + ship, bill or self.BILL], supplier=supplier)

    def test_the_name_in_another_word_order(self):
        self.assertEqual(self.match("BP NORTH ORANGE | NSW")["site"], "BP Orange North")

    def test_a_small_street_typo_with_the_same_number(self):
        self.assertEqual(self.match("5 HANRAHAN PLACE | ORANGE NSW 2800")["site"], "BP Orange North")

    def test_a_different_street_number_is_not_accepted(self):
        self.assertEqual(self.match("7 HANRAHAN PLACE | ORANGE NSW 2800")["site"], settings.UNASSIGNED)

    def test_word_order_is_never_read_across_two_lines(self):
        """'... ORANGE | PEARL ENERGY ...' must not become 'Pearl Energy Orange'."""
        self.assertEqual(self.match("BP NORTH ORANGE | PEARL ENERGY NO.3 | 5 HANRAHAN PLACE")["site"], "BP Orange North")

    def test_an_account_number_from_the_bill_to_is_not_used(self):
        self.add_rows("Accounts", [["Unilever", "15537500", "Pearl Energy Harlaxton"]])
        self.assertEqual(self.match("1B BELL STREET | PRESTON VIC 3072")["site"], settings.UNASSIGNED)

    def test_an_account_number_inside_the_ship_to_block_is_used(self):
        self.add_rows("Accounts", [["Unilever", "15471177", "BP Orange North"]])
        self.assertEqual(self.match("15471177 | UNIT 9 | SOMEWHERE NSW 2800")["site"], "BP Orange North")

    def test_a_ship_to_with_only_an_unknown_site_name_still_waits(self):
        m = self.match("PEARL ENERGY PRESTON")                              # no address lines at all: still never the bill-to
        self.assertEqual(m["site"], settings.UNASSIGNED)
        self.assertTrue(sites.ship_hold(m["source"]))

    def test_a_stray_deliver_to_phrase_is_ignored(self):
        self.assertEqual(self.match("as instructed by the driver", bill=self.BILL + " | HARLAXTON QLD 4350")["site"], "Pearl Energy Harlaxton")

    def test_without_a_ship_to_the_bill_to_still_works(self):
        m = sites.match("x", blocks=[self.BILL + " | HARLAXTON QLD 4350"], supplier="Unilever")
        self.assertEqual(m["site"], "Pearl Energy Harlaxton")


if __name__ == "__main__":
    unittest.main()
