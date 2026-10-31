"""5.7.0, step 2 and 3: the month folder follows the invoice date within 3 months; highlights and notes; the next month's folders are built in
the last two days of the month; future-dated invoices are moved into their month; a site chosen in To_Fix > New sites moves the file."""
import datetime
import os
import unittest

from openpyxl import load_workbook

from tests import make_data
from tests.helpers import ONE, Isolated, connect_db, gateway, mail, run

NOW = datetime.datetime(2026, 10, 3, 14, 0, 0)
LAST_DAYS = datetime.datetime(2026, 10, 30, 9, 0, 0)
NOV1 = datetime.datetime(2026, 11, 1, 9, 0, 0)


class MonthTests(Isolated, unittest.TestCase):
    def inv(self, number, date, name=None):
        path = os.path.join(self.home, name or f"{number}.pdf")
        make_data.invoice(path, "Acme Foods Pty Ltd", number, date=date, net=100.0)
        return path

    def one(self, number, date, **kw):
        gw = gateway()
        gw.add(ONE, mail(f"Invoice {number}", [self.inv(number, date)]))
        run(gw, **kw)
        (r,) = [dict(x) for x in connect_db(self.home).execute("SELECT * FROM docs")]
        return r

    def folder_of(self, r):
        return os.path.relpath(r["dest_dir"], self.path("TestOutput")).split(os.sep)[0]

    def record(self, year, month):
        wb = load_workbook(self.path("Reports", f"Record_{year}-{month:02d}.xlsx"))
        ws = wb["Record"]
        rows = [[c.value for c in row] for row in ws.iter_rows(min_row=2)]
        fills = [row[0].fill.fgColor.rgb for row in ws.iter_rows(min_row=2)]
        wb.close()
        return rows, fills

    # ---------------------------------------------------------------- the window
    def test_this_month_goes_to_this_month_without_a_note(self):
        r = self.one("INV-2101", "02/10/2026")
        self.assertEqual((self.folder_of(r), r["month_case"] or "", r["remarks"] or ""), ("OCTOBER 2026 EMAIL INVOICES", "", ""))

    def test_an_older_invoice_within_three_months_goes_to_its_own_month_and_the_row_is_yellow(self):
        r = self.one("INV-2102", "15/08/2026")
        self.assertEqual((self.folder_of(r), r["month_case"]), ("AUGUST 2026 EMAIL INVOICES", "past"))
        self.assertIn("INVOICE DATED AUGUST 2026", r["remarks"])
        rows, fills = self.record(2026, 10)                         # the row stays in the month it ARRIVED
        self.assertEqual(len(rows), 1)
        self.assertTrue(fills[0].endswith("FFF2A8"))               # yellow

    def test_an_invoice_older_than_three_months_goes_to_the_current_month_and_is_lilac(self):
        r = self.one("INV-2103", "10/05/2026")
        self.assertEqual((self.folder_of(r), r["month_case"]), ("OCTOBER 2026 EMAIL INVOICES", "far"))
        self.assertIn("MORE THAN 3 MONTHS OLD", r["remarks"])
        self.assertTrue(self.record(2026, 10)[1][0].endswith("E1D5F0"))

    def test_a_future_invoice_is_filed_in_the_current_month_and_is_orange_not_unrecognised(self):
        r = self.one("INV-2104", "05/11/2026")
        self.assertEqual((r["status"], self.folder_of(r), r["month_case"], r["month_target"]), ("Filed", "OCTOBER 2026 EMAIL INVOICES", "future", "2026-11"))
        self.assertIn("FUTURE-DATED NOVEMBER 2026", r["remarks"])
        self.assertTrue(self.record(2026, 10)[1][0].endswith("F8CBAD"))

    def test_an_absurd_future_date_goes_to_the_current_month(self):
        r = self.one("INV-2105", "11/10/2099")
        self.assertEqual((r["status"], self.folder_of(r), r["month_case"]), ("Filed", "OCTOBER 2026 EMAIL INVOICES", "far"))
        self.assertIn("MORE THAN 3 MONTHS AHEAD", r["remarks"])

    def test_the_window_can_be_switched_off(self):
        self.set_run(**{"Month folder follows": "arrival"})
        r = self.one("INV-2106", "15/08/2026")
        self.assertEqual(self.folder_of(r), "OCTOBER 2026 EMAIL INVOICES")

    # ---------------------------------------------------------------- next month's folders and the move
    def make_october(self):
        os.makedirs(self.path("TestOutput", "OCTOBER 2026 EMAIL INVOICES", "ONE PEARL ENERGY", "MOBIL A"))
        os.makedirs(self.path("TestOutput", "OCTOBER 2026 EMAIL INVOICES", "TWO PEARL ENERGY"))
        with open(self.path("TestOutput", "OCTOBER 2026 EMAIL INVOICES", "TWO PEARL ENERGY", "keep.txt"), "w") as f:
            f.write("a file")

    def test_in_the_last_two_days_the_next_months_folders_are_built_folders_only(self):
        self.make_october()
        run(gateway(), now=LAST_DAYS)
        nov = self.path("TestOutput", "NOVEMBER 2026 EMAIL INVOICES")
        self.assertTrue(os.path.isdir(os.path.join(nov, "ONE PEARL ENERGY", "MOBIL A")))
        self.assertTrue(os.path.isdir(os.path.join(nov, "TWO PEARL ENERGY")))
        self.assertFalse(os.path.exists(os.path.join(nov, "TWO PEARL ENERGY", "keep.txt")))           # folders only, never files

    def test_not_in_the_middle_of_the_month(self):
        self.make_october()
        run(gateway(), now=NOW)
        self.assertFalse(os.path.exists(self.path("TestOutput", "NOVEMBER 2026 EMAIL INVOICES")))

    def test_the_first_run_of_a_new_month_builds_a_missing_folder(self):
        self.make_october()
        run(gateway(), now=NOV1)
        self.assertTrue(os.path.isdir(self.path("TestOutput", "NOVEMBER 2026 EMAIL INVOICES", "ONE PEARL ENERGY", "MOBIL A")))

    def test_it_can_be_switched_off(self):
        self.set_run(**{"Create next month's folders": "No"})
        self.make_october()
        run(gateway(), now=LAST_DAYS)
        self.assertFalse(os.path.exists(self.path("TestOutput", "NOVEMBER 2026 EMAIL INVOICES")))

    def test_a_waiting_future_invoice_is_moved_in_the_run_that_builds_its_month(self):
        r = self.one("INV-2107", "05/11/2026")
        old = r["current_path"]
        self.assertTrue(os.path.exists(old))
        run(gateway(), now=LAST_DAYS)                                            # builds NOVEMBER and moves
        (r2,) = [dict(x) for x in connect_db(self.home).execute("SELECT * FROM docs")]
        self.assertEqual((r2["month_case"], self.folder_of(r2)), ("moved", "NOVEMBER 2026 EMAIL INVOICES"))
        self.assertTrue(os.path.exists(r2["current_path"]))
        self.assertFalse(os.path.exists(old))
        self.assertIn("MOVED TO NOVEMBER 2026", r2["remarks"])
        oct_rows, oct_fills = self.record(2026, 10)                              # October's row stays, orange, note changed
        self.assertEqual(len(oct_rows), 1)
        self.assertTrue(oct_fills[0].endswith("F8CBAD"))
        self.assertIn("MOVED TO NOVEMBER 2026", oct_rows[0][11])
        nov_rows, nov_fills = self.record(2026, 11)                              # November lists it again: arrived early, light blue
        self.assertEqual(len(nov_rows), 1)
        self.assertTrue(nov_fills[0].endswith("BDD7EE"))
        self.assertIn("ARRIVED EARLY", nov_rows[0][11])

    def test_a_file_that_staff_moved_is_left_and_reported(self):
        r = self.one("INV-2108", "05/11/2026")
        os.remove(r["current_path"])                                             # someone moved it away
        s = run(gateway(), now=LAST_DAYS)
        (r2,) = [dict(x) for x in connect_db(self.home).execute("SELECT * FROM docs")]
        self.assertEqual(r2["month_case"], "future")
        self.assertIn("COULD NOT MOVE", r2["remarks"])
        self.assertTrue(any("a person must move it" in w for w in s.warnings))

    def test_the_moved_invoice_is_not_counted_twice_in_the_quarterly_archive(self):
        from invoicebot import db, quarterly
        self.one("INV-2109", "05/11/2026")
        run(gateway(), now=LAST_DAYS)
        conn = connect_db(self.home)
        source = db.DocSource(conn, *db.between_clause(datetime.date(2026, 10, 1), datetime.date(2026, 12, 31), "arrival"))
        self.assertEqual(source.count(), 1)


MONTH = "OCTOBER 2026 EMAIL INVOICES"


class NewSitesTests(Isolated, unittest.TestCase):
    def setUp(self):
        super().setUp()
        self.set_run(Destination_pattern="{MONTH_FOLDER}\\{COMPANY}\\{SITE}")
        self.add_rows("Sites", [["Mobil Ballina", "Pearl Energy No 1 Pty Ltd", "", "", "", "", "Yes", "", "", "MOBIL BALLINA"]])
        gw = gateway()
        gw.add(ONE, mail("FW: Invoice NS-4040", ["invoice_company_named.pdf"]))              # the site (Robina) is not in the list
        run(gw)

    def to_fix(self):
        return load_workbook(self.path("Reports", "To_Fix.xlsx"))

    def type_site(self, text):
        wb = self.to_fix()
        ws = wb["New sites"]
        ws.cell(2, 8).value = text
        wb.save(self.path("Reports", "To_Fix.xlsx"))
        wb.close()

    def row(self):
        return dict(connect_db(self.home).execute("SELECT * FROM docs").fetchone())

    def test_the_document_is_listed_and_every_site_by_company(self):
        wb = self.to_fix()
        self.assertIn("New sites", wb.sheetnames)
        rows = [[c for c in r] for r in wb["New sites"].iter_rows(min_row=2, values_only=True)]
        self.assertEqual(rows[0][1], "RADHE KRISHNA INVESTMENT")
        listing = [list(r) for r in wb["Sites by company"].iter_rows(min_row=2, values_only=True)]
        self.assertTrue(any(r[2] == "Mobil Ballina" and r[3] == "MOBIL BALLINA" for r in listing))
        self.assertTrue(any(r[0] == "ONE PEARL ENERGY" for r in listing))               # a heading row for each company
        wb.close()

    def test_a_site_typed_in_new_sites_moves_the_file(self):
        before = self.row()
        self.type_site("Mobil Ballina")
        run(gateway())
        r = self.row()
        self.assertEqual(r["site"], "Mobil Ballina")
        self.assertTrue(r["current_path"].replace("\\", "/").endswith(f"{MONTH}/RADHE KRISHNA INVESTMENT/MOBIL BALLINA/Nova Snacks Pty Ltd_NS-4040.pdf"))
        self.assertTrue(os.path.exists(r["current_path"]))
        self.assertFalse(os.path.exists(before["current_path"]))
        self.assertNotIn("SITE NOT FOUND", r["remarks"] or "")

    def test_a_site_that_is_not_on_the_list_creates_nothing(self):
        before = self.row()
        self.type_site("Mobil Narnia")
        run(gateway())
        r = self.row()
        self.assertEqual(r["current_path"], before["current_path"])
        self.assertTrue(os.path.exists(r["current_path"]))
        self.assertIn("SITE NAME NOT ON LIST: Mobil Narnia", r["remarks"])
        self.assertFalse(os.path.exists(os.path.join(os.path.dirname(r["current_path"]), "MOBIL NARNIA")))


if __name__ == "__main__":
    unittest.main()
