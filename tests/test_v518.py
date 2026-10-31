"""5.18.0: TAX INVOICE pages that also say Delivery Note / Delivery Docket stay invoices (Lactalis, PacPlus); a street number range
(102-104 Wollombi Road) matches the site (Metcash POD for Cessnock, arrived in the Shri Hari mailbox); Metcash Customer Activity
Reports are filed in the site folder; security reports with 'Client No: 9307(1)-SA' (area 1) are read (Hillcrest, Hampstead
Gardens, Torrensville, Sheidow Park - the real October reports)."""
import datetime
import glob
import os
import shutil
import tempfile
import unittest

from tests import docs_v518 as D
from tests.helpers import DATA, ONE, Isolated, connect_db, gateway, mail, run
from tests.test_site_layout import site_row

from invoicebot import db, doctypes, exports, paths, security, security_reader, security_settings as ss, settings, sites

SHRI_HARI = "shrihari@pearlenergy.com.au"
AREA1 = sorted(glob.glob(os.path.join(DATA, "security_area1_*.pdf")))
SITES = [site_row("Pearl Energy Cessnock", "102 Wollombi Road, Cessnock NSW 2325", company="Pearl Energy No 3 Pty Ltd"),
         site_row("Ampol Greta", "2021 New England Highway, Greta NSW 2334", company="Shri Hari Investment Pty Ltd"),
         site_row("Ampol Farleigh", "971 Bruce Highway, Farleigh QLD 4741", company="Pearl Energy No 2 Pty Ltd"),
         site_row("Mobil Torrensville", "239 Henley Beach Road, Torrensville SA 5031", company="Pearl Energy No 1 Pty Ltd")]


class DocTests(Isolated, unittest.TestCase):
    def setUp(self):
        super().setUp()
        self.tmp = tempfile.mkdtemp(prefix="v518_")
        self.set_run(Destination_pattern="{MONTH_FOLDER}\\{COMPANY}\\{SITE}")
        self.add_rows("Sites", SITES)

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)
        super().tearDown()

    def deliver(self, builder, subject, box=ONE, name=None):
        p = os.path.join(self.tmp, name or builder.__name__ + ".pdf")
        builder(p)
        g = gateway()
        g.add(box, mail(subject, [p]))
        run(g)
        return dict(connect_db(self.home).execute("SELECT * FROM docs").fetchone())

    def test_lactalis_tax_invoice_with_a_delivery_note_line_is_an_invoice(self):
        d = self.deliver(D.lactalis_invoice, "Lactalis Tax Invoice")
        self.assertEqual((d["doc_type"], d["number"], d["dest_name"]), ("Invoice", "244596854", "Lactalis_244596854.pdf"))

    def test_the_title_rules(self):
        h = "Lactalis Australia Pty Ltd TAX INVOICE\nDelivery Note 433201 for 6.10.2026 Purchase Order 2359665"
        self.assertEqual(doctypes.decide("", "", h, None, h)[0], "Invoice")
        h = "MACKAY CLEANING PacPlus Payne Pty Ltd Delivery Docket\nTAX INVOICE\nSOLD TO: Ampol Farleigh SHIP TO: Ampol Farleigh"
        self.assertEqual(doctypes.decide("", "", h, None, h)[0], "Invoice")
        h = "Acme Foods\nDELIVERY DOCKET\nDocket No 5566"
        self.assertEqual(doctypes.decide("", "", h, None, h)[0], "POD")                     # a real docket stays a POD

    def test_cessnock_pod_with_a_street_number_range_goes_to_pearl_3(self):
        d = self.deliver(D.metcash_pod, "Metcash POD", box=SHRI_HARI)
        self.assertEqual((d["doc_type"], d["site"], d["company_folder"], d["status"]),
                         ("POD", "Pearl Energy Cessnock", "THREE PEARL ENERGY", "Filed"))         # never the Shri Hari mailbox

    def test_ranges(self):
        self.assertIn("102 wollombi road", sites._range_variants("102 104 wollombi road")[1])
        self.assertIn("104 wollombi road", sites._range_variants("102 104 wollombi road")[2])
        self.assertEqual(sites._range_variants("unit 5 wollombi")[1:], [])

    def test_metcash_activity_report_is_filed_in_the_site_folder(self):
        d = self.deliver(D.metcash_activity_report, "Customer Activity Report MOBIL TORRENSVILLE  A/C 85852440 19-10-2026",
                         name="Statement_TransReport.85852440.1000.20261019133532.pdf")
        self.assertEqual((d["doc_type"], d["site"], d["status"]), ("Customer Activity Report", "Mobil Torrensville", "Filed"))
        self.assertTrue(d["dest_name"].endswith("_19-10-2026_CAR.pdf"), d["dest_name"])
        self.assertIn(os.path.join("OCTOBER 2026 EMAIL INVOICES", "ONE PEARL ENERGY", "MOBIL TORRENSVILLE"), d["dest_dir"])
        conn = connect_db(self.home)
        self.assertEqual([r for r in exports.ac_rows(conn, "2026-01-01", "2027-01-01")], [])      # never in the A/C Excel

    def test_older_settings_files_stop_parking_activity_reports(self):
        from openpyxl import load_workbook
        wb = load_workbook(settings.settings_file())
        for ws_name, key, col, value in (("DocTypes", "Customer Activity Report", 2, "Unrecognised"),
                                         ("Run", "Settings version", 2, "5.17.0")):
            ws = wb[ws_name]
            for r in range(2, ws.max_row + 1):
                if ws.cell(r, 1).value == key:
                    ws.cell(r, col).value = value
        wb.save(settings.settings_file())
        settings.reset()
        added = settings.upgrade_workbook()
        settings.reset()
        self.assertTrue(any("Customer Activity Report" in a for a in added))
        car = next(r for r in settings.rows("DocTypes") if r["Document type"] == "Customer Activity Report")
        self.assertEqual(car["Route"], "Company folder")


def site_rows():
    def row(no, name, client, addr):
        return ["PEARL ENERGY NO.1", no, name, client, addr] + ["5:00 AM", "11:00 PM"] * 5 + ["6:00 AM", "11:00 PM", "6:00 AM", "10:00 PM"] \
            + ["Yes", "No"]
    return [row(26, "Mobil - Hillcrest SA", "9307-SA", "473-477 North East Road Hillcrest 5086 SA"),
            row(28, "Mobil - Hampstead Gardens SA", "5652-SA", "Cnr North East Road & Muller Road Hampstead Gardens 5086 SA"),
            row(29, "Mobil - Torrensville SA", "5657-SA", "239 Henley Beach Road Torrensville 5031 SA"),
            row(30, "Mobil - Sheidow Park SA", "5679-SA", "14 Commercial Road Sheidow Park 5158 SA")]


class SecurityAreaTests(Isolated, unittest.TestCase):
    def test_the_reader_joins_area_1_to_its_site(self):
        self.assertEqual(len(AREA1), 4)
        for p in AREA1:
            secs = security_reader.parse(p)
            self.assertEqual(len(secs), 1, p)                                              # page 1 (periodic) + page 2 = one site
            self.assertRegex(secs[0]["client"], r"^\d{4}-SA$")
            self.assertGreaterEqual(len(secs[0]["events"]), 14)

    def test_the_four_sites_are_no_longer_no_report(self):
        ss.create_default_file(paths.p(ss.FILE_NAME), site_rows())
        g = gateway()
        g.add_mailbox("Security")
        for f in AREA1:
            g.add("Security", mail("Weekly report", [f], hours_ago=24))
        r = security.SecurityRun(g, log=lambda *a: None, from_date=datetime.date(2026, 9, 28), to_date=datetime.date(2026, 12, 31)).run()
        self.assertFalse(r.warnings, r.warnings)
        conf = ss.load()
        conn = connect_db(self.home)
        start, end = datetime.date(2026, 10, 12), datetime.date(2026, 10, 18)
        cover = security.coverage(conn, conf["sites"], start, end)
        for client, mon_open, mon_close in (("9307-SA", "04:59", "23:03"), ("5652-SA", "04:58", "23:04"),
                                            ("5657-SA", "04:56", "23:03"), ("5679-SA", "04:56", "23:03")):
            site = [s for s in conf["sites"] if s["client"] == client][0]
            days = security.day_results(conn, site, start, end, (10, 4, 3.5), {w: v.get(client, False) for w, v in cover.items()})
            self.assertFalse(any(x["state"] == "no report" for x in days.values()), client)
            mon = days[start]
            self.assertEqual((mon["in"].strftime("%H:%M"), mon["out"].strftime("%H:%M")), (mon_open, mon_close), client)


if __name__ == "__main__":
    unittest.main()
