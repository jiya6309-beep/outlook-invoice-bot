"""The bot finds the synced EMAIL INVOICES folder by itself: no path is typed anywhere."""
import os
import unittest

from tests.helpers import ONE, Isolated, connect_db, gateway, mail, run

from invoicebot import autofind, panel, settings
from tests.test_site_layout import COLYTON

LIBRARY = ("pearlenergy.com.au", "Patel Group - Documents", "OFFICE INDIA")


class FinderTests(Isolated, unittest.TestCase):
    def tree(self, *parts):
        path = os.path.join(self.profile, *parts)
        os.makedirs(path, exist_ok=True)
        return path

    def test_finds_the_folder_inside_the_synced_library(self):
        want = self.tree(*LIBRARY, "EMAIL INVOICES")
        self.tree(*LIBRARY, "MAIL INVOICES", "SEPTEMBER 2026 EMAIL INVOICES")            # the other tree: must not match
        self.tree("Documents", "EMAIL INVOICES")                                          # same name, wrong parent: must not match
        self.tree("AppData", "Local", "x", "OFFICE INDIA", "EMAIL INVOICES")              # inside a skipped folder
        self.assertEqual(autofind.find(), [want])

    def test_onedrive_style_locations_and_depth(self):
        want = self.tree("OneDrive - Pearl Energy", "pearlenergy.com.au", "Patel Group - Documents", "OFFICE INDIA", "EMAIL INVOICES")
        self.assertEqual(autofind.find(), [want])
        self.assertEqual(autofind.find(max_depth=3), [])                                  # too deep for the limit: not found, no crash

    def test_several_matches_shortest_first(self):
        near = self.tree("lib", "OFFICE INDIA", "EMAIL INVOICES")
        far = self.tree("a", "b", "c", "OFFICE INDIA", "EMAIL INVOICES")
        self.assertEqual(autofind.find(), [near, far])

    def test_the_names_are_settings(self):
        self.tree("lib", "HEAD OFFICE", "INVOICE SITES")
        self.assertEqual(autofind.find(), [])
        self.set_run(**{"Auto-find_folder_name": "INVOICE SITES", "Auto-find_parent_folder": "HEAD OFFICE"})
        self.assertEqual(len(autofind.find()), 1)
        self.set_run(**{"Auto-find_parent_folder": " "})
        self.assertEqual(len(autofind.find(parent="")), 1)                                # no parent required

    def test_remembered_only_while_it_exists(self):
        want = self.tree(*LIBRARY, "EMAIL INVOICES")
        self.assertEqual(autofind.cached(), "")
        found, hits = autofind.discover()
        self.assertEqual((found, hits), (want, [want]))
        autofind.session_root = ""
        self.assertEqual(autofind.cached(), want)                                         # read back from data/destination.txt
        os.rmdir(want)
        self.assertEqual(autofind.cached(), "")                                           # folder gone: forgotten

    def test_a_search_with_nothing_there_is_quiet(self):
        self.assertEqual(autofind.discover(), ("", []))


def no_reports(files):
    """5.15.0: in LIVE the Reports folder sits inside the destination folder; these tests look at the filed documents."""
    return [f for f in files if not f.startswith("Reports/")]


class AutomaticRunTests(Isolated, unittest.TestCase):
    def setUp(self):
        super().setUp()
        self.dest = os.path.join(self.profile, *LIBRARY, "EMAIL INVOICES")
        os.makedirs(os.path.join(self.dest, "MOBIL COLYTON"))
        self.set_run(Destination_mode="live", Destination_pattern="{SITE}")                # NO path anywhere
        self.add_rows("Sites", [COLYTON])

    def test_a_live_run_with_no_path_finds_the_folder_and_files_there(self):
        gw = gateway()
        gw.add(ONE, mail("Tax Invoice INV-1001", ["invoice_acme.pdf"]))
        s = run(gw)
        self.assertEqual(s.c["Filed"], 1)
        self.assertEqual(no_reports(self.files_under("profile", *LIBRARY, "EMAIL INVOICES")), ["MOBIL COLYTON/Acme Foods Pty Ltd_INV-1001.pdf"])
        self.assertIn("Reports/Record_2026-10.xlsx", self.files_under("profile", *LIBRARY, "EMAIL INVOICES"))   # Reports inside it
        with open(self.path("data", "destination.txt"), encoding="utf-8") as f:
            self.assertEqual(f.read(), self.dest)                                         # remembered for the next run

    def test_a_rehearsal_finds_it_without_writing_anything(self):
        gw = gateway()
        gw.add(ONE, mail("Tax Invoice INV-1001", ["invoice_acme.pdf"]))
        s = run(gw, shadow=True)
        self.assertEqual(s.c["Filed"], 1)
        self.assertFalse(os.path.exists(self.path("data", "destination.txt")))
        self.assertEqual(no_reports(self.files_under("profile", *LIBRARY, "EMAIL INVOICES")), [])    # only the rehearsal report

    def test_a_typed_path_always_wins(self):
        other = os.path.join(self.home, "Elsewhere")
        os.makedirs(os.path.join(other, "MOBIL COLYTON"))
        self.set_run(Live_destination_root=other)
        gw = gateway()
        gw.add(ONE, mail("Tax Invoice INV-1001", ["invoice_acme.pdf"]))
        run(gw)
        self.assertEqual(no_reports(self.files_under("Elsewhere")), ["MOBIL COLYTON/Acme Foods Pty Ltd_INV-1001.pdf"])

    def test_not_found_stops_with_a_clear_message(self):
        os.rmdir(os.path.join(self.dest, "MOBIL COLYTON"))
        os.rmdir(self.dest)
        gw = gateway()
        gw.add(ONE, mail("Tax Invoice INV-1001", ["invoice_acme.pdf"]))
        lines = []
        s = run(gw, log=lambda t="", **k: lines.append(t))
        self.assertEqual((s.stop_reason, gw.scan_calls), ("setup incomplete", 0))
        self.assertTrue(any("was not found" in line for line in lines))

    def test_switching_the_search_off(self):
        self.set_run(Find_the_destination_folder_automatically="No")
        gw = gateway()
        s = run(gw)
        self.assertEqual(s.stop_reason, "setup incomplete")

    def test_the_whole_chain_company_then_site_into_the_found_folder(self):
        """EMAIL INVOICES (found) -> company (named on the document) -> site of that company -> folder."""
        self.add_rows("Sites", [["Mobil Robina", "Radhe Krishna Investment Pty Ltd", "", "", "", "", "Yes", "", "", "MOBIL ROBINA"],
                                ["Mobil Dubbo", "Radhe Krishna Investment Pty Ltd", "", "", "", "", "Yes", "", "", ""],
                                ["Mobil Moss Vale", "Radhe Krishna Investment Pty Ltd", "", "", "", "", "Yes", "", "", ""]])
        os.makedirs(os.path.join(self.dest, "MOBIL ROBINA"))
        gw = gateway()
        gw.add(ONE, mail("FW: Invoice NS-4040", ["invoice_company_named.pdf"]))           # RADHE KRISHNA INVESTMENT / PEARL ENERGY ROBINA
        gw.add(ONE, mail("Tax Invoice INV-1001", ["invoice_acme.pdf"]))                    # the Colyton address
        s = run(gw)
        self.assertEqual(s.c["Filed"], 2)
        self.assertEqual(sorted(no_reports(self.files_under("profile", *LIBRARY, "EMAIL INVOICES"))),
                         ["MOBIL COLYTON/Acme Foods Pty Ltd_INV-1001.pdf", "MOBIL ROBINA/Nova Snacks Pty Ltd_NS-4040.pdf"])
        rows = {r["number"]: (r["company_folder"], r["site"]) for r in connect_db(self.home).execute("SELECT * FROM docs")}
        self.assertEqual(rows["NS-4040"], ("RADHE KRISHNA INVESTMENT", "Mobil Robina"))


class PanelAutomaticTests(Isolated, unittest.TestCase):
    def test_an_empty_box_means_find_it(self):
        c = panel.Controller(opener=lambda p: None)
        self.assertFalse(c.set_destination("live", "", "site")["ok"])                       # nothing synced yet: a clear message
        dest = os.path.join(self.profile, *LIBRARY, "EMAIL INVOICES")
        os.makedirs(dest)
        r = c.set_destination("live", "", "site")
        self.assertTrue(r["ok"], r)
        self.assertIn("found automatically", r["message"])
        settings.reset()
        self.assertEqual((settings.destination_mode(), settings.destination_root(), settings.run_value("Live destination root")),
                         ("live", dest, settings.LIVE_ROOT))             # the typed default is kept; the folder found on this PC is used
        self.assertEqual(c.overview()["root"], dest)


if __name__ == "__main__":
    unittest.main()
