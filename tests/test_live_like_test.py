"""LIVE mode folders (5.11.0): a missing month or company folder is created, a missing SITE folder is not (the invoice waits);
an existing folder with another spelling is reused. The old 'like test mode' switch still creates everything."""
import datetime
import os
import unittest

from tests.helpers import CAREFUL_LIVE, ONE, Isolated, gateway, mail, run
from invoicebot.router import Router

D = datetime.datetime(2026, 10, 9)


class LiveLikeTestTests(Isolated, unittest.TestCase):
    def live(self, root, pattern="{MONTH_FOLDER}\\{COMPANY}\\{SITE}", **extra):
        self.set_run(Destination_pattern=pattern, Destination_mode="live", Live_destination_root=root, **extra)
        return Router()

    def test_live_creates_month_and_company_folders_but_not_site_folders_by_default(self):
        from invoicebot.router import PlacementError
        root = self.path("EMAIL INVOICES")
        os.makedirs(root)
        r = self.live(root, **CAREFUL_LIVE)
        with self.assertRaises(PlacementError) as ctx:
            r.prepare_dir(D, company="RADHE KRISHNA INVESTMENT", site="BP NYNGAN")
        self.assertIn("site folder does not exist", ctx.exception.reason)
        self.assertTrue(os.path.isdir(os.path.join(root, "OCTOBER 2026 EMAIL INVOICES", "RADHE KRISHNA INVESTMENT")))
        self.assertFalse(os.path.exists(os.path.join(root, "OCTOBER 2026 EMAIL INVOICES", "RADHE KRISHNA INVESTMENT", "BP NYNGAN")))

    def test_an_existing_site_folder_is_used(self):
        root = self.path("EMAIL INVOICES")
        real = os.path.join(root, "OCTOBER 2026 EMAIL INVOICES", "RADHE KRISHNA INVESTMENT", "BP NYNGAN")
        os.makedirs(real)
        folder = self.live(root).prepare_dir(D, company="RADHE KRISHNA INVESTMENT", site="BP NYNGAN")
        self.assertEqual(os.path.normpath(folder), os.path.normpath(real))

    def test_the_default_creates_every_folder_like_test(self):
        root = self.path("EMAIL INVOICES")
        os.makedirs(root)
        folder = self.live(root).prepare_dir(D, company="RADHE KRISHNA INVESTMENT", site="BP NYNGAN")      # 5.15.0 default
        self.assertTrue(os.path.isdir(folder))
        self.assertTrue(folder.endswith(os.path.join("OCTOBER 2026 EMAIL INVOICES", "RADHE KRISHNA INVESTMENT", "BP NYNGAN")))

    def test_the_like_test_switch_creates_every_folder(self):
        root = self.path("EMAIL INVOICES")
        os.makedirs(root)
        folder = self.live(root, Live_mode_creates_missing_folders_like_test_mode="Yes").prepare_dir(D, company="X", site="Y")
        self.assertTrue(os.path.isdir(folder))

    def test_month_and_company_switches_set_to_no_create_nothing(self):
        from invoicebot.router import PlacementError
        root = self.path("EMAIL INVOICES")
        os.makedirs(root)
        r = self.live(root, Create_month_folder_in_live_mode="No", Create_company_folder_in_live_mode="No", **CAREFUL_LIVE)
        with self.assertRaises(PlacementError):
            r.prepare_dir(D, company="X", site="Y")

    def test_an_existing_folder_with_another_spelling_is_reused_not_duplicated(self):
        root = self.path("EMAIL INVOICES")
        real = os.path.join(root, "OCTOBER 2026 EMAIL INVOICES", "RADHE KRISHNA INVESTMENT", "AMPOL  KELSO")      # two spaces, as people type it
        os.makedirs(real)
        folder = self.live(root).prepare_dir(D, company="RADHE KRISHNA INVESTMENT", site="AMPOL KELSO")
        self.assertEqual(os.path.normpath(folder), os.path.normpath(real))
        self.assertEqual(os.listdir(os.path.join(root, "OCTOBER 2026 EMAIL INVOICES", "RADHE KRISHNA INVESTMENT")), ["AMPOL  KELSO"])

    def test_a_whole_run_in_live_mode_files_the_document_and_builds_its_folders(self):
        root = self.path("EMAIL INVOICES")
        os.makedirs(root)
        self.set_run(Destination_mode="live", Live_destination_root=root)
        gw = gateway()
        gw.add(ONE, mail("Tax Invoice 778899", ["invoice_bidfood.pdf"]))
        s = run(gw)
        self.assertEqual((s.c["documents"], s.c["Unassigned"]), (1, 0))
        filed = [os.path.join(r_, f) for r_, _, fs in os.walk(root) for f in fs if f.endswith(".pdf")]
        self.assertEqual(len(filed), 1)
        self.assertIn("OCTOBER 2026 EMAIL INVOICES", filed[0])


if __name__ == "__main__":
    unittest.main()
