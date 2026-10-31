"""Site folders: <root>\\<SITE FOLDER>\\file - the layout setting, the real folder names, matching, and the To_Fix site choice."""
import datetime
import os
import subprocess
import sys
import unittest

from openpyxl import load_workbook

from tests.helpers import ONE, ROOT, Isolated, connect_db, gateway, mail, run
from tests import make_data

from invoicebot import panel, settings, sites
from invoicebot.router import ConfigError, PlacementError, Router

D = datetime.date(2026, 10, 3)


def site_row(name, address=None, folder="", company="Pearl Energy No 1 Pty Ltd"):
    frags = "|".join(sites.make_fragments(address)) if address else ""
    return [name, company, "", frags, "", "", "Yes", "", "", folder]


COLYTON = site_row("Mobil Colyton", "88-90 Great Western Hwy, Colyton NSW 2760", "MOBIL COLYTON")


class LayoutTests(Isolated, unittest.TestCase):
    def router(self, pattern, **run_settings):
        self.set_run(Destination_pattern=pattern, **run_settings)
        return Router()

    def test_rendering(self):
        self.assertEqual(self.router("{SITE}").render(D, "COLYTON FUELS", "MOBIL COLYTON"), ["MOBIL COLYTON"])
        self.assertEqual(self.router("{COMPANY}\\{SITE}").render(D, "COLYTON FUELS", "MOBIL COLYTON"), ["COLYTON FUELS", "MOBIL COLYTON"])
        self.assertEqual(self.router("{SITE}\\{MONTH} {YEAR}").render(D, "", "AMPOL  KELSO"), ["AMPOL  KELSO", "OCTOBER 2026"])
        self.assertEqual(self.router("{MONTH_FOLDER}/{COMPANY}").render(D, "P & O INVESTMENT"), ["OCTOBER 2026 EMAIL INVOICES", "P & O INVESTMENT"])
        self.assertEqual(self.router("{SITE}").needs(), {"site"})
        self.assertEqual(self.router("{COMPANY}\\{SITE}").needs(), {"company", "site"})

    def test_a_bad_layout_stops_the_run_with_a_message(self):
        for bad in ("{SITES}", "{SITE}\\{FOO}"):
            with self.assertRaises(ConfigError) as cm:
                self.router(bad).check_config()
            self.assertIn("unknown word", str(cm.exception))
        self.set_run(Destination_pattern="{SITE}")
        Router().check_config()                                                  # a good one passes

    def test_test_mode_builds_the_site_folder(self):
        r = self.router("{SITE}")
        folder = r.prepare_dir(D, site="MOBIL COLYTON")
        self.assertTrue(folder.endswith(os.path.join("TestOutput", "MOBIL COLYTON")))
        self.assertTrue(os.path.isdir(folder))

    def test_live_mode_never_creates_site_folders_by_default(self):
        root = self.path("EMAIL INVOICES")
        os.makedirs(os.path.join(root, "AMPOL KELSO"))
        r = self.router("{SITE}", Destination_mode="live", Live_destination_root=root, Live_mode_creates_missing_folders_like_test_mode="No", Create_site_folder_in_live_mode="No")
        self.assertEqual(os.path.basename(r.prepare_dir(D, site="AMPOL KELSO")), "AMPOL KELSO")
        with self.assertRaises(PlacementError) as cm:
            r.prepare_dir(D, site="MOBIL NOWHERE")
        self.assertIn("site folder does not exist", cm.exception.reason)
        self.assertFalse(os.path.exists(os.path.join(root, "MOBIL NOWHERE")))
        r = self.router("{SITE}", Destination_mode="live", Live_destination_root=root, Live_mode_creates_missing_folders_like_test_mode="No", Create_site_folder_in_live_mode="Yes")
        self.assertTrue(os.path.isdir(r.prepare_dir(D, site="MOBIL NOWHERE")))

    def test_each_level_of_a_longer_layout_has_its_own_rule(self):
        root = self.path("SP")
        os.makedirs(os.path.join(root, "COLYTON FUELS"))
        r = self.router("{COMPANY}\\{SITE}", Destination_mode="live", Live_destination_root=root, Live_mode_creates_missing_folders_like_test_mode="No",
                        Create_company_folder_in_live_mode="No", Create_site_folder_in_live_mode="No")
        with self.assertRaises(PlacementError) as cm:
            r.prepare_dir(D, company="COLYTON FUELS", site="MOBIL COLYTON")
        self.assertIn("site folder", cm.exception.reason)
        with self.assertRaises(PlacementError) as cm:
            r.prepare_dir(D, company="HIGHLAND ENERGY 1", site="X")
        self.assertIn("company folder", cm.exception.reason)

    def test_the_month_warning_only_applies_to_a_layout_that_starts_with_the_month(self):
        root = self.path("SP")
        os.makedirs(root)
        self.assertIsNone(self.router("{SITE}", Destination_mode="live", Live_destination_root=root).month_warning(datetime.date(2026, 10, 30)))
        self.assertIsNotNone(self.router("{MONTH_FOLDER}\\{COMPANY}", Destination_mode="live", Live_destination_root=root).month_warning(datetime.date(2026, 10, 30)))


class SiteFolderTests(Isolated, unittest.TestCase):
    def test_default_folder_name_and_the_recorded_one(self):
        self.add_rows("Sites", [site_row("Ampol Kelso", None, "AMPOL  KELSO"), site_row("Mobil Dubbo")])
        self.assertEqual(sites.folder_for("Ampol Kelso"), "AMPOL  KELSO")              # the real name, double space kept
        self.assertEqual(sites.folder_for("Mobil Dubbo"), "MOBIL DUBBO")                # derived
        self.assertEqual(sites.find_site("ampol kelso")["name"], "Ampol Kelso")
        self.assertEqual(sites.find_site("AMPOL KELSO")["name"], "Ampol Kelso")
        self.assertIsNone(sites.find_site("nothing"))

    def test_a_town_with_two_sites_never_resolves_by_the_generic_name_alone(self):
        """Found by a 2,500-document test: 'PEARL ENERGY WAGGA WAGGA' with no address went to the site literally called that,
        although 'Mobil Wagga Wagga' is in the same town and company."""
        self.add_rows("Sites", [site_row("Mobil Wagga Wagga", "10 Docker Street, Wagga Wagga NSW 2650"), site_row("Pearl Energy Wagga wagga", "20 Baylis Street, Wagga Wagga NSW 2650"),
                                site_row("Mobil Dubbo"), site_row("Mobil Ballina"), site_row("Mobil Orange")])
        m = sites.match("Deliver To\nPEARL ENERGY WAGGA WAGGA", blocks=["PEARL ENERGY WAGGA WAGGA"])
        self.assertEqual(m["site"], settings.UNASSIGNED)
        self.assertIn("more than one site", m["source"])
        self.assertEqual(sites.match("Deliver To\nMOBIL WAGGA WAGGA", blocks=["MOBIL WAGGA WAGGA"])["site"], "Mobil Wagga Wagga")           # the brand name is specific
        m = sites.match("Deliver To\nPEARL ENERGY WAGGA WAGGA\n20 BAYLIS STREET WAGGA WAGGA NSW 2650", blocks=["PEARL ENERGY WAGGA WAGGA 20 BAYLIS STREET WAGGA WAGGA NSW 2650"])
        self.assertEqual(m["site"], "Pearl Energy Wagga wagga")                                                                            # an address settles it

    def test_a_site_also_answers_to_pearl_energy_and_its_town_when_the_town_is_unique(self):
        self.add_rows("Sites", [site_row("Mobil Colyton"), site_row("Mobil Dubbo"), site_row("Mobil Ballina"), site_row("Ampol Greta"),
                                site_row("Pearl Energy Greta"), site_row("Ampol Nyngan"), site_row("Bp Nyngan"), site_row("Ampol Kelso"), site_row("Bp Orange"),
                                site_row("Bp Quirindi")])                                               # 3+ sites make a brand, as in the real list
        find = lambda text: sites.match(text)["site"]
        self.assertEqual(find("Delivered to PEARL ENERGY COLYTON"), "Mobil Colyton")
        self.assertEqual(find("Statement for Pearl Energy Ballina"), "Mobil Ballina")
        self.assertEqual(find("PEARL ENERGY NYNGAN"), settings.UNASSIGNED)             # two sites in Nyngan: never a guess
        self.assertEqual(find("Pearl Energy Greta"), settings.UNASSIGNED)              # two Gretas and no address: wait, even though one is literally called that
        self.assertEqual(find("Delivered to AMPOL GRETA"), "Ampol Greta")              # the brand name is specific: accepted
        self.assertEqual(find("Colyton"), settings.UNASSIGNED)                          # a town alone is not enough

    def real_folders(self):
        root = self.path("EMAIL INVOICES")
        for f in ["AMPOL  KELSO", "AMPOL GRETA", "BP NYNGAN", "MOBIL BROKEN HILL", "AMPOL NYNGAN", "MOBIL CHILDERS", "RANDOM STUFF"]:
            os.makedirs(os.path.join(root, f))
        return root

    def test_matching_the_site_list_to_the_real_folders(self):
        self.add_rows("Sites", [site_row("Ampol Kelso"), site_row("Ampol Greta"), site_row("Bp Nyngan"), site_row("Mobil Broken Hills"),
                                site_row("Mobil Dubbo")])
        res = sites.match_folders(self.real_folders())
        self.assertEqual(res["matched"], {"Ampol Kelso": "AMPOL  KELSO", "Ampol Greta": "AMPOL GRETA", "Bp Nyngan": "BP NYNGAN"})
        self.assertEqual(list(res["fuzzy"]), ["Mobil Broken Hills"])                      # the site list has a typo; the folder is MOBIL BROKEN HILL
        self.assertEqual(res["fuzzy"]["Mobil Broken Hills"][0], "MOBIL BROKEN HILL")
        self.assertEqual(res["missing"], ["Mobil Dubbo"])
        self.assertEqual(res["extra"], ["AMPOL NYNGAN", "MOBIL CHILDERS", "RANDOM STUFF"])

    def test_matching_writes_the_folder_names_and_keeps_manual_ones(self):
        self.add_rows("Sites", [site_row("Ampol Kelso"), site_row("Ampol Greta", None, "AMPOL GRETA"), site_row("Bp Nyngan", None, "AMPOL NYNGAN")])
        root = self.real_folders()
        res = sites.match_folders(root)
        written = sites.apply_folder_matches(res, root)
        settings.reset()
        self.assertEqual(sites.folder_for("Ampol Kelso"), "AMPOL  KELSO")
        self.assertEqual(sites.folder_for("Bp Nyngan"), "AMPOL NYNGAN")                    # typed by hand and the folder exists: kept
        self.assertEqual(written, 1)                                                      # only Kelso was new

    def test_the_command_reports_in_plain_words(self):
        self.add_rows("Sites", [site_row("Ampol Kelso"), site_row("Mobil Dubbo")])
        out = subprocess.run([sys.executable, os.path.join(ROOT, "run_bot.py"), "--match-folders", self.real_folders()],
                             capture_output=True, text=True, timeout=120, cwd=ROOT)
        self.assertEqual(out.returncode, 0, out.stderr[-400:])
        self.assertIn("1 of 2 sites matched", out.stdout)
        self.assertIn("Sites WITHOUT a folder", out.stdout)
        self.assertIn("Mobil Dubbo", out.stdout)
        self.assertIn("Folders WITHOUT a site", out.stdout)
        self.assertIn("MOBIL CHILDERS", out.stdout)
        bad = subprocess.run([sys.executable, os.path.join(ROOT, "run_bot.py"), "--match-folders", self.path("nope")], capture_output=True, text=True, cwd=ROOT)
        self.assertEqual(bad.returncode, 2)
        self.assertIn("not found", bad.stdout)


class SiteRoutingTests(Isolated, unittest.TestCase):
    def setUp(self):
        super().setUp()
        self.set_run(Destination_pattern="{SITE}")
        self.add_rows("Sites", [COLYTON, site_row("Mobil Dubbo", None, "MOBIL DUBBO"), site_row("Mobil Moss Vale"), site_row("Mobil Hampstead")])   # several Mobil sites, like the real list; no Robina

    def test_invoice_credit_note_and_pod_go_to_the_site_folder_and_statements_to_their_own(self):
        gw = gateway()
        for subject, f in [("Tax Invoice 778899", "invoice_bidfood.pdf"), ("Credit Note CN-9001", "credit_note.pdf"),
                           ("POD D-77120", "pod.pdf"), ("Statement for September", "statement.pdf")]:
            gw.add(ONE, mail(subject, [f]))
        s = run(gw)
        self.assertEqual(sorted(self.files_under("TestOutput")), ["MOBIL COLYTON/Acme Foods Pty Ltd_CN-9001_CN.pdf",
                         "MOBIL COLYTON/Bidfood Australia Ltd_778899.pdf", "MOBIL COLYTON/Haulage Express Pty Ltd_D-77120_POD.pdf"])
        self.assertEqual(self.files_under("Holding"), ["_Statements/August 2026/ONE PEARL ENERGY/Fresh Fruit Co Pty Ltd_Mobil Colyton_12-08-2026_STM.pdf"])
        self.assertEqual((s.c["Filed"], s.c["Parked"], s.c["Unassigned"]), (3, 1, 0))
        row = connect_db(self.home).execute("SELECT site, site_source FROM docs WHERE number='778899'").fetchone()
        self.assertEqual(row["site"], "Mobil Colyton")

    def test_an_invoice_addressed_to_pearl_energy_and_a_town_finds_the_site(self):
        self.add_rows("Sites", [site_row("Mobil Robina", None, "MOBIL ROBINA", "Radhe Krishna Investment Pty Ltd")])
        gw = gateway()
        gw.add(ONE, mail("FW: Invoice NS-4040", ["invoice_company_named.pdf"]))              # to RADHE KRISHNA INVESTMENT, "PEARL ENERGY ROBINA"
        run(gw)
        self.assertEqual(self.files_under("TestOutput"), ["MOBIL ROBINA/Nova Snacks Pty Ltd_NS-4040.pdf"])

    def test_an_unknown_site_waits_and_the_person_chooses_it_in_to_fix(self):
        gw = gateway()
        gw.add(ONE, mail("FW: Invoice NS-4040", ["invoice_company_named.pdf"]))              # addressed to Pearl Energy Robina: not in the site list
        s = run(gw)
        self.assertEqual((s.c["Unassigned"], s.c["Filed"]), (1, 0))
        self.assertEqual(self.files_under("TestOutput"), [])
        self.assertIn("Nova Snacks Pty Ltd_NS-4040.pdf", self.files_under("Holding", "_Unassigned"))
        reason = connect_db(self.home).execute("SELECT reason FROM docs").fetchone()[0]
        self.assertIn("site could not be determined", reason)

        wb = load_workbook(self.path("Reports", "To_Fix.xlsx"))
        ws = wb["To fix"]
        self.assertEqual(ws.cell(1, 16).value, "Site folder (TYPE HERE)")
        self.assertTrue(any(dv.sqref and "P2" in str(dv.sqref) or "P" in str(dv.sqref) for dv in ws.data_validations.dataValidation))
        self.assertIn("MOBIL DUBBO", [c.value for c in wb["Lists"]["C"]])                   # the drop-down offers the real folder names
        ws.cell(2, 16).value = "MOBIL DUBBO"
        wb.save(self.path("Reports", "To_Fix.xlsx"))
        wb.close()
        s = run(gw)
        self.assertEqual((s.c["fixes"], s.c["released"]), (1, 1))
        self.assertEqual(self.files_under("TestOutput"), ["MOBIL DUBBO/Nova Snacks Pty Ltd_NS-4040.pdf"])
        self.assertEqual(self.files_under("Holding", "_Unassigned"), [])
        conn = connect_db(self.home)
        self.assertEqual(tuple(conn.execute("SELECT site, site_source, status FROM docs").fetchone()), ("Mobil Dubbo", "manual", "Filed"))
        self.assertEqual(conn.execute("SELECT field, new FROM corrections").fetchone()["field"], "site")

    def test_live_mode_waits_for_a_missing_site_folder_then_files_it(self):
        root = self.path("EMAIL INVOICES")
        os.makedirs(root)
        self.set_run(Destination_mode="live", Live_destination_root=root, Live_mode_creates_missing_folders_like_test_mode="No", Create_site_folder_in_live_mode="No")
        gw = gateway()
        gw.add(ONE, mail("Tax Invoice 778899", ["invoice_bidfood.pdf"]))
        s = run(gw)
        self.assertEqual(s.c["Unassigned"], 1)
        self.assertEqual(self.docs_under("EMAIL INVOICES"), [])                               # nothing created in SharePoint
        self.assertIn("site folder does not exist", connect_db(self.home).execute("SELECT reason FROM docs").fetchone()[0])
        os.makedirs(os.path.join(root, "MOBIL COLYTON"))                                      # somebody creates the site folder
        s = run(gw)
        self.assertEqual(s.c["released"], 1)
        self.assertEqual(self.docs_under("EMAIL INVOICES"), ["MOBIL COLYTON/Bidfood Australia Ltd_778899.pdf"])

    def test_the_real_folder_name_is_used_even_with_a_double_space(self):
        """The real folder is 'AMPOL  KELSO' (two spaces). The bot must write into THAT folder, not into 'AMPOL KELSO'."""
        self.add_rows("Sites", [site_row("Ampol Kelso", "5 Kelso Street, Kelso QLD 4815", "AMPOL  KELSO")])
        root = self.path("SP")
        os.makedirs(os.path.join(root, "AMPOL  KELSO"))
        self.set_run(Destination_mode="live", Live_destination_root=root)
        self.assertEqual(sites.folder_for("Ampol Kelso"), "AMPOL  KELSO")
        self.assertEqual(Router().render(D, site=sites.folder_for("Ampol Kelso")), ["AMPOL  KELSO"])
        gw = gateway()
        gw.add(ONE, mail("Tax Invoice INV-1001", ["invoice_acme.pdf"]))                       # delivered to the Colyton address
        self.add_rows("Sites", [COLYTON])
        os.makedirs(os.path.join(root, "MOBIL COLYTON"))
        run(gw)
        self.assertEqual(self.docs_under("SP"), ["MOBIL COLYTON/Acme Foods Pty Ltd_INV-1001.pdf"])
        self.assertEqual(sorted(os.listdir(root)), ["AMPOL  KELSO", "HOLDINGS", "MOBIL COLYTON", "Reports"])   # no second Kelso folder

    def test_company_then_site_layout(self):
        self.set_run(Destination_pattern="{COMPANY}\\{SITE}")
        gw = gateway()
        gw.add(ONE, mail("Tax Invoice 778899", ["invoice_bidfood.pdf"]))
        run(gw)
        self.assertEqual(self.files_under("TestOutput"), ["ONE PEARL ENERGY/MOBIL COLYTON/Bidfood Australia Ltd_778899.pdf"])


class CompanyThenSiteTests(Isolated, unittest.TestCase):
    """Step 1 the company, step 2 the site among that company's sites."""

    def setUp(self):
        super().setUp()
        self.set_run(Destination_pattern="{SITE}")
        # the real situation: two Gretas that belong to different companies, plus enough Mobil/Ampol sites for the brand to be known
        self.add_rows("Sites", [site_row("Pearl Energy Greta", None, "PEARL ENERGY GRETA", "Radhe Krishna Investment Pty Ltd"),
                                site_row("Ampol Greta", None, "AMPOL GRETA", "Shri Hari Investment Pty Ltd"),
                                site_row("Ampol Dubbo", None, "AMPOL DUBBO", "Shri Hari Investment Pty Ltd"),
                                site_row("Ampol Moss Vale", None, "AMPOL MOSS VALE", "Shri Hari Investment Pty Ltd"),
                                site_row("Ampol Springvale", None, "AMPOL SPRINGVALE", "Shri Hari Investment Pty Ltd"),      # two sites in one town,
                                site_row("Mobil Springvale", None, "MOBIL SPRINGVALE", "Radhe Krishna Investment Pty Ltd"),   # two different companies
                                site_row("Mobil Alpha", None, "MOBIL ALPHA", "Radhe Krishna Investment Pty Ltd"),             # (3+ sites make a brand,
                                site_row("Mobil Beta", None, "MOBIL BETA", "Radhe Krishna Investment Pty Ltd"),               #  as in the real list)
                                site_row("Mobil Robina", None, "MOBIL ROBINA", "Pearl Energy No 1 Pty Ltd")])

    def pdf_to(self, customer_company, town, number, name):
        """An invoice PDF addressed to a company and a 'PEARL ENERGY <TOWN>' site."""
        path = os.path.join(self.home, name)
        make_data.invoice(path, "Orbit Wholesale Pty Ltd", number, net=80.0, customer=(customer_company, f"PEARL ENERGY {town}"),
                          address="1 Nowhere Lane Nowhereville NSW 2000")
        return path

    def test_the_same_town_resolves_by_company(self):
        gw = gateway()
        gw.add(ONE, mail("Invoice 1", [self.pdf_to("SHRI HARI INVESTMENT PTY LTD", "GRETA", "G-1", "a.pdf")], hours_ago=2))
        gw.add(ONE, mail("Invoice 2", [self.pdf_to("RADHE KRISHNA INVESTMENT PTY LTD", "GRETA", "G-2", "b.pdf")], hours_ago=3))
        run(gw)
        self.assertEqual(sorted(self.files_under("TestOutput")), ["AMPOL GRETA/Orbit Wholesale Pty Ltd_G-1.pdf",
                                                                   "PEARL ENERGY GRETA/Orbit Wholesale Pty Ltd_G-2.pdf"])
        rows = dict(connect_db(self.home).execute("SELECT number, site_source FROM docs").fetchall())
        self.assertIn("within SHRI HARI INVESTMENT", rows["G-1"])

    def test_the_company_decides_between_two_sites_of_one_town(self):
        gw = gateway()
        gw.add(ONE, mail("Invoice 6", [self.pdf_to("SHRI HARI INVESTMENT PTY LTD", "SPRINGVALE", "S-6", "f.pdf")], hours_ago=2))
        gw.add(ONE, mail("Invoice 7", [self.pdf_to("RADHE KRISHNA INVESTMENT PTY LTD", "SPRINGVALE", "S-7", "g.pdf")], hours_ago=3))
        run(gw)
        self.assertEqual(sorted(self.files_under("TestOutput")), ["AMPOL SPRINGVALE/Orbit Wholesale Pty Ltd_S-6.pdf",
                                                                   "MOBIL SPRINGVALE/Orbit Wholesale Pty Ltd_S-7.pdf"])

    def test_a_town_with_two_companies_and_no_company_clue_waits(self):
        gw = gateway()
        gw.add(ONE, mail("Invoice 3", [self.pdf_to("SOMEONE ELSE PTY LTD", "SPRINGVALE", "S-3", "c.pdf")]))   # no known company: both Springvale sites fit
        s = run(gw)
        self.assertEqual((s.c["Unassigned"], s.c["Filed"]), (1, 0))

    def test_the_company_on_the_document_is_trusted_over_the_site_name(self):
        gw = gateway()
        # Robina belongs to Pearl Energy No 1, but the document is addressed to Shri Hari: do not guess across companies
        gw.add(ONE, mail("Invoice 4", [self.pdf_to("SHRI HARI INVESTMENT PTY LTD", "ROBINA", "R-4", "d.pdf")]))
        s = run(gw)
        self.assertEqual(s.c["Unassigned"], 1)
        reason = connect_db(self.home).execute("SELECT reason FROM docs").fetchone()[0]
        self.assertIn("no site of SHRI HARI INVESTMENT fits", reason)

    def test_a_weaker_company_clue_falls_back_to_all_sites_and_says_so(self):
        # no company on the document; the MAILBOX (One = Pearl Energy No 1) is a weak clue; Greta (Shri Hari) is found in another company
        path = os.path.join(self.home, "e.pdf")
        make_data.invoice(path, "Orbit Wholesale Pty Ltd", "G-5", net=80.0, customer=("PEARL ENERGY MOSS VALE",))
        gw = gateway()
        gw.add(ONE, mail("Invoice 5", [path]))
        s = run(gw)
        self.assertEqual(self.files_under("TestOutput"), ["AMPOL MOSS VALE/Orbit Wholesale Pty Ltd_G-5.pdf"])
        warn = connect_db(self.home).execute("SELECT warnings FROM docs").fetchone()[0]
        self.assertIn("belongs to", warn)

    def test_the_switch_turns_the_two_steps_off(self):
        self.set_run(Choose_the_site_within_the_company_first="No")
        gw = gateway()
        gw.add(ONE, mail("Invoice 8", [self.pdf_to("SHRI HARI INVESTMENT PTY LTD", "SPRINGVALE", "S-8", "h.pdf")]))
        s = run(gw)
        self.assertEqual(s.c["Unassigned"], 1)                                             # the old way: two Springvale sites, so no certain site


MONTH = "OCTOBER 2026 EMAIL INVOICES"


class MonthCompanySiteTests(Isolated, unittest.TestCase):
    """The real SharePoint shape:  <root> \\ OCTOBER 2026 EMAIL INVOICES \\ ONE PEARL ENERGY \\ MOBIL BALLINA \\ file"""

    def setUp(self):
        super().setUp()
        self.set_run(Destination_pattern="{MONTH_FOLDER}\\{COMPANY}\\{SITE}")
        self.add_rows("Sites", [COLYTON, site_row("Mobil Dubbo", None, "MOBIL DUBBO"), site_row("Mobil Moss Vale"), site_row("Mobil Ballina")])

    def test_the_file_goes_month_company_site(self):
        gw = gateway()
        gw.add(ONE, mail("Tax Invoice INV-1001", ["invoice_acme.pdf"]))                           # the Colyton address, arriving in the One mailbox
        s = run(gw)
        self.assertEqual(self.files_under("TestOutput"), [f"{MONTH}/ONE PEARL ENERGY/MOBIL COLYTON/Acme Foods Pty Ltd_INV-1001.pdf"])
        self.assertEqual((s.c["Filed"], s.c["Unassigned"]), (1, 0))

    def test_when_the_site_is_not_found_it_is_filed_in_the_company_folder_and_says_so(self):
        gw = gateway()
        gw.add(ONE, mail("FW: Invoice NS-4040", ["invoice_company_named.pdf"]))                    # RADHE KRISHNA INVESTMENT; Robina is not in the list
        s = run(gw)
        self.assertEqual(self.files_under("TestOutput"), [f"{MONTH}/RADHE KRISHNA INVESTMENT/Nova Snacks Pty Ltd_NS-4040.pdf"])
        row = connect_db(self.home).execute("SELECT status, remarks FROM docs").fetchone()
        self.assertEqual((row["status"], row["remarks"]), ("Filed", "SITE NOT FOUND"))
        wb = load_workbook(self.path("Reports", "Record_2026-10.xlsx"))
        head = [c.value for c in wb["Record"][1]]
        self.assertEqual([r[head.index("Remarks")] for r in wb["Record"].iter_rows(min_row=2, values_only=True)], ["SITE NOT FOUND"])
        wb.close()

    def test_the_fallback_can_be_switched_off(self):
        self.set_run(File_in_the_company_folder_when_the_site_is_not_found="No")
        gw = gateway()
        gw.add(ONE, mail("FW: Invoice NS-4040", ["invoice_company_named.pdf"]))
        s = run(gw)
        self.assertEqual((s.c["Unassigned"], s.c["Filed"]), (1, 0))
        self.assertEqual(self.files_under("TestOutput"), [])

    def test_the_remarks_combine_when_both_the_site_and_the_amounts_are_missing(self):
        from tests import make_data
        path = os.path.join(self.home, "none.pdf")
        make_data.simple(path, "TAX INVOICE", ["Invoice No: 7001", "Invoice Date: 02/10/2026", "SHRI HARI INVESTMENT PTY LTD"], "Orbit Wholesale Pty Ltd")
        gw = gateway()
        gw.add(ONE, mail("Tax Invoice 7001", [path]))
        run(gw)
        row = connect_db(self.home).execute("SELECT remarks FROM docs").fetchone()
        self.assertEqual(row["remarks"], "GST NOT FOUND; SITE NOT FOUND")

    def test_a_live_run_uses_existing_folders_and_creates_nothing(self):
        root = self.path("SP", "MAIL INVOICES")
        site = os.path.join(root, MONTH, "ONE PEARL ENERGY", "MOBIL COLYTON")
        os.makedirs(site)
        self.set_run(Destination_mode="live", Live_destination_root=root)
        gw = gateway()
        gw.add(ONE, mail("Tax Invoice INV-1001", ["invoice_acme.pdf"]))
        s = run(gw)
        self.assertEqual(s.c["Filed"], 1)
        self.assertEqual(sorted(os.listdir(os.path.join(root, MONTH))), ["ONE PEARL ENERGY"])      # no second folder was made
        self.assertEqual(os.listdir(site), ["Acme Foods Pty Ltd_INV-1001.pdf"])
        s = run(gw)                                                                                # and the same again changes nothing
        self.assertEqual(s.c["documents"], 0)

    def test_matching_reads_month_then_company_then_site(self):
        root = self.path("SP", "MAIL INVOICES")
        for company, sites_ in (("ONE PEARL ENERGY", ["MOBIL COLYTON", "MOBIL DUBBO", "SOMETHING ELSE"]), ("TWO PEARL ENERGY", ["MOBIL BALLINA"])):
            for name in sites_:
                os.makedirs(os.path.join(root, MONTH, company, name))
        os.makedirs(os.path.join(root, "SEPTEMBER 2026 EMAIL INVOICES", "ONE PEARL ENERGY", "OLD FOLDER"))
        res = sites.match_folders(root)
        self.assertEqual(res["matched"]["Mobil Colyton"], "MOBIL COLYTON")
        self.assertEqual(res["matched"]["Mobil Dubbo"], "MOBIL DUBBO")
        self.assertIn("Mobil Ballina", res["missing"] + list(res["matched"]) + list(res["fuzzy"]))   # Ballina belongs to company No 1 here: not under TWO
        self.assertIn("ONE PEARL ENERGY\\SOMETHING ELSE", res["extra"])                           # a folder with no site, shown with its company
        self.assertNotIn("ONE PEARL ENERGY\\OLD FOLDER", res["extra"])                            # only the current month is read
        self.assertTrue(res["looked_in"].endswith(MONTH))


class PanelLayoutTests(Isolated, unittest.TestCase):
    def test_layout_choices_and_the_match_button(self):
        c = panel.Controller(opener=lambda p: None)
        run_py = os.path.join(ROOT, "run_bot.py")
        self.assertEqual(c.args_for("match", {}), [run_py, "--match-folders"])
        self.assertFalse(c.set_destination("test", "", "sideways")["ok"])
        self.assertEqual(panel.LAYOUTS["month_company_site"], "{MONTH_FOLDER}\\{COMPANY}\\{SITE}")
        self.assertTrue(c.set_destination("test", "", "month_company_site")["ok"])
        settings.reset()
        self.assertEqual(settings.run_value("Destination pattern"), "{MONTH_FOLDER}\\{COMPANY}\\{SITE}")
        os.makedirs(self.path("SP"))
        r = c.set_destination("live", self.path("SP"), "site")
        self.assertTrue(r["ok"], r)
        settings.reset()
        self.assertEqual(settings.run_value("Destination pattern"), "{SITE}")
        self.assertEqual(c.overview()["layout"], "{SITE}")
        self.assertTrue(c.set_destination("test", "", "company_site")["ok"])
        settings.reset()
        self.assertEqual(settings.run_value("Destination pattern"), "{COMPANY}\\{SITE}")


if __name__ == "__main__":
    unittest.main()
