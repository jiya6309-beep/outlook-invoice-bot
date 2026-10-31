"""Setup step 2: every mailbox is a company, a special mailbox (Fuel invoices / Fleet Card statements / Security reports), TEST, or
Never read - saved to the right sheet, one place per mailbox, one mailbox per special role."""
import os
import unittest

from tests.helpers import Isolated

from invoicebot import companies, mailboxes, settings


def special():
    return {sp["role"]: sp["names"] for sp in companies.special_rows()}


class MailboxRoleTests(Isolated, unittest.TestCase):
    def test_suggestions_show_the_roles(self):
        out = {x["outlook"]: x for x in mailboxes.suggest(["P I Fuel Invoice", "Security", "mstatement@pearlenergy.com.au",
                                                           "office@pearlenergy.com.au"])}
        self.assertEqual(out["P I Fuel Invoice"]["special"], "fuel invoices")
        self.assertEqual(out["Security"]["special"], "security")
        self.assertEqual(out["mstatement@pearlenergy.com.au"]["special"], "fleet card")
        self.assertEqual(out["office@pearlenergy.com.au"]["special"], "never")

    def test_choosing_a_role_for_another_mailbox_moves_it(self):
        mailboxes.apply([{"outlook": "Fuel Box 2", "action": "special", "role": "fuel invoices"}])
        settings.reset()
        self.assertEqual(special()["fuel invoices"], ["Fuel Box 2"])                 # the old P I Fuel Invoice row is replaced
        self.assertIsNone(companies.special_for("pifuelinvoice@pearlenergy.com.au"))

    def test_never_read_and_back_to_a_company(self):
        mailboxes.apply([{"outlook": "accounts@anantaa.com.au", "action": "special", "role": "never"}])
        settings.reset()
        self.assertEqual(companies.special_for("accounts@anantaa.com.au")["role"], "never")
        mailboxes.apply([{"outlook": "accounts@anantaa.com.au", "action": "company", "company": "COLYTON FUELS PTY LTD"}])
        settings.reset()
        self.assertIsNone(companies.special_for("accounts@anantaa.com.au"))
        self.assertIn("accounts@anantaa.com.au", companies.mailbox_names(companies.by_name("COLYTON FUELS PTY LTD")))

    def test_a_fuel_mailbox_set_to_never_is_not_read_at_all(self):
        mailboxes.apply([{"outlook": "P I Fuel Invoice", "action": "special", "role": "never"}])
        settings.reset()
        self.assertNotIn("fuel invoices", special())
        self.assertIsNone(companies.special_for("pifuelinvoice@pearlenergy.com.au"))   # its address alias went too
        self.assertEqual(companies.special_for("P I Fuel Invoice")["role"], "never")

    def test_a_company_mailbox_made_special_leaves_the_companies_sheet(self):
        name = companies.mailbox_names(companies.by_name("COLYTON FUELS PTY LTD"))[0]
        mailboxes.apply([{"outlook": name, "action": "special", "role": "security"}])
        settings.reset()
        self.assertNotIn(name, companies.mailbox_names(companies.by_name("COLYTON FUELS PTY LTD")))
        self.assertEqual(special()["security"], [name])

    def test_only_one_mailbox_per_special_role(self):
        with self.assertRaises(settings.SettingsError):
            mailboxes.apply([{"outlook": "A", "action": "special", "role": "fleet card"},
                             {"outlook": "B", "action": "special", "role": "fleet card"}])


if __name__ == "__main__":
    unittest.main()


class HoldingsAndFuelFolderTests(Isolated, unittest.TestCase):
    def test_holding_root_moves_into_holdings_once(self):
        from openpyxl import load_workbook
        path = settings.settings_file()
        wb = load_workbook(path)
        ws = wb["Run"]
        for r in range(ws.max_row, 1, -1):
            k = str(ws.cell(r, 1).value or "")
            if k == "Settings version":
                ws.cell(r, 2).value = "5.12.0"
            if k == "Holding root":
                ws.cell(r, 2).value = settings.OLD_MAILS
        wb.save(path)
        settings.reset()
        added = settings.upgrade_workbook()
        settings.reset()
        self.assertTrue(any("HOLDINGS" in a for a in added))
        # 5.13.0 moved it to MAILS\HOLDINGS; 5.15.0 then empties it: LIVE uses <EMAIL INVOICES>\HOLDINGS
        self.assertEqual(settings.run_value("Holding root"), "")
        self.set_run(Destination_mode="live")
        self.assertEqual(settings.holding_root(), os.path.join(settings.LIVE_ROOT, "HOLDINGS"))
        self.assertEqual(settings.upgrade_workbook(), [])

    def test_a_holding_root_someone_chose_is_kept(self):
        self.set_run(Holding_root="D:\\Somewhere\\Else", Settings_version="5.12.0")
        settings.upgrade_workbook()
        settings.reset()
        self.assertEqual(settings.run_value("Holding root"), "D:\\Somewhere\\Else")

    def test_a_fuel_customer_that_is_not_a_company_gets_its_own_folder(self):
        import os
        from unittest import mock
        from tests.helpers import gateway, mail, run
        from invoicebot import fuel
        real = fuel.parse

        def skm(path):
            got = real(path)
            if got:
                got.update(customer="SKM DIESEL PTY LTD - SKM DIESEL DEPOT", company_text="SKM DIESEL PTY LTD", site_text="SKM DIESEL DEPOT")
            return got
        g = gateway()
        g.add_mailbox("P I Fuel Invoice")
        g.add("P I Fuel Invoice", mail("Fuel", ["fuel_INV-115316.pdf"]))
        with mock.patch.object(fuel, "parse", skm):
            s = run(g)
        self.assertEqual(s.c["Filed"], 1)
        self.assertEqual(self.files_under("TestOutput", "FUEL INVOICES"), ["SKM DIESEL PTY LTD/INV-115316_SKM DIESEL DEPOT.pdf"])
        os.makedirs(self.path("TestOutput", "FUEL INVOICES", "SKM Diesel Pty. Ltd."))   # an existing folder with another spelling
        g2 = gateway()
        g2.add_mailbox("P I Fuel Invoice")
        g2.add("P I Fuel Invoice", mail("Fuel", ["fuel_INV-114870.pdf"], hours_ago=0.5))
        with mock.patch.object(fuel, "parse", skm):
            run(g2)
        self.assertEqual(len(os.listdir(self.path("TestOutput", "FUEL INVOICES"))), 2)    # the 2nd went into one of the two, no 3rd folder
