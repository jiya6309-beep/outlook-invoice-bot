"""5.8.0: statements / order confirmations / _Unrecognised go into month folders; statements and order confirmations are named
supplier_site_date_TAG (NoSite + a To_Fix flag when no site is found); the one-time tidy of old files."""
import os
import shutil
import unittest

from openpyxl import load_workbook

from tests import make_data
from tests.helpers import ONE, Isolated, connect_db, gateway, mail, run
from invoicebot import db, exports, settings, tidy


class HoldingTests(Isolated, unittest.TestCase):
    def setUp(self):
        super().setUp()
        self.add_rows("Sites", [["Pearl Energy Colyton", "Pearl Energy No 1 Pty Ltd", "", "", "", "", "Yes", "", "", ""]])

    def pdf(self, name, title, lines, supplier="Zed Supplies Pty Ltd"):
        p = os.path.join(self.home, name)
        make_data.simple(p, title, lines, supplier=supplier)
        return p

    def one(self, subject, title, lines, **kw):
        gw = gateway()
        gw.add(ONE, mail(subject, [self.pdf("d.pdf", title, lines)], **kw))
        s = run(gw)
        rows = [dict(r) for r in connect_db(self.home).execute("SELECT * FROM docs ORDER BY id")]
        return s, rows[-1], gw

    def rel(self, r):
        return os.path.relpath(r["current_path"], self.home).replace(os.sep, "/")

    def test_statement_with_a_site_is_named_supplier_site_date_and_sits_in_its_month(self):
        s, r, _ = self.one("Statement for October", "WEEKLY STATEMENT", ["Statement Date: 02/10/2026", "PEARL ENERGY COLYTON", "Total 500.00"])
        self.assertEqual((r["status"], r["doc_type"]), ("Parked", "Statement"))
        self.assertIn("Holding/_Statements/October 2026/", self.rel(r))
        self.assertTrue(os.path.basename(r["current_path"]).endswith("_02-10-2026_STM.pdf"), r["current_path"])
        self.assertEqual(os.path.basename(r["current_path"]), "Zed Supplies Pty Ltd_Pearl Energy Colyton_02-10-2026_STM.pdf")
        self.assertNotIn("NoSite", r["current_path"])
        self.assertNotIn(settings.MISSING_SITE, r["remarks"] or "")

    def test_statement_without_a_site_says_nosite_and_is_flagged_in_to_fix(self):
        s, r, _ = self.one("Statement for October", "WEEKLY STATEMENT", ["Statement Date: 02/10/2026", "Total 500.00"])
        self.assertEqual(os.path.basename(r["current_path"]), "Zed Supplies Pty Ltd_NoSite_02-10-2026_STM.pdf")
        self.assertIn(settings.MISSING_SITE, r["remarks"])
        conn = connect_db(self.home)
        exports.write_to_fix(conn)
        ws = load_workbook(os.path.join(exports.reports_dir(), "To_Fix.xlsx"))["To fix"]
        self.assertEqual(ws.cell(1, 18).value, "Statement missing data")
        self.assertEqual(ws.cell(2, 18).value, "Site not found in this document")
        self.assertEqual(ws.cell(2, 2).value, "Parked")

    def test_a_statement_with_a_site_is_not_listed_in_to_fix(self):
        self.one("Statement for October", "WEEKLY STATEMENT", ["Statement Date: 02/10/2026", "PEARL ENERGY COLYTON", "Total 500.00"])
        exports.write_to_fix(connect_db(self.home))
        ws = load_workbook(os.path.join(exports.reports_dir(), "To_Fix.xlsx"))["To fix"]
        self.assertIsNone(ws.cell(2, 1).value)

    def test_order_confirmation_goes_to_its_own_holding_folder(self):
        s, r, _ = self.one("Order Confirmation 4411", "ORDER CONFIRMATION", ["Order No: 4411", "Date: 02/10/2026", "Total 500.00"])
        self.assertEqual((r["status"], r["doc_type"]), ("Parked", "Order Confirmation"))
        self.assertIn("Holding/_Order Confirmations/October 2026/", self.rel(r))
        self.assertTrue(os.path.basename(r["current_path"]).endswith("_NoSite_02-10-2026_OC.pdf") or "_02-10-2026_OC.pdf" in r["current_path"])

    def test_a_blocked_word_still_blocks_an_order_confirmation(self):
        gw = gateway()
        em = gw.add(ONE, mail("URGENT Order Confirmation 4412", [self.pdf("e.pdf", "ORDER CONFIRMATION", ["Order No: 4412"])]))
        s = run(gw)
        self.assertEqual(s.c["blocked"], 1)
        self.assertEqual(em.flag, 2)

    def test_statement_dated_last_month_goes_to_that_month_folder(self):
        s, r, _ = self.one("Statement Sept", "WEEKLY STATEMENT", ["Statement Date: 25/09/2026", "Total 500.00"])
        self.assertIn("Holding/_Statements/September 2026/", self.rel(r))

    def test_a_statement_dated_far_away_uses_the_arrival_month(self):
        s, r, _ = self.one("Statement old", "WEEKLY STATEMENT", ["Statement Date: 25/01/2026", "Total 500.00"])
        self.assertIn("Holding/_Statements/October 2026/", self.rel(r))

    def test_unrecognised_documents_get_a_month_folder_too(self):
        gw = gateway()
        gw.add(ONE, mail("Tax Invoice 1", ["corrupt.pdf"]))
        run(gw)
        r = [dict(x) for x in connect_db(self.home).execute("SELECT * FROM docs")][-1]
        self.assertEqual(r["status"], "Unrecognised")
        self.assertIn("Holding/_Unrecognised/October 2026/", self.rel(r))

    def test_parked_on_purpose_types_sit_in_month_then_type(self):
        s, r, _ = self.one("Account Summary", "ACCOUNT SUMMARY", ["Date: 02/10/2026", "Total 500.00"])
        self.assertEqual(r["status"], "Parked")
        self.assertIn("Holding/_Unrecognised/October 2026/Account Summary/", self.rel(r))

    def test_tidy_lists_first_and_then_moves_and_renames_old_files(self):
        s, r, _ = self.one("Statement for October", "WEEKLY STATEMENT", ["Statement Date: 02/10/2026", "PEARL ENERGY COLYTON", "Total 500.00"])
        flat = os.path.join(settings.holding_root(), "_Statements")
        old = os.path.join(flat, "Zed Supplies Pty Ltd_02-10-2026_STM.pdf")           # the way 5.7 named and placed it
        shutil.move(r["current_path"], old)
        conn = connect_db(self.home)
        db.update_doc(conn, r["ref"], current_path=old, dest_dir=flat, dest_name=os.path.basename(old))
        conn.commit()
        conn.close()
        moved, failed = tidy.run(apply=False, log=lambda *a: None)
        self.assertEqual((moved, failed), (1, 0))
        self.assertTrue(os.path.exists(old))                                              # list only: nothing changed
        self.assertTrue(os.path.exists(os.path.join(exports.reports_dir(), "Tidy_Holding_Preview.xlsx")))
        moved, failed = tidy.run(apply=True, log=lambda *a: None)
        self.assertEqual((moved, failed), (1, 0))
        self.assertFalse(os.path.exists(old))
        row = db.get_doc(connect_db(self.home), r["ref"])
        self.assertIn("_Statements/October 2026/", row["current_path"].replace(os.sep, "/"))
        self.assertTrue(os.path.isfile(row["current_path"]))
        self.assertTrue(row["current_path"].endswith("_02-10-2026_STM.pdf") and row["current_path"].count("_") >= 3)
        self.assertEqual(tidy.run(apply=True, log=lambda *a: None), (0, 0))               # nothing left to do: safe to repeat

    def test_old_settings_files_still_get_the_new_statement_name_and_the_order_confirmation_type(self):
        from invoicebot import doctypes
        self.assertEqual(doctypes.rule("Order Confirmation")["subfolder"], "_Order Confirmations")
        self.assertIn("{site}", doctypes.rule("Statement")["pattern"])


if __name__ == "__main__":
    unittest.main()
