"""Tests for saved .msg e-mails: the reader, the 'run the bot on a folder of .msg files' mode, scorecard and diagnose."""
import datetime
import os
import shutil
import subprocess
import sys
import unittest

from tests.helpers import DATA, ROOT, Isolated, run, write

from invoicebot import msgfile, msgtest
from tests.msgwriter import write_msg

WHEN = datetime.datetime(2026, 10, 3, 9, 30)


def pdf(name):
    with open(os.path.join(DATA, name), "rb") as f:
        return f.read()


class MsgTests(Isolated, unittest.TestCase):
    def setUp(self):
        super().setUp()
        self.folder = self.path("msgs")
        os.makedirs(self.folder)
        m = lambda name, *a, **k: write_msg(os.path.join(self.folder, name), *a, **k)
        m("invoice.msg", "Tax Invoice 778899", WHEN, "Bidfood", "accounts@bidfood.example", "Pearl Energy Colyton",
          attachments=[dict(filename="invoice_bidfood.pdf", data=pdf("invoice_bidfood.pdf")),
                       dict(filename="logo.png", data=b"p" * 5000, hidden=True)])
        m("overdue.msg", "Overdue Account Reminder", WHEN - datetime.timedelta(hours=1), "Orbit", "a@orbit.example", "Someone Unknown",
          attachments=[dict(filename="invoice_other.pdf", data=pdf("invoice_other.pdf"))])
        m("statement.msg", "Statement for September", WHEN - datetime.timedelta(hours=2), "Fresh", "a@fresh.example", "One Pearl Energy",
          attachments=[dict(filename="statement.pdf", data=pdf("statement.pdf"))])
        write(os.path.join(self.folder, "broken.msg"), b"this is not a msg file")

    def test_reader(self):
        m = msgfile.load(os.path.join(self.folder, "invoice.msg"))
        self.assertEqual((m["subject"], m["sender_name"], m["sender_email"], m["to"], m["received"]),
                         ("Tax Invoice 778899", "Bidfood", "accounts@bidfood.example", "Pearl Energy Colyton", WHEN))
        self.assertEqual([(a["filename"], a["method"], a["hidden"]) for a in m["attachments"]],
                         [("invoice_bidfood.pdf", 1, False), ("logo.png", 1, True)])
        self.assertTrue(m["attachments"][0]["data"].startswith(b"%PDF"))

    def test_bad_files_are_reported_not_fatal(self):
        with self.assertRaises(msgfile.MsgError):
            msgfile.load(os.path.join(self.folder, "broken.msg"))
        cut = os.path.join(self.folder, "cut.msg")
        with open(os.path.join(self.folder, "invoice.msg"), "rb") as f:
            write(cut, f.read()[:3000])
        with self.assertRaises(msgfile.MsgError):
            msgfile.load(cut)

    def test_gateway_maps_recipients_to_company_mailboxes(self):
        gw, first, last, n, problems = msgtest.build(self.folder)
        self.assertEqual((n, first, last), (3, datetime.date(2026, 10, 3), datetime.date(2026, 10, 3)))
        self.assertEqual([p[0] for p in problems], ["broken.msg"])
        subjects = {box: sorted(x.subject for x in ms) for box, ms in gw.boxes.items() if ms}
        self.assertEqual(subjects["pearl.colyton@pearlenergy.com.au"], ["Tax Invoice 778899"])        # recipient "Pearl Energy Colyton"
        self.assertEqual(subjects["One@pearlenergy.com.au"], ["Overdue Account Reminder", "Statement for September"])   # known + fallback

    def test_the_bot_runs_on_msg_files(self):
        gw, first, last, _n, _p = msgtest.build(self.folder)
        s = run(gw, from_date=first, to_date=last)
        self.assertEqual((s.c["blocked"], s.c["Filed"], s.c["Parked"]), (1, 1, 1))
        files = self.files_under("TestOutput")
        self.assertEqual(len(files), 1)
        self.assertTrue(files[0].endswith("COLYTON FUELS/Bidfood Australia Ltd_778899.pdf"))
        self.assertIn("_Statements/August 2026/ONE PEARL ENERGY/Fresh Fruit Co Pty Ltd_NoSite_12-08-2026_STM.pdf", self.files_under("Holding"))
        overdue = next(x for x in gw.boxes["One@pearlenergy.com.au"] if x.subject.startswith("Overdue"))
        self.assertEqual((overdue.flag, overdue.unread), (2, True))

    def test_command_line_msg_folder(self):
        work = os.path.join(ROOT, "msg_test_run")
        self.addCleanup(shutil.rmtree, work, True)
        out = subprocess.run([sys.executable, os.path.join(ROOT, "run_bot.py"), "--msg-folder", self.folder, "--workers", "1"],
                             capture_output=True, text=True, timeout=180, cwd=ROOT)
        self.assertEqual(out.returncode, 0, out.stderr[-600:])
        self.assertIn("3 .msg e-mails loaded", out.stdout)
        self.assertIn("could not read broken.msg", out.stdout)
        self.assertIn("the real e-mails are unchanged", out.stdout)
        self.assertTrue(os.path.isdir(os.path.join(work, "TestOutput")))

    def test_command_line_rejects_a_missing_folder(self):
        out = subprocess.run([sys.executable, os.path.join(ROOT, "run_bot.py"), "--msg-folder", self.path("nope")],
                             capture_output=True, text=True, timeout=60, cwd=ROOT)
        self.assertEqual(out.returncode, 2)
        self.assertIn("does not exist", out.stdout)

    def test_scorecard_reads_msg_files(self):
        out = subprocess.run([sys.executable, os.path.join(ROOT, "scorecard.py"), self.folder, "--out", self.path("sc.xlsx")],
                             capture_output=True, text=True, timeout=120, cwd=ROOT)
        self.assertEqual(out.returncode, 0, out.stderr[-600:])
        self.assertIn("invoice.msg > invoice_bidfood.pdf", out.stdout)
        self.assertIn("BLOCKED ('overdue')", out.stdout)
        self.assertIn("3 documents", out.stdout)

    def test_diagnose_reads_a_msg(self):
        out = subprocess.run([sys.executable, os.path.join(ROOT, "diagnose.py"), os.path.join(self.folder, "overdue.msg")],
                             capture_output=True, text=True, timeout=120, cwd=ROOT)
        self.assertEqual(out.returncode, 0, out.stderr[-600:])
        self.assertIn("BLOCKED by 'overdue'", out.stdout)
        self.assertIn("Orbit Wholesale Pty Ltd", out.stdout)


if __name__ == "__main__":
    unittest.main()
