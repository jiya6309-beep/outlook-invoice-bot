"""The security bot: the alarm reports (both layouts) -> the fortnightly STORE OPEN & CLOSE Excel, with the agreed rules.
Uses the company's real reports (tests/data/security)."""
import datetime
import glob
import os
import unittest

from openpyxl import load_workbook

from tests.helpers import DATA, Isolated, connect_db, gateway, mail

from invoicebot import db, paths, security, security_reader as reader, security_settings as ss

SEC = os.path.join(DATA, "security")
REPORTS = sorted(glob.glob(os.path.join(SEC, "*.pdf")))
SITES = ["Tomingley", "Deeragun", "Wangaratta", "Nanango", "Tomago"]
D = datetime.date


def site_rows():
    def row(company, no, name, client, addr, wk, we):
        out = [company, no, name, client, addr]
        for i in range(7):
            o, c = we if i >= 5 else wk
            out += [o, c]
        return out + ["Yes", "No"]
    return [row("RADHE KRISHNA INVESTMENT", 115, "BP - Tomingley NSW", "4894-25", "36 Myall Street Tomingley 2869",
                ("4:00 AM", "12:00 AM"), ("4:30 AM", "11:00 PM")),
            row("PEARL ENERGY NO.2", 62, "Pearl Energy - Deeragun QLD", "0873-25", "Deeragun", ("5:00 AM", "9:00 PM"), ("6:00 AM", "8:00 PM")),
            row("GOPINATH INVESTMENT", 99, "BP - Wangaratta VIC", "7444-25", "Wangaratta", ("5:00 AM", "10:00 PM"), ("6:00 AM", "8:00 PM")),
            row("PEARL ENERGY NO.3", 80, "Pearl Energy - Nanango QLD", "0920-25", "Nanango", ("5:30 AM", "9:30 PM"), ("6:00 AM", "9:00 PM")),
            row("PEARL ENERGY NO.2", 64, "Pearl Energy - Tomago NSW", "0989-25", "Tomago", ("4:00 AM", "10:00 PM"), ("6:00 AM", "9:00 PM")),
            row("PEARL ENERGY NO.1", 1, "Pearl Energy - Never Reports NSW", "9999-25", "Nowhere", ("5:00 AM", "9:00 PM"), ("6:00 AM", "8:00 PM"))]


class ReaderTests(unittest.TestCase):
    def test_both_layouts_read_every_arm_and_disarm(self):
        got = {}
        for f in REPORTS:
            for s in reader.parse(f):
                got[s["client"]] = s
        self.assertEqual(sorted(got), ["0873-25", "0920-25", "0989-25", "4894-25", "7444-25"])
        for s in got.values():
            self.assertEqual(len(s["events"]), 14)                         # one disarm + one arm a day, nothing else
        self.assertEqual(got["4894-25"]["start"], D(2026, 9, 28))
        self.assertEqual(got["4894-25"]["name"], "BP - Tomingley NSW")
        self.assertEqual(got["0989-25"]["start"], D(2026, 10, 12))

    def test_keyswitch_counts_and_noise_is_ignored(self):
        s = [x for f in REPORTS for x in reader.parse(f) if x["client"] == "0873-25"][0]
        kinds = [(e[0].strftime("%d %H:%M"), e[1]) for e in s["events"]]
        self.assertIn(("17 00:01", "disarm"), kinds)                        # Disarmed, Keyswitch
        for line in ("Activation 27057261: Battery Missing Or Dead, created", "Test Message: Periodic Test 901",
                     "Late to Close", ": Activation Completed:", "fmartinah: Picked Up Activation 27005701",
                     "65052 2 Early to Close User 2", "458 0 User on Premises User 0", "65051 0 Late to Open User 0"):
            self.assertIsNone(reader.classify(line), line)
        self.assertEqual(reader.classify("1401 2 Armed User 2"), "arm")
        self.assertEqual(reader.classify("Armed: Armed User 0 0"), "arm")
        self.assertEqual(reader.classify("401 2 Disarmed User 2"), "disarm")

    def test_seconds_are_dropped_and_12_hour_times_read(self):
        s = [x for f in REPORTS for x in reader.parse(f) if x["client"] == "4894-25"][0]
        self.assertEqual(s["events"][1][0], datetime.datetime(2026, 9, 28, 22, 3))

    def test_two_sites_in_one_file_are_two_sections(self):
        a = reader.read_text([f for f in REPORTS if "0873" in f][0])
        b = reader.read_text([f for f in REPORTS if "7444" in f][0])
        secs = reader.parse_text(a + "\n" + b)
        self.assertEqual([s["client"] for s in secs], ["0873-25", "7444-25"])


class TimingTests(unittest.TestCase):
    def test_timing_texts(self):
        for text, want in (("5:00 AM - 10:00 PM", (300, 1320)), ("05:00AM - 10:00PM", (300, 1320)), ("5 AM - 10 PM", (300, 1320)),
                           ("06:00:AM - 09:00:PM", (360, 1260)), ("04:00 AM - 12:00 AM", (240, 1440))):
            self.assertEqual(ss.split_timing(text), want, text)
        self.assertEqual(ss.timing_text(240, 1440), "04:00 AM - 12:00 AM")

    def test_fortnights_and_names(self):
        a = D(2026, 10, 5)
        self.assertEqual(security.fortnight(D(2026, 10, 14), a)[:2], (D(2026, 10, 5), D(2026, 10, 18)))
        s, e, lab = security.fortnight(D(2026, 10, 28), a)
        self.assertEqual((s, e), (D(2026, 10, 19), D(2026, 11, 1)))
        self.assertEqual(security.file_name(s, e, lab), "SITE_DATA_OCT_2026_19-10_TO_01-11.xlsx")
        s, e, lab = security.fortnight(D(2026, 11, 30), a)                  # a fifth week starts the next month's file
        self.assertEqual(security.file_name(s, e, lab), "SITE_DATA_DEC_2026_30-11_TO_13-12.xlsx")
        self.assertEqual(security.month_folder(lab), "DECEMBER 2026")


class RunTests(Isolated, unittest.TestCase):
    def setUp(self):
        super().setUp()
        ss.create_default_file(paths.p(ss.FILE_NAME), site_rows())
        self.g = gateway()
        self.g.add_mailbox("Security")
        self.mails = [self.g.add("Security", mail("Weekly report", [f], hours_ago=24)) for f in REPORTS]

    def go(self, **kw):
        kw.setdefault("from_date", D(2026, 9, 28))
        kw.setdefault("to_date", D(2026, 12, 31))
        return security.SecurityRun(self.g, log=lambda *a: None, **kw).run()

    def results(self, client, start, end):
        conf = ss.load()
        site = [s for s in conf["sites"] if s["client"] == client][0]
        conn = connect_db(self.home)
        cover = security.coverage(conn, conf["sites"], start, end)
        return security.day_results(conn, site, start, end, (10, 4, 3.5), {w: v.get(client, False) for w, v in cover.items()})

    def test_the_agreed_results_on_the_real_reports(self):
        r = self.go()
        self.assertFalse(r.warnings, r.warnings)
        tom = self.results("4894-25", D(2026, 9, 28), D(2026, 10, 4))
        self.assertTrue(all(x["close_flag"] and x["close_remark"] == "EARLY" for x in tom.values()))   # closed early every day
        self.assertFalse(any(x["open_flag"] for x in tom.values()))
        thu = tom[D(2026, 10, 1)]
        self.assertEqual((thu["in"].strftime("%H:%M"), thu["open_remark"]), ("03:58", "EARLY"))      # the 3.5 h rule
        dee = self.results("0873-25", D(2026, 10, 12), D(2026, 10, 18))
        self.assertEqual(dee[D(2026, 10, 17)]["open_remark"], "MISSING")                             # keyswitch 00:01 = Friday
        self.assertTrue(dee[D(2026, 10, 17)]["open_flag"])
        self.assertEqual((dee[D(2026, 10, 18)]["close_remark"], dee[D(2026, 10, 18)]["close_diff"]), ("EARLY", -2))
        self.assertFalse(dee[D(2026, 10, 16)]["open_flag"] or dee[D(2026, 10, 16)]["close_flag"])
        wan = self.results("7444-25", D(2026, 10, 12), D(2026, 10, 18))
        self.assertEqual([d.day for d, x in wan.items() if x["open_flag"]], [17, 18])                 # 6:12 and 7:00
        tomago = self.results("0989-25", D(2026, 10, 12), D(2026, 10, 18))
        self.assertFalse(tomago[D(2026, 10, 16)]["open_flag"])                                       # 4:09 = 9 min, allowed
        never = self.results("9999-25", D(2026, 10, 12), D(2026, 10, 18))
        self.assertTrue(all(x["state"] == "no report" for x in never.values()))

    def test_files_folders_and_excel_layout(self):
        self.go()
        root = self.path("TestOutput", "SECURITY REPORTS")
        self.assertTrue(os.path.exists(os.path.join(root, "SEPTEMBER 2026", "SITE_DATA_SEP_2026_21-09_TO_04-10.xlsx")))
        oct_file = os.path.join(root, "OCTOBER 2026", "SITE_DATA_OCT_2026_05-10_TO_18-10.xlsx")
        self.assertTrue(os.path.exists(oct_file))
        self.assertEqual(len(os.listdir(os.path.join(root, "OCTOBER 2026", "Downloaded Reports"))), 4)
        ws = load_workbook(oct_file).active
        rows = {ws.cell(r, 4).value: r for r in range(6, ws.max_row + 1) if ws.cell(r, 4).value}
        r = rows["0873-25"]
        labels = [ws.cell(r + k, 6).value for k in range(6)]
        self.assertEqual(labels, ["Timing", None, "EXPECTED", "ACTUAL", "HOURS", "Remark"])
        col = {ws.cell(2, c).value: c for c in range(7, ws.max_column + 1) if ws.cell(2, c).value}
        sat, sun = col[17], col[18]
        self.assertEqual(ws.cell(r + 1, sat).fill.fgColor.rgb, "00FFFF00")                       # OPEN label yellow (missing)
        self.assertNotEqual(ws.cell(r + 1, sat + 1).fill.fgColor.rgb, "00FFFF00")                # CLOSE label not
        self.assertEqual(ws.cell(r + 1, sun + 1).fill.fgColor.rgb, "00FFFF00")                   # CLOSE label yellow (7:58 PM)
        self.assertNotEqual(ws.cell(r, sun).fill.fgColor.rgb, "00FFFF00")                        # Timing never highlighted
        self.assertEqual(ws.cell(r + 3, sun + 1).value, "7:58 PM")
        self.assertEqual(ws.cell(r + 5, sat).value, "MISSING")
        self.assertIn("Disarmed", ws.cell(r + 3, sun).comment.text)                              # the original line

    def test_emails_ticked_and_a_second_run_does_nothing(self):
        self.go()
        self.assertTrue(all(not m.unread for m in self.mails))
        r2 = self.go()
        self.assertEqual(r2.c["already done"], 5)
        self.assertEqual(r2.c["sites read"], 0)

    def test_unknown_or_periodic_only_reports_leave_the_email_untouched(self):
        conf_rows = [r for r in site_rows() if r[3] != "0920-25"]
        ss.create_default_file(paths.p(ss.FILE_NAME), conf_rows)
        r = self.go()
        nanango = [m for m in self.mails if "0920" in m.attachments[0]["name"]][0]
        self.assertTrue(nanango.unread)
        self.assertIn("0920-25", " ".join(r.unknown))

    def test_rehearsal_changes_nothing(self):
        r = self.go(shadow=True)
        self.assertEqual(r.c["sites read"], 5)
        self.assertTrue(all(m.unread for m in self.mails))
        self.assertFalse(os.path.exists(self.path("TestOutput", "SECURITY REPORTS", "OCTOBER 2026")))

    def test_invoice_runs_never_read_the_security_mailbox(self):
        from tests.helpers import run
        s = run(self.g)
        self.assertEqual(s.c["documents"], 0)
        self.assertTrue(all(m.unread for m in self.mails))


class ImportTests(Isolated, unittest.TestCase):
    def test_settings_built_from_the_october_workbook(self):
        target, n = ss.import_from_workbook(os.path.join(SEC, "SITE_OPEN_CLOSE_DATA_OCT_2026.xlsx"), paths.p("Imported.xlsx"))
        self.assertEqual(n, 119)
        sites = {s["client"]: s for s in ss.load(target)["sites"]}
        tom = sites["4894-25"]
        self.assertEqual((tom["company"], tom["name"]), ("RADHE KRISHNA INVESTMENT", "BP - Tomingley NSW"))
        self.assertEqual(tom["hours"][0], (240, 1440))
        self.assertEqual(tom["hours"][5], (270, 1380))


if __name__ == "__main__":
    unittest.main()
