"""Control panel tests: the controller (buttons' logic) and the real local web server, including its security checks."""
import json
import os
import shutil
import sys
import threading
import time
import unittest
from unittest import mock
import urllib.error
import urllib.request

from tests.helpers import ONE, ROOT, Isolated, gateway, mail, run, write

from invoicebot import mailboxes, panel, settings


def wait_until_finished(controller, seconds=120):
    end = time.time() + seconds
    while time.time() < end:
        s = controller.status(0)
        if not s["running"] and s["job"] and s["job"]["finished"]:
            return s
        time.sleep(0.2)
    raise AssertionError("the run did not finish in time")


class SummaryParsingTests(unittest.TestCase):
    def test_numbers_and_headline(self):
        lines = ["RUN SUMMARY", "E-mails scanned          : 72", "  blocked (red flag)     : 6", "  filed to company folder: 58",
                 "  in _Unrecognised       : 9", "Time: 5 s   Peak memory: main process 77 MB", "STOPPED EARLY: time limit reached"]
        nums, headline = panel.parse_summary(lines)
        self.assertEqual((nums["scanned"], nums["blocked"], nums["filed"], nums["unrecognised"]), (72, 6, 58, 9))
        self.assertEqual(headline, "STOPPED EARLY: time limit reached")
        _n, h = panel.parse_summary(["CANNOT START: Live mode needs the real SharePoint path."])
        self.assertTrue(h.startswith("CANNOT START"))
        _n, h = panel.parse_summary(["Not started: another run is in progress"])
        self.assertTrue(h.startswith("Not started"))


class OutlookTestButtonTests(Isolated, unittest.TestCase):
    def test_the_marking_test_needs_both_boxes(self):
        c = panel.Controller(opener=lambda p: None)
        run_py = os.path.join(ROOT, "run_bot.py")
        self.assertEqual(c.args_for("outlooktest", {"mailbox": "jiya", "subject": "BOTTEST"}),
                         [run_py, "--test-outlook-marking", "--mailbox", "jiya", "--subject", "BOTTEST"])
        for opts in ({}, {"mailbox": "jiya"}, {"subject": "BOTTEST"}):
            with self.assertRaises(ValueError) as cm:
                c.args_for("outlooktest", opts)
            self.assertIn("BOTH boxes", str(cm.exception))


class ControllerTests(Isolated, unittest.TestCase):
    def setUp(self):
        super().setUp()
        self.c = panel.Controller(opener=lambda p: self.opened.append(p))
        self.opened = []
        self.addCleanup(shutil.rmtree, os.path.join(ROOT, "demo_run"), True)
        self.addCleanup(shutil.rmtree, os.path.join(ROOT, "msg_test_run"), True)

    # ---- running
    def test_arguments_are_validated_with_friendly_messages(self):
        for kind, opts, text in [("dates", {}, "both dates"), ("dates", {"from": "2026-10-05", "to": "2026-10-01"}, "later than"),
                                 ("msg", {"folder": self.path("nope")}, "does not exist"), ("nonsense", {}, "unknown")]:
            r = self.c.start(kind, opts)
            self.assertFalse(r["ok"], kind)
            self.assertIn(text, r["message"])
        os.makedirs(self.path("empty"))
        self.assertIn("no .msg files", self.c.start("msg", {"folder": self.path("empty")})["message"])

    def test_arguments(self):
        run_py = os.path.join(ROOT, "run_bot.py")
        self.assertEqual(self.c.args_for("live", {}), [run_py])
        self.assertEqual(self.c.args_for("shadow", {"mailbox": "jiya"}), [run_py, "--shadow", "--mailbox", "jiya"])
        self.assertEqual(self.c.args_for("dates", {"from": "2026-10-01", "to": "2026-10-05", "shadow": True}),
                         [run_py, "--shadow", "--from", "2026-10-01", "--to", "2026-10-05"])
        self.assertEqual(self.c.args_for("demo", {}), [run_py, "--demo"])
        self.assertEqual(self.c.args_for("live", {"mailbox": "jiya", "subject": " BOTTEST "}), [run_py, "--mailbox", "jiya", "--subject", "BOTTEST"])
        self.assertEqual(self.c.args_for("shadow", {"subject": "BOTTEST"}), [run_py, "--shadow", "--subject", "BOTTEST"])

    def test_the_demo_runs_to_the_end(self):
        self.assertTrue(self.c.start("demo")["ok"])
        self.assertFalse(self.c.start("demo")["ok"])                         # only one run at a time
        s = wait_until_finished(self.c)
        self.assertEqual(s["job"]["exit_code"], 0)
        self.assertGreaterEqual(s["summary"]["filed"], 5)
        self.assertEqual(s["summary"]["blocked"], 2)
        self.assertTrue(any("RUN SUMMARY" in line for line in s["lines"]))
        tail = self.c.status(len(s["lines"]) - 2)                            # incremental reading
        self.assertEqual(len(tail["lines"]), 2)

    def test_stop_ends_the_run_and_removes_the_lock(self):
        write(self.path("data", "run.lock"), "pid=1")
        self.c.args_for = lambda kind, opts: ["-c", "import time; time.sleep(60)"]
        self.assertTrue(self.c.start("live")["ok"])
        self.assertTrue(self.c.running())
        self.assertTrue(self.c.stop()["ok"])
        s = wait_until_finished(self.c, 20)
        self.assertTrue(s["job"]["stopped"])
        self.assertFalse(os.path.exists(self.path("data", "run.lock")))
        self.assertFalse(self.c.stop()["ok"])                                # nothing left to stop

    def test_report_for_dates(self):
        gw = gateway()
        gw.add(ONE, mail("Tax Invoice 778899", ["invoice_bidfood.pdf"]))
        gw.add(ONE, mail("Credit Note CN-9001", ["credit_note.pdf"]))
        run(gw)
        self.assertIn("both dates", self.c.start("report", {})["message"])
        self.assertTrue(self.c.start("report", {"from": "2026-10-01", "to": "2026-10-31"})["ok"])
        s = wait_until_finished(self.c)
        self.assertEqual(s["job"]["exit_code"], 0, s["lines"])
        self.assertTrue(any("2 documents" in line for line in s["lines"]))
        self.assertEqual(self.files_under("Reports").count("Report_2026-10-01_2026-10-31.xlsx"), 1)
        from openpyxl import load_workbook
        wb = load_workbook(self.path("Reports", "Report_2026-10-01_2026-10-31.xlsx"))
        self.assertEqual(wb.sheetnames, ["Record", "Needs attention"])
        wb.close()

    def test_check_my_pc(self):
        self.assertTrue(self.c.start("check")["ok"])
        s = wait_until_finished(self.c)
        self.assertTrue(any("setup check" in line for line in s["lines"]))

    # ---- overview
    def test_overview_without_any_run(self):
        o = self.c.overview()
        self.assertEqual((o["mode"], o["counts"], o["last_run"]), ("test", {}, None))
        self.assertFalse(o["reports"]["to_fix"])
        self.assertEqual(len(o["companies"]), 13)

    def test_overview_after_a_run(self):
        gw = gateway()
        gw.add(ONE, mail("Tax Invoice 778899", ["invoice_bidfood.pdf"]))
        gw.add(ONE, mail("inv", ["corrupt.pdf"]))
        run(gw)
        o = self.c.overview()
        self.assertEqual(o["counts"], {"Filed": 1, "Unrecognised": 1})
        self.assertIsNotNone(o["last_run"])
        self.assertTrue(o["reports"]["to_fix"])

    def test_overview_reports_a_broken_settings_file(self):
        write(settings.settings_file(), "not excel")
        settings.reset()
        o = self.c.overview()
        self.assertIn("cannot be read", o["settings_error"])
        self.assertIsNone(o["mode"])

    # ---- opening
    def test_open_shortcuts(self):
        self.assertFalse(self.c.open_path("nonsense")["ok"])
        self.assertFalse(self.c.open_path("to_fix")["ok"])                   # not created yet
        self.assertTrue(self.c.open_path("reports")["ok"])                   # folders are created when missing
        self.assertTrue(self.c.open_path("settings")["ok"])
        self.assertEqual(self.opened, [self.path("Reports"), settings.settings_file()])

    # ---- destination
    def test_switching_destination(self):
        self.assertIn("not found automatically", self.c.set_destination("live", "")["message"])
        self.assertIn("not found", self.c.set_destination("live", self.path("not_synced"))["message"])
        self.assertFalse(self.c.set_destination("sideways")["ok"])
        os.makedirs(self.path("SharePoint"))
        r = self.c.set_destination("live", '"' + self.path("SharePoint") + '"')
        self.assertTrue(r["ok"], r)
        settings.reset()
        self.assertEqual((settings.destination_mode(), settings.destination_root()), ("live", self.path("SharePoint")))
        self.assertTrue(self.c.set_destination("test")["ok"])
        settings.reset()
        self.assertEqual(settings.destination_mode(), "test")
        self.assertTrue(os.path.exists(self.path("Bot_Settings.backup.xlsx")))

    def test_settings_open_in_excel_gives_a_friendly_message(self):
        os.makedirs(self.path("SP"))
        with mock.patch("openpyxl.workbook.workbook.Workbook.save", side_effect=PermissionError):         # what Windows does for a file open in Excel
            r = self.c.set_destination("live", self.path("SP"))
        self.assertFalse(r["ok"])
        self.assertIn("open in Excel", r["message"])
        settings.reset()
        self.assertEqual(settings.destination_mode(), "test")                  # nothing was half-changed

    # ---- mailboxes
    def test_detecting_and_saving_mailboxes(self):
        names = ["jiya@pearlenergy.com.au", "Pearl Energy Colyton", "Archive"]
        self.c.detect_cmd = [sys.executable, "-c", f"import json; print(json.dumps({{'outlook': {names!r}}}))"]
        r = self.c.detect_mailboxes()
        self.assertTrue(r["ok"], r)
        by = {s["outlook"]: s for s in r["suggestions"]}
        self.assertEqual(by["Pearl Energy Colyton"]["why"], "already set")
        self.assertEqual(by["jiya@pearlenergy.com.au"]["company"], "")
        self.assertEqual(len(r["companies"]), 13)
        self.assertNotIn("Pearl Energy Colyton", r["missing"])          # signed in on this PC
        self.assertTrue(r["missing"])                                   # the other companies are named as "not signed in here"
        saved = self.c.save_mailboxes([{"outlook": "jiya@pearlenergy.com.au", "action": "test"}, {"outlook": "Archive", "action": "skip"}])
        self.assertTrue(saved["ok"])
        settings.reset()
        row = next(x for x in settings.rows("Companies") if x["Mailbox (as in Outlook)"] == "jiya@pearlenergy.com.au")
        self.assertEqual((row["Folder name"], row["Active"], row["Text to look for"]), ("TEST JIYA", "Yes", "zz test only"))
        self.assertEqual(self.c.save_mailboxes([{"outlook": "jiya@pearlenergy.com.au", "action": "test"}])["messages"], ["jiya@pearlenergy.com.au: already in the Companies sheet"])

    def test_outlook_problems_are_explained(self):
        self.c.detect_cmd = [sys.executable, "-c", "import json; print(json.dumps({'error': 'Could not connect to Outlook: closed'}))"]
        self.assertIn("Could not connect", self.c.detect_mailboxes()["message"])
        self.c.detect_cmd = [sys.executable, "-c", "print('garbage')"]
        self.assertIn("did not answer", self.c.detect_mailboxes()["message"])

    def test_test_mailbox_files_go_to_its_own_folder(self):
        mailboxes.apply([{"outlook": "jiya@pearlenergy.com.au", "action": "test"}])
        gw = gateway("jiya@pearlenergy.com.au")
        gw.add("jiya@pearlenergy.com.au", mail("Tax Invoice 778899", ["invoice_bidfood.pdf"]))
        gw.add("jiya@pearlenergy.com.au", mail("FW: Invoice NS-4040", ["invoice_company_named.pdf"]))
        s = run(gw, only="jiya")
        self.assertEqual(s.c["Filed"], 2)
        files = self.files_under("TestOutput")
        self.assertIn("OCTOBER 2026 EMAIL INVOICES/TEST JIYA/Bidfood Australia Ltd_778899.pdf", files)
        self.assertIn("OCTOBER 2026 EMAIL INVOICES/RADHE KRISHNA INVESTMENT/Nova Snacks Pty Ltd_NS-4040.pdf", files)   # the document names it

    # ---- schedule
    def test_schedule_is_windows_only(self):
        if sys.platform.startswith("win"):
            self.skipTest("Windows")
        self.assertFalse(self.c.schedule_status()["supported"])
        self.assertIn("only works on Windows", self.c.set_schedule(True, 30)["message"])


class MailboxSuggestionTests(Isolated, unittest.TestCase):
    def test_suggestions(self):
        by = {s["outlook"]: s for s in mailboxes.suggest(["Tweed Heads", "Gopinath", "Shared Mailbox", "One@pearlenergy.com.au"])}
        self.assertEqual(by["Tweed Heads"]["company"], "Tweed Heads Energy Pty Ltd")
        self.assertEqual(by["Gopinath"]["company"], "GOPINATH INVESTMENT PTY LTD")
        self.assertEqual(by["Shared Mailbox"]["company"], "")
        self.assertEqual(by["One@pearlenergy.com.au"]["why"], "already set")

    def test_mapping_an_outlook_mailbox_to_a_company_keeps_the_old_names(self):
        mailboxes.apply([{"outlook": "Pearl Energy Tweed", "action": "company", "company": "Tweed Heads Energy Pty Ltd"}])
        settings.reset()
        row = next(r for r in settings.rows("Companies") if r["Company name"] == "Tweed Heads Energy Pty Ltd")
        self.assertEqual(row["Mailbox (as in Outlook)"].split("|")[0], "Pearl Energy Tweed")
        self.assertIn("pearl.tweedheads@pearlenergy.com.au", row["Mailbox (as in Outlook)"])


class WebServerTests(Isolated, unittest.TestCase):
    def setUp(self):
        super().setUp()
        self.addCleanup(shutil.rmtree, os.path.join(ROOT, "demo_run"), True)
        self.controller = panel.Controller(opener=lambda p: None)
        self.server, self.token = panel.make_server(self.controller, port=0)
        self.port = self.server.server_address[1]
        threading.Thread(target=self.server.serve_forever, daemon=True).start()
        self.addCleanup(self.server.server_close)
        self.addCleanup(self.server.shutdown)

    def call(self, path, body=None, token="good", host=None):
        headers = {"X-Token": self.token if token == "good" else (token or "")}
        if host:
            headers["Host"] = host
        data = None
        if body is not None:
            data = json.dumps(body).encode()
            headers["Content-Type"] = "application/json"
        req = urllib.request.Request(f"http://127.0.0.1:{self.port}{path}", data=data, headers=headers)
        try:
            with urllib.request.urlopen(req, timeout=30) as r:
                return r.status, r.read().decode()
        except urllib.error.HTTPError as e:
            return e.code, e.read().decode()

    def test_page_needs_the_secret(self):
        code, _ = self.call("/", token=None)
        self.assertEqual(code, 403)
        code, html = self.call(f"/?t={self.token}", token=None)
        self.assertEqual(code, 200)
        self.assertIn("Outlook Invoice Bot", html)
        self.assertIn(self.token, html)                                       # the page carries the secret for its own requests
        self.assertNotIn("__TOKEN__", html)

    def test_api_needs_the_secret_and_a_local_host_name(self):
        self.assertEqual(self.call("/api/overview", token="wrong")[0], 403)
        self.assertEqual(self.call("/api/overview", token=None)[0], 403)
        self.assertEqual(self.call("/api/run", {"kind": "demo"}, token="wrong")[0], 403)
        self.assertEqual(self.call("/api/overview", host="evil.example.com")[0], 403)      # DNS-rebinding style request
        code, body = self.call("/api/overview")
        self.assertEqual(code, 200)
        self.assertEqual(json.loads(body)["mode"], "test")

    def test_run_the_demo_through_the_web_api(self):
        code, body = self.call("/api/run", {"kind": "demo"})
        self.assertEqual((code, json.loads(body)["ok"]), (200, True))
        s = wait_until_finished(self.controller)
        code, body = self.call("/api/status?since=0")
        data = json.loads(body)
        self.assertEqual((code, data["running"], data["job"]["exit_code"]), (200, False, 0))
        self.assertGreaterEqual(data["summary"]["filed"], 5)
        self.assertEqual(json.loads(self.call("/api/run", {"kind": "dates"})[1])["ok"], False)

    def test_bad_requests_do_not_crash_the_server(self):
        self.assertEqual(self.call("/api/nothing")[0], 404)
        req = urllib.request.Request(f"http://127.0.0.1:{self.port}/api/run", data=b"{not json", headers={"X-Token": self.token})
        try:
            urllib.request.urlopen(req, timeout=10)
        except urllib.error.HTTPError as e:
            self.assertEqual(e.code, 500)
        self.assertEqual(self.call("/api/overview")[0], 200)                    # still alive

    @unittest.skipUnless(shutil.which("node"), "Node.js is not installed (it is only needed for this test)")
    def test_the_pages_javascript_works_against_the_real_server(self):
        """Runs panel_ui.html's script in Node with a fake DOM, clicks the buttons and checks what the page shows."""
        import subprocess
        names = ["jiya@pearlenergy.com.au", "Pearl Energy Colyton"]
        self.controller.detect_cmd = [sys.executable, "-c", f"import json; print(json.dumps({{'outlook': {names!r}}}))"]
        out = subprocess.run(["node", os.path.join(ROOT, "tests", "panel_dom_test.js"), str(self.port), self.token],
                             capture_output=True, text=True, timeout=240)
        self.assertEqual(out.returncode, 0, out.stdout + out.stderr)
        self.assertIn("DOM-OK", out.stdout)

    def test_quit(self):
        code, body = self.call("/api/quit", {})
        self.assertTrue(json.loads(body)["ok"])


if __name__ == "__main__":
    unittest.main()
