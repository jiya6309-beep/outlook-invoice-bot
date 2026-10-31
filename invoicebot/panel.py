"""The control panel: a small web page on THIS PC with big buttons (no commands to type).

    Start_Bot.bat  ->  app.py  ->  this module serves  http://127.0.0.1:<port>/?t=<secret>  and opens it in the browser.

Uses only Python's standard library (no extra install, no internet). Safe by design:
  * listens on 127.0.0.1 only (not reachable from other PCs);
  * every request needs a random secret token (other web pages cannot press the buttons);
  * the Host header must be 127.0.0.1 / localhost (blocks DNS-rebinding tricks);
  * a run is a separate process (python run_bot.py ...), so the page can never hang or corrupt a run.
Controller holds all the logic and is tested without any web server."""
import datetime
import json
import os
import re
import secrets
import sqlite3
import subprocess
import sys
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from . import __version__, exports, mailboxes, paths, quarterly, settings
from . import companies as companies_mod

IS_WINDOWS = sys.platform.startswith("win")
NO_WINDOW = 0x08000000 if IS_WINDOWS else 0
TASK_NAME = "Outlook Invoice Bot"
SECURITY_TASK = "Outlook Invoice Bot - Security"
WEEKDAYS = {"monday": "MON", "tuesday": "TUE", "wednesday": "WED", "thursday": "THU", "friday": "FRI", "saturday": "SAT", "sunday": "SUN"}
MAX_LINES = 20000

SUMMARY_KEYS = {
    "E-mails scanned": "scanned", "blocked (red flag)": "blocked", "downloaded": "downloaded", "image-only": "image_only",
    "no attachment": "no_attachment", "red category only": "category_only", "failed (will retry)": "failed", "already ticked or flagged in Outlook (not read again)": "already_ticked",
    "already handled by the bot before": "already_handled",
    "Documents": "documents", "filed to company folder": "filed", "parked (statements...)": "parked",
    "waiting in _Unassigned": "unassigned", "in _Unrecognised": "unrecognised", "duplicates skipped": "duplicates",
    "Released from waiting": "released", "Fixes applied from To_Fix": "fixes", "Outlook update problems": "outlook_problems"}
LAYOUTS = {"month_company_site": "{MONTH_FOLDER}\\{COMPANY}\\{SITE}", "site": "{SITE}", "company_site": "{COMPANY}\\{SITE}", "site_month": "{SITE}\\{MONTH} {YEAR}",
           "month_company": "{MONTH_FOLDER}\\{COMPANY}"}
HEADLINES = ("CANNOT START", "Not started", "STOPPED EARLY", "Could not connect", "The Python package", "DEMO", "WARNING")


def parse_summary(lines):
    """Numbers and the headline problem from the text a run prints."""
    numbers, headline = {}, ""
    for line in lines:
        m = re.match(r"^\s*([^:]+?)\s*:\s*(\d+)\s*$", line)
        if m and m.group(1) in SUMMARY_KEYS:
            numbers[SUMMARY_KEYS[m.group(1)]] = int(m.group(2))
        elif not headline and line.strip().startswith(HEADLINES[:5]):
            headline = line.strip()
    return numbers, headline


class Controller:
    """Everything the buttons do. No HTTP in here."""

    def __init__(self, opener=None):
        self.lock = threading.RLock()
        self.proc = None
        self.job = None
        self.lines = []
        self.opener = opener or self._open_with_os
        self.detect_cmd = [sys.executable, "-m", "invoicebot.mailboxes"]
        self.project = paths.project_dir()

    # ------------------------------------------------------------------ running the bot
    def running(self):
        return self.proc is not None and self.proc.poll() is None

    def args_for(self, kind, opts):
        """The command (after python) for a button. Raises ValueError with a friendly message."""
        run = os.path.join(self.project, "run_bot.py")
        extra = []
        mailbox = str(opts.get("mailbox") or "").strip()
        if mailbox:
            extra = ["--mailbox", mailbox]
        subject = str(opts.get("subject") or "").strip()
        if subject:
            extra += ["--subject", subject]
        if kind in ("security", "security_shadow", "security_dates"):          # the security run: never invoices, no mailbox filters
            cmd = [run, "--security"] + (["--shadow"] if kind == "security_shadow" or opts.get("shadow") else [])
            if kind == "security_dates":
                try:
                    a, b = datetime.date.fromisoformat(str(opts.get("from"))), datetime.date.fromisoformat(str(opts.get("to")))
                except ValueError:
                    raise ValueError("Please choose both dates (From and To).")
                if a > b:
                    raise ValueError("The From date is later than the To date.")
                cmd += ["--from", a.isoformat(), "--to", b.isoformat()]
            return cmd
        if kind == "live":
            return [run] + extra
        if kind == "shadow":
            return [run, "--shadow"] + extra
        if kind == "dates":
            try:
                a, b = datetime.date.fromisoformat(str(opts.get("from"))), datetime.date.fromisoformat(str(opts.get("to")))
            except ValueError:
                raise ValueError("Please choose both dates (From and To).")
            if a > b:
                raise ValueError("The From date is later than the To date.")
            return [run] + (["--shadow"] if opts.get("shadow") else []) + ["--from", a.isoformat(), "--to", b.isoformat()] + extra
        if kind == "report":
            try:
                a, b = datetime.date.fromisoformat(str(opts.get("from"))), datetime.date.fromisoformat(str(opts.get("to")))
            except ValueError:
                raise ValueError("Please choose both dates (From and To).")
            if a > b:
                raise ValueError("The From date is later than the To date.")
            return [run, "--report", "--from", a.isoformat(), "--to", b.isoformat()]
        if kind == "ac":
            try:
                a, b = datetime.date.fromisoformat(str(opts.get("from"))), datetime.date.fromisoformat(str(opts.get("to")))
            except ValueError:
                raise ValueError("Please choose both dates (From and To).")
            if a > b:
                raise ValueError("The From date is later than the To date.")
            return [run, "--ac", "--from", a.isoformat(), "--to", b.isoformat()]
        if kind == "match":
            return [run, "--match-folders"]
        if kind == "outlooktest":
            if not mailbox or not subject:
                raise ValueError("Fill in BOTH boxes: 'Only one mailbox?' (for example jiya) and 'Only touch e-mails whose subject contains' (for example BOTTEST). The test changes one e-mail and puts it back.")
            return [run, "--test-outlook-marking"] + extra
        if kind == "check":
            return [os.path.join(self.project, "check_setup.py")]
        if kind == "demo":
            return [run, "--demo"]
        if kind == "msg":
            folder = str(opts.get("folder") or "").strip().strip('"')
            if not folder or not os.path.isdir(folder):
                raise ValueError("That folder does not exist. Type the folder that contains the saved .msg e-mails.")
            if not any(f.lower().endswith(".msg") for f in os.listdir(folder)):
                raise ValueError("There are no .msg files in that folder.")
            return [run, "--msg-folder", folder] + (["--msg-mailbox", mailbox] if mailbox else [])
        raise ValueError(f"unknown action '{kind}'")

    def start(self, kind, opts=None):
        opts = opts or {}
        with self.lock:
            if self.running():
                return {"ok": False, "message": "A run is already in progress. Wait for it to finish (or press Stop)."}
            try:
                args = self.args_for(kind, opts)
            except ValueError as e:
                return {"ok": False, "message": str(e)}
            env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUNBUFFERED="1")
            self.proc = subprocess.Popen([sys.executable] + args, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True,
                                         encoding="utf-8", errors="replace", bufsize=1, cwd=self.project, env=env,
                                         creationflags=NO_WINDOW)
            self.lines = []
            self.job = {"kind": kind, "started": time.time(), "finished": None, "exit_code": None, "stopped": False}
            threading.Thread(target=self._pump, args=(self.proc, self.job), daemon=True).start()
            return {"ok": True, "message": "Started."}

    def _pump(self, proc, job):
        for line in proc.stdout:
            with self.lock:
                self.lines.append(line.rstrip("\r\n"))
                if len(self.lines) > MAX_LINES:
                    del self.lines[:1000]
        code = proc.wait()
        try:
            proc.stdout.close()                                  # one handle per run would otherwise leak in a panel left open all day
        except OSError:
            pass
        with self.lock:
            job["exit_code"], job["finished"] = code, time.time()

    def stop(self):
        with self.lock:
            if not self.running():
                return {"ok": False, "message": "Nothing is running."}
            self.job["stopped"] = True
            pid = self.proc.pid
        if IS_WINDOWS:                                          # the run has reader processes: stop the whole tree
            subprocess.run(["taskkill", "/PID", str(pid), "/T", "/F"], capture_output=True, creationflags=NO_WINDOW)
        else:
            self.proc.terminate()
        try:
            self.proc.wait(timeout=15)
        except subprocess.TimeoutExpired:
            self.proc.kill()
        try:
            os.remove(paths.p("data", "run.lock"))               # otherwise the next run would wait for the stale lock
        except OSError:
            pass
        return {"ok": True, "message": "Stopped. Nothing is lost: the next run continues from where this one reached."}

    def status(self, since=0):
        with self.lock:
            lines = self.lines[since:]
            job = dict(self.job) if self.job else None
            nums, headline = parse_summary(self.lines) if job and not self.running() else ({}, "")
            return {"running": self.running(), "job": job, "lines": lines, "next": since + len(lines),
                    "summary": nums, "headline": headline}

    # ------------------------------------------------------------------ overview
    def overview(self):
        info = {"version": __version__, "project": paths.home(), "running": self.running(), "mode": None, "root": "",
                "holding": "", "settings_error": "", "counts": {}, "last_run": None, "reports": {}, "schedule": self.schedule_status()}
        try:
            settings.reset()
            info["mode"], info["root"], info["holding"] = settings.destination_mode(), settings.destination_root(), settings.holding_root()
            info["layout"] = settings.run_value("Destination pattern", LAYOUTS["month_company"])
            info["switched_off"] = self.switched_off()
            info["live_root_set"] = bool(settings.run_value("Live destination root") or settings._config_value("LIVE_DESTINATION_ROOT"))
            info["companies"] = [{"name": r["Company name"], "folder": r["Folder name"], "mailbox": r["Mailbox (as in Outlook)"]}
                                 for r in settings.rows("Companies")]
        except settings.SettingsError as e:
            info["settings_error"] = str(e)
        dbfile = paths.p("data", "bot.db")
        if os.path.exists(dbfile):
            try:
                conn = sqlite3.connect(Path(dbfile).as_uri() + "?mode=ro", uri=True)
                today = datetime.date.today().isoformat()
                # the dashboard shows TODAY only (documents first seen today), not the all-time totals
                info["counts"] = dict(conn.execute("SELECT status, COUNT(*) FROM docs WHERE substr(created_at,1,10)=? GROUP BY status", (today,)).fetchall())
                info["today"] = conn.execute("SELECT COUNT(*) FROM docs WHERE substr(created_at,1,10)=?", (today,)).fetchone()[0]
                info["waiting_total"] = conn.execute("SELECT COUNT(*) FROM docs WHERE status='Unassigned'").fetchone()[0]
                info["unrecognised_total"] = conn.execute("SELECT COUNT(*) FROM docs WHERE status='Unrecognised'").fetchone()[0]
                row = conn.execute("SELECT started, mode FROM runs ORDER BY id DESC LIMIT 1").fetchone()
                info["last_run"] = {"started": row[0], "mode": row[1]} if row else None
                conn.close()
            except sqlite3.Error:
                pass
        try:
            rep = settings.reports_root()
        except settings.SettingsError:
            rep = paths.p("Reports")
        month = Path(os.path.join(rep, f"Record_{datetime.date.today():%Y-%m}.xlsx"))
        info["reports"] = {"to_fix": os.path.exists(os.path.join(rep, "To_Fix.xlsx")), "record": month.exists(), "ac": bool(self._newest_ac(rep))}
        return info

    # ------------------------------------------------------------------ opening folders and files
    def _targets(self):
        logs = paths.p("logs")
        newest = None
        if os.path.isdir(logs):
            files = sorted((os.path.join(logs, f) for f in os.listdir(logs) if f.startswith("run_")), key=os.path.getmtime)
            newest = files[-1] if files else None
        today = datetime.date.today()
        rep = settings.reports_root()
        hold = settings.holding_root()
        dest = settings.destination_root()
        month_folder = ""
        if dest:
            from .router import month_folder_name, folder_clean
            month_folder = os.path.join(dest, folder_clean(month_folder_name(datetime.datetime.now())))
        return {"reports": rep, "to_fix": os.path.join(rep, "To_Fix.xlsx"),
                "record": os.path.join(rep, f"Record_{today:%Y-%m}.xlsx"), "ac": self._newest_ac(rep),
                "unassigned": os.path.join(hold, "_Unassigned"), "unrecognised": os.path.join(hold, "_Unrecognised"),
                "statements": os.path.join(hold, "_Statements"), "filed": month_folder if month_folder and os.path.isdir(month_folder) else dest,
                "logs": logs, "last_log": newest,
                "settings": settings.settings_file(), "project": self.project, "readme": os.path.join(self.project, "README.md"),
                "holding": settings.holding_root(), "destination": settings.destination_root(),
                "records_archive": quarterly.folder(),
                "security_excel": self._newest_security_excel(), "security_folder": self._security_root(),
                "security_settings": os.path.join(self.project, "Security_Settings.xlsx"),
                "fuel_list": exports.fuel_list_path(today.year, today.month), "fleet_record": exports.fleet_record_path(today.year, today.month),
                "fuel_folder": self._special_root("fuel invoices"), "fleet_folder": self._special_root("fleet card")}

    def _special_root(self, role):
        from . import companies
        for sp in companies.special_rows():
            if sp["role"] == role:
                return sp["dest"]
        return ""

    @staticmethod
    def _newest_ac(folder):
        try:
            files = [os.path.join(folder, f) for f in os.listdir(folder) if f.startswith("AC_") and f.endswith(".xlsx")]
        except OSError:
            return None
        return max(files, key=os.path.getmtime) if files else None

    @staticmethod
    def _open_with_os(path):
        if IS_WINDOWS:
            os.startfile(path)                                    # noqa: S606 (Windows only)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])

    def open_path(self, what):
        try:
            target = self._targets().get(what)
        except settings.SettingsError as e:
            return {"ok": False, "message": str(e)}
        if what not in self._targets():
            return {"ok": False, "message": "unknown item"}
        if not target:
            return {"ok": False, "message": "Not available yet - run the bot first."}
        if what in ("reports", "logs", "records_archive", "unassigned", "unrecognised", "statements") and not os.path.exists(target):
            os.makedirs(target, exist_ok=True)
        if not os.path.exists(target):
            return {"ok": False, "message": "It does not exist yet" + (" - run the bot first." if what in ("to_fix", "record", "last_log") else
                                                                    " - make one with 'Make the A/C Excel for these dates'." if what == "ac" else ".")}
        try:
            self.opener(target)
        except Exception as e:
            return {"ok": False, "message": f"Could not open it: {e}"}
        return {"ok": True, "message": target}

    # ------------------------------------------------------------------ settings: destination
    def set_destination(self, mode, root="", layout=None):
        mode, root = str(mode).lower(), str(root or "").strip().strip('"')
        if layout and layout not in LAYOUTS:
            return {"ok": False, "message": "Unknown folder layout."}
        if mode not in ("test", "live"):
            return {"ok": False, "message": "Mode must be test or live."}
        auto = False
        if mode == "live" and not root:
            from . import autofind
            try:
                known = settings.destination_root()                       # a path already typed in the workbook / config.py / found earlier
            except settings.SettingsError as e:
                return {"ok": False, "message": str(e)}
            hits = [] if (known and os.path.isdir(known)) else autofind.find()
            if known and os.path.isdir(known):
                root = known
            elif hits:
                root, auto = hits[0], True
                autofind.remember(root)
            else:
                return {"ok": False, "message": "The synced SharePoint folder was not found automatically (looking for '" + settings.run_value("Auto-find folder name", "EMAIL INVOICES") + "' inside '" + settings.run_value("Auto-find parent folder", "OFFICE INDIA") + "' in C:\\SHAREPOINT and under your user folder). "
                                                "Make sure OneDrive has finished syncing it, or type the folder path."}
        if mode == "live":
            if not os.path.isdir(root):
                return {"ok": False, "message": "That folder was not found on this PC. Is OneDrive synced and signed in? Check the path."}
        try:
            def change(wb):
                if mode == "live" and not auto and root != settings.destination_root():
                    settings.set_run_value(wb, "Live destination root", root)
                settings.set_run_value(wb, "Destination mode", mode)
                if layout:
                    settings.set_run_value(wb, "Destination pattern", LAYOUTS[layout])
            settings.edit(change)
        except settings.SettingsError as e:
            return {"ok": False, "message": str(e)}
        where = ("  Layout: " + LAYOUTS[layout]) if layout else ""
        return {"ok": True, "message": (("LIVE: documents now go to " + root + ("   (found automatically)" if auto else "")) if mode == "live" else "TEST: documents go to the TestOutput folder.") + where}

    # ------------------------------------------------------------------ mailboxes
    def detect_mailboxes(self):
        try:
            out = subprocess.run(self.detect_cmd, capture_output=True, text=True, timeout=90, cwd=self.project,
                                 creationflags=NO_WINDOW, env=dict(os.environ, PYTHONIOENCODING="utf-8"))
            data = json.loads((out.stdout.strip().splitlines() or ["{}"])[-1])
        except (subprocess.TimeoutExpired, json.JSONDecodeError, OSError):
            return {"ok": False, "message": "Outlook did not answer. Make sure classic Outlook is open and signed in, then try again."}
        if data.get("error"):
            return {"ok": False, "message": data["error"]}
        try:
            settings.reset()
            suggestions = mailboxes.suggest(data.get("outlook", []))
            companies = [r["Company name"] for r in settings.rows("Companies") if r["Text to look for"] != mailboxes.TEST_TEXT]
        except settings.SettingsError as e:
            return {"ok": False, "message": str(e)}
        # companies that are switched on but whose mailbox is not signed in on this PC (the run would skip them)
        missing = []
        try:
            outlook_names = [settings.norm(n) for n in data.get("outlook", [])]
            for r in settings.rows("Companies"):
                if r["Text to look for"] == mailboxes.TEST_TEXT or str(r.get("Active", "Yes")).strip().lower() not in ("yes", "y", "true", "1", ""):
                    continue
                wanted = [settings.norm(n) for n in companies_mod.mailbox_names(r)]
                if wanted and not any(w and o and (w == o or w in o or o in w) for w in wanted for o in outlook_names):
                    missing.append(r["Company name"])
        except Exception:
            missing = []
        inactive = [r["Company name"] for r in settings.rows("Companies")
                    if r["Text to look for"] != mailboxes.TEST_TEXT and not companies_mod._active(r)]
        return {"ok": True, "suggestions": suggestions, "companies": companies, "missing": missing, "inactive": inactive}

    def switched_off(self):
        """Companies whose Active is No: they are NOT read by any run (shown as a warning banner in the panel)."""
        try:
            return [r["Company name"] for r in settings.rows("Companies")
                    if r["Text to look for"] != mailboxes.TEST_TEXT and not companies_mod._active(r)]
        except Exception:
            return []

    def activate_all(self):
        def change(wb):
            ws = wb["Companies"]
            head = {str(c.value or "").strip(): i + 1 for i, c in enumerate(ws[1])}
            for r in range(2, ws.max_row + 1):
                if str(ws.cell(r, head["Text to look for"]).value or "").strip() != mailboxes.TEST_TEXT:
                    ws.cell(r, head["Active"]).value = "Yes"
        try:
            settings.edit(change)
            settings.reset()
            return {"ok": True, "message": "All companies are switched on again: every run reads them all."}
        except settings.SettingsError as e:
            return {"ok": False, "message": str(e)}

    def save_mailboxes(self, items):
        try:
            return {"ok": True, "messages": mailboxes.apply(items) or ["Nothing to save."]}
        except settings.SettingsError as e:
            return {"ok": False, "message": str(e)}

    # ------------------------------------------------------------------ security (its own schedule and its own numbers)
    def security_overview(self):
        """The latest week of security reports that was read: how many sites, late openings, early closings, missing, no report."""
        from . import db, security, security_settings as ss
        out = {"week": "", "reports": 0, "ok": 0, "open_late": 0, "close_early": 0, "missing": 0, "no_report": 0, "excel": ""}
        dbfile = paths.p("data", "bot.db")
        if not os.path.exists(dbfile):
            return out
        try:
            conf = ss.load()
            conn = db.connect(dbfile)
            security.ensure_schema(conn)
            row = conn.execute("SELECT MAX(week) FROM security_reports").fetchone()
            if not row or not row[0]:
                conn.close()
                return out
            wk = datetime.date.fromisoformat(row[0])
            out["week"] = f"{wk:%d-%m} TO {wk + datetime.timedelta(days=6):%d-%m}"
            out["reports"] = conn.execute("SELECT COUNT(*) FROM security_reports WHERE week=?", (row[0],)).fetchone()[0]
            rules = (ss.number(conf["run"], "Late opening allowance (minutes)", 10), ss.number(conf["run"], "Late arming window (hours)", 4),
                     ss.number(conf["run"], "Early disarm window (hours)", 3.5))
            sites = [x for x in conf["sites"] if x["active"]]
            cover = security.coverage(conn, sites, wk, wk)
            for site in sites:
                res = security.day_results(conn, site, wk, wk + datetime.timedelta(days=6), rules,
                                           {w: v.get(site["client"], False) for w, v in cover.items()})
                for r in res.values():
                    if r["state"] == "no report":
                        out["no_report"] += 1
                    elif r["state"] == "ok":
                        if not r["in"] or not r["out"]:
                            out["missing"] += 1
                        if r["in"] and r["open_flag"]:
                            out["open_late"] += 1
                        if r["out"] and r["close_flag"]:
                            out["close_early"] += 1
                        if not r["open_flag"] and not r["close_flag"]:
                            out["ok"] += 1
            conn.close()
            out["excel"] = self._newest_security_excel() or ""
        except Exception as e:
            out["error"] = str(e)
        return out

    def _security_root(self):
        from . import companies
        for sp in companies.special_rows():
            if sp["role"] == "security":
                return sp["dest"]
        return ""

    def _newest_security_excel(self):
        root = self._security_root()
        best = None
        for base, _dirs, files in os.walk(root) if root and os.path.isdir(root) else []:
            for f in files:
                if f.startswith("SITE_DATA_") and f.endswith(".xlsx"):
                    p = os.path.join(base, f)
                    if best is None or os.path.getmtime(p) > os.path.getmtime(best):
                        best = p
        return best

    def security_schedule_status(self):
        if not IS_WINDOWS:
            return {"supported": False, "enabled": False}
        r = subprocess.run(["schtasks", "/Query", "/TN", SECURITY_TASK], capture_output=True, text=True, creationflags=NO_WINDOW)
        return {"supported": True, "enabled": r.returncode == 0}

    def set_security_schedule(self, enable, day="Monday", at="09:00"):
        """Windows Task Scheduler, once a week: run_bot.bat --security (separate from the invoice task)."""
        if not IS_WINDOWS:
            return {"ok": False, "message": "Automatic running is set up with Windows Task Scheduler, so it only works on Windows."}
        code = WEEKDAYS.get(str(day or "").strip().lower())
        try:
            hh, mm = (int(x) for x in str(at or "").strip().split(":"))
            assert 0 <= hh < 24 and 0 <= mm < 60
        except Exception:
            return {"ok": False, "message": "Time must look like 09:00."}
        if enable and not code:
            return {"ok": False, "message": "Choose a day of the week."}
        bat = os.path.join(self.project, "run_bot.bat")
        cmd = (["schtasks", "/Create", "/TN", SECURITY_TASK, "/TR", f'"{bat}" --security', "/SC", "WEEKLY", "/D", code,
                "/ST", f"{hh:02d}:{mm:02d}", "/F"] if enable else ["schtasks", "/Delete", "/TN", SECURITY_TASK, "/F"])
        r = subprocess.run(cmd, capture_output=True, text=True, creationflags=NO_WINDOW)
        if r.returncode != 0:
            return {"ok": False, "message": (r.stderr or r.stdout).strip()[:300] or "Windows refused the request."}
        return {"ok": True, "message": f"The security check will run every {day} at {hh:02d}:{mm:02d} (Outlook must be open)."
                if enable else "The automatic security check is switched off."}

    # ------------------------------------------------------------------ schedule (Windows Task Scheduler)
    def schedule_status(self):
        if not IS_WINDOWS:
            return {"supported": False, "enabled": False}
        r = subprocess.run(["schtasks", "/Query", "/TN", TASK_NAME], capture_output=True, text=True, creationflags=NO_WINDOW)
        return {"supported": True, "enabled": r.returncode == 0}

    def set_schedule(self, enable, minutes=30):
        if not IS_WINDOWS:
            return {"ok": False, "message": "Automatic running is set up with Windows Task Scheduler, so it only works on Windows."}
        try:
            minutes = max(5, min(int(minutes), 1440))
        except (TypeError, ValueError):
            return {"ok": False, "message": "Minutes must be a number (5 or more)."}
        bat = os.path.join(self.project, "run_bot.bat")
        cmd = (["schtasks", "/Create", "/TN", TASK_NAME, "/TR", f'"{bat}"', "/SC", "MINUTE", "/MO", str(minutes), "/F"]
               if enable else ["schtasks", "/Delete", "/TN", TASK_NAME, "/F"])
        r = subprocess.run(cmd, capture_output=True, text=True, creationflags=NO_WINDOW)
        if r.returncode != 0:
            return {"ok": False, "message": (r.stderr or r.stdout).strip()[:300] or "Windows refused the request."}
        return {"ok": True, "message": f"The bot will run every {minutes} minutes while you are signed in to Windows (Outlook must be open)."
                if enable else "Automatic running is switched off."}


# ---------------------------------------------------------------------------------------- web server

def make_server(controller, host="127.0.0.1", port=0):
    """Returns (server, token). The caller runs server.serve_forever()."""
    token = secrets.token_urlsafe(24)
    html = (Path(__file__).with_name("panel_ui.html")).read_text(encoding="utf-8")

    class Handler(BaseHTTPRequestHandler):
        server_version = "InvoiceBotPanel"

        def log_message(self, *args):
            pass

        def _send(self, code, body, ctype="application/json; charset=utf-8"):
            data = body if isinstance(body, bytes) else body.encode("utf-8")
            self.send_response(code)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(len(data)))
            self.send_header("Cache-Control", "no-store")
            self.send_header("X-Content-Type-Options", "nosniff")
            self.end_headers()
            self.wfile.write(data)

        def _json(self, code, obj):
            self._send(code, json.dumps(obj))

        def _allowed(self, query):
            host_ok = self.headers.get("Host", "").split(":")[0] in ("127.0.0.1", "localhost")
            given = (query.get("t") or [""])[0] or self.headers.get("X-Token", "")
            return host_ok and secrets.compare_digest(str(given), token)

        def do_GET(self):
            url = urlparse(self.path)
            query = parse_qs(url.query)
            if not self._allowed(query):
                return self._send(403, "<h3>Not allowed</h3><p>Open the control panel with Start_Bot.bat.</p>", "text/html; charset=utf-8")
            try:
                if url.path == "/":
                    return self._send(200, html.replace("__TOKEN__", token).replace("__VERSION__", __version__), "text/html; charset=utf-8")
                if url.path == "/api/overview":
                    return self._json(200, controller.overview())
                if url.path == "/api/security":
                    return self._json(200, dict(controller.security_overview(), schedule=controller.security_schedule_status()))
                if url.path == "/api/status":
                    return self._json(200, controller.status(int((query.get("since") or ["0"])[0] or 0)))
                if url.path == "/api/mailboxes":
                    return self._json(200, controller.detect_mailboxes())
            except Exception as e:
                return self._json(500, {"ok": False, "message": f"{type(e).__name__}: {e}"})
            self._json(404, {"ok": False, "message": "not found"})

        def do_POST(self):
            url = urlparse(self.path)
            if not self._allowed(parse_qs(url.query)):
                return self._json(403, {"ok": False, "message": "not allowed"})
            try:
                size = min(int(self.headers.get("Content-Length") or 0), 200_000)
                body = json.loads(self.rfile.read(size) or b"{}")
                route = url.path
                if route == "/api/run":
                    return self._json(200, controller.start(body.get("kind", ""), body))
                if route == "/api/stop":
                    return self._json(200, controller.stop())
                if route == "/api/open":
                    return self._json(200, controller.open_path(body.get("what", "")))
                if route == "/api/destination":
                    return self._json(200, controller.set_destination(body.get("mode", ""), body.get("root", ""), body.get("layout")))
                if route == "/api/mailboxes/save":
                    return self._json(200, controller.save_mailboxes(body.get("items", [])))
                if route == "/api/schedule":
                    return self._json(200, controller.set_schedule(bool(body.get("enable")), body.get("minutes", 30)))
                if route == "/api/companies/activate_all":
                    return self._json(200, controller.activate_all())
                if route == "/api/security/schedule":
                    return self._json(200, controller.set_security_schedule(bool(body.get("enable")), body.get("day", "Monday"),
                                                                            body.get("time", "09:00")))
                if route == "/api/quit":
                    threading.Thread(target=self.server.shutdown, daemon=True).start()
                    return self._json(200, {"ok": True, "message": "Control panel closed. You can close this tab."})
            except Exception as e:
                return self._json(500, {"ok": False, "message": f"{type(e).__name__}: {e}"})
            self._json(404, {"ok": False, "message": "not found"})

    server = ThreadingHTTPServer((host, port), Handler)
    server.daemon_threads = True
    return server, token
