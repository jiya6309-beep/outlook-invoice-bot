"""One complete bot run.

    1. take the lock, load the settings once, open the database
    2. apply the fixes staff typed in To_Fix.xlsx, and retry the files waiting in _Unassigned
    3. for every mailbox:  work out the date window -> scan e-mails (newest first, bounded)
         classify each e-mail:  blocked word | wanted | nothing
         wanted e-mails are handled in small batches:
            download the attachments to the local scratch folder
            read them (several reader processes in parallel)
            decide where each goes, write it there (verified, never overwritten)
            save the results in the database
            only then update Outlook (read + tick, categories, flags)
    4. exports (Excel), housekeeping (archive, backup), run summary

SHADOW MODE does everything except change anything: no Outlook change, no file written to a destination or
holding folder, a throw-away database. It produces a report of what a real run WOULD do.
Every loop is bounded (e-mails, attachments, run time, retries).
"""
import collections
import datetime
import multiprocessing
import os
import re
import time
from concurrent.futures import ProcessPoolExecutor

from . import (companies, db, doctypes, exports, memory, naming, notify, ocr, paths, quarterly, settings, sites, worker)
from .logs import Logger
from .settings import SettingsError
from . import router as router_mod
from .router import ConfigError, PlacementError, Router, UNASSIGNED_DIR, UNRECOGNISED_DIR
from .runlock import LockBusy, RunLock

DOC_EXT = (".pdf", ".xlsx", ".xlsm", ".xls")
IMG_EXT = (".png", ".jpg", ".jpeg")
FWRE = re.compile(r"^\s*(fw|fwd|re)\s*:", re.IGNORECASE)
MAX_EMAIL_ATTEMPTS = 3
PARALLEL_TIMEOUT_SECONDS = 120       # the longest a reader process may take for one file before it is given up on


# ----------------------------------------------------------------------------------------- helpers

def month_folder_clean(when):
    from .router import folder_clean, month_folder_name
    return folder_clean(month_folder_name(when))


def router_months(d):
    from .router import MONTHS
    return MONTHS[d.month - 1]


def parse_date(text):
    try:
        return datetime.date.fromisoformat(str(text).strip()[:10]) if str(text).strip() else None
    except ValueError:
        return None


def compute_window(now, from_date=None, to_date=None, last_success=None, default_days=2, max_days=31, overlap_minutes=60):
    """The period of e-mails to read. Returns (start, end, kind).
    With dates: From 00:00 to To 23:59:59, both inclusive (never beyond 'now').
    Without: automatic catch-up from the end of the last successful run (minus an overlap), capped."""
    if from_date or to_date:
        start = datetime.datetime.combine(from_date or to_date, datetime.time.min)
        end = datetime.datetime.combine(to_date, datetime.time(23, 59, 59)) if to_date else now
        return start, min(end, now), "dates"
    start = (last_success - datetime.timedelta(minutes=overlap_minutes)) if last_success \
        else now - datetime.timedelta(days=default_days)
    start = max(start, now - datetime.timedelta(days=max_days))
    return start, now, "catch-up"


def _reason_key(text):
    """A reason without the numbers and names in it, so the same reason counts together."""
    t = re.sub(r"'[^']*'", "'...'", str(text))
    t = re.sub(r"\d[\d\-/.,:]*", "#", t)
    return re.sub(r"\s+", " ", t).strip()[:80]


class Summary:
    def __init__(self, shadow=False):
        self.shadow = shadow
        self.c = collections.Counter()
        self.blocked_words = collections.Counter()
        self.new_suppliers, self.notes, self.mailbox_lines, self.warnings = [], [], [], []
        self.reasons, self.first_outlook_error = collections.Counter(), ""
        self.mailbox_data = []          # (mailbox, e-mails, from, to, kind, blocked, document e-mails, documents, note)
        self.peak_mb, self.worker_peak_mb, self.stop_reason, self.t0 = None, None, "", time.time()

    def text(self):
        c = self.c
        lines = [f"{'SHADOW RUN (nothing was changed)' if self.shadow else 'RUN SUMMARY'}",
                 "=" * 60]
        lines += self.mailbox_lines
        lines += ["",
                  f"Mailboxes read           : {c['mailboxes read']} ({c['quiet mailboxes']} had no new e-mail)",
                  f"E-mails scanned          : {c['emails scanned']}",
                  f"  blocked (red flag)     : {c['blocked']}",
                  f"  downloaded             : {c['emails downloaded']}",
                  f"  image-only             : {c['image-only']}",
                  f"  no attachment          : {c['no-attachment']}",
                  f"  red category only      : {c['category-only']}",
                  f"  forwards/replies left unread (no tick): {c['forwards left unread']}",
                  f"  already ticked or flagged in Outlook (not read again): {c['already ticked']}",
                  f"  already handled by the bot before: {c['already handled']}",
                  f"  failed (will retry)    : {c['email errors']}",
                  f"Documents                : {c['documents']}",
                  f"  filed to company folder: {c['Filed']}",
                  f"  parked (statements...) : {c['Parked']}",
                  f"  waiting in _Unassigned : {c['Unassigned']}",
                  f"  in _Unrecognised       : {c['Unrecognised']}",
                  f"  duplicates skipped     : {c['duplicates']}",
                  f"Released from waiting    : {c['released']}",
                  f"Fixes applied from To_Fix: {c['fixes']}",
                  f"Future-dated invoices moved to their month: {c['future invoices moved']}",
                  f"Outlook update problems  : {c['outlook errors']}"]
        if c["outlook errors"] and self.first_outlook_error:
            lines.append(f"  first problem: {self.first_outlook_error}")
        if self.reasons:
            lines += ["", "Why documents are waiting or unrecognised (this run):"] + [f"  {n:4} x {why}" for why, n in self.reasons.most_common(8)]
        if self.new_suppliers:
            lines += ["", "First-time supplier names (check them): " + "; ".join(n for n, _ in self.new_suppliers[:15])]
        if self.stop_reason:
            lines += ["", f"STOPPED EARLY: {self.stop_reason} (the rest is picked up next run)"]
        for w in self.warnings[:10]:
            lines.append(f"WARNING: {w}")
        for n in self.notes:
            lines.append(n)
        lines += ["", f"Time: {time.time() - self.t0:.0f} s   Peak memory: main process {self.peak_mb} MB, largest reader process {self.worker_peak_mb} MB"]
        return "\n".join(lines)


# ----------------------------------------------------------------------------------------- the run

class Pipeline:
    def __init__(self, gateway, shadow=False, from_date=None, to_date=None, only=None, now=None,
                 log=None, workers=None, db_path=None, quiet=False, subject_filter=None):
        self.gw, self.shadow = gateway, shadow
        self.from_date, self.to_date, self.only = from_date, to_date, only
        self.subject_filter = (subject_filter or "").strip().lower()      # testing aid: only e-mails whose subject contains this
        self.now = now or datetime.datetime.now().replace(microsecond=0)
        self.log = log or Logger(to_file=True, quiet=quiet)
        self.workers_override, self.db_path = workers, db_path
        self.summary = Summary(shadow)
        self.conn, self.router, self.pool = None, None, None
        self.seen, self.known = {}, set()
        self.report_rows, self.would, self.skipped_sample = [], collections.defaultdict(list), []
        self.touched_months, self.filed_now = set(), []
        self.pool_tasks, self.parallel_timeout = 0, PARALLEL_TIMEOUT_SECONDS
        self.changed = False                    # did this run change anything the Excel files show?
        self.retry_pending = collections.Counter()
        self.roles = {}                         # mailbox label -> (special role, destination folder)
        self.run_last = {}                      # mailbox -> (received, ref) of its newest e-mail with documents in this run
        self.started_iso = db.now()             # documents changed since this moment = changed by this run

    # ------------------------------------------------------------------------ setup / teardown
    def _setup(self):
        settings.reset()
        self.snap = settings.snapshot()
        ocr.configure(settings.run_bool("OCR enabled", True), settings.run_value("Tesseract path"),
                      settings.run_int("OCR max pages", 3))
        self._find_destination()
        self.router = Router(dry=self.shadow, log=self.log)
        self.router.check_config()
        paths.ensure("Scratch")
        if not self.shadow:
            self.router.ensure_holding_folders()
        if not self.shadow:
            added = settings.upgrade_workbook()                   # an older Bot_Settings.xlsx gets the new sheets / settings (nothing existing changes)
            if added:
                self.summary.notes.append("Bot_Settings.xlsx was brought up to date: " + "; ".join(added[:12]))
        self.conn = db.connect(":memory:" if self.shadow else (self.db_path or paths.p("data", "bot.db")))
        self.probe_ids = set()                                    # e-mails opened only because an attachment's file name says "order"
        self.seen = {} if self.shadow else db.seen_emails(self.conn)
        self.known = {settings.supplier_key(s) for s in db.known_suppliers(self.conn)} | {
            settings.supplier_key(r["Record name"]) for r in settings.rows("Suppliers") if r["Record name"]} | {
            settings.supplier_key(r["Supplier name"]) for r in settings.rows("Senders") if r.get("Supplier name")} | {
            settings.supplier_key(r["Supplier (record name)"]) for r in settings.rows("SupplierIDs") if r.get("Supplier (record name)")}
        cfg = settings.run_value
        self.conf = dict(
            max_emails=settings.run_int("Max emails per mailbox per run", 5000),
            max_atts=settings.run_int("Max attachments per email", 20),
            max_minutes=settings.run_int("Max run minutes", 45),
            batch=max(1, settings.run_int("Batch size (emails)", 25)),
            max_mb=settings.run_int("Max attachment MB", 25),
            image_kb=settings.run_int("Image min KB", 40),
            default_days=settings.run_int("Default catch-up days", 2),
            max_days=settings.run_int("Catch-up max days", 31),
            overlap=settings.run_int("Overlap minutes", 60),
            low_ram=settings.run_int("Min free RAM MB (use 1 worker)", 1500),
            stop_ram=settings.run_int("Min free RAM MB (stop run)", 500),
            archive_days=settings.run_int("Archive days", 7),
            max_wait=settings.run_int("Max attempts for waiting files", 50),
            flag_blocked_no_att=settings.run_bool("Flag blocked emails without attachment", True),
            scan_body=settings.run_bool("Scan body for blocked words", False),
            hold_first=settings.run_bool("Hold first-time suppliers", False),
            red_cat=cfg("Red category name", "Forwarded / Reply"), img_cat=cfg("Image category name", "Image attachment"),
            red_all=cfg("Red category applies to", "document").lower() == "all",
            fwre_download=settings.run_bool("Download attachments of FW/RE e-mails", True),
            skip_ticked=settings.run_bool("Skip e-mails already ticked in Outlook", True),
            order_probe=settings.run_bool("Check attachment names for order confirmations", True),
            month_basis=cfg("Month basis", "arrival").lower(),
            month_follow=cfg("Month folder follows", "window").lower(), make_trees=settings.run_bool("Create next month's folders", True), month_window=max(0, settings.run_int("Month window (months)", 3)),
            path_mode=cfg("Email path mode", "off").lower(),
            name_date=cfg("File name date format", naming.DEFAULT_DATE_FORMAT),
            site_in_company=settings.run_bool("Choose the site within the company first", True),
            site_fallback=settings.run_bool("File in the company folder when the site is not found", True),
        )
        if self.conf["month_basis"] == "invoice":                  # an older setting that asked for the invoice's own month keeps working
            self.conf["month_follow"] = "invoice"
        if self.from_date is None and self.to_date is None:
            self.from_date, self.to_date = parse_date(cfg("From date")), parse_date(cfg("To date"))
        self.deadline = time.time() + self.conf["max_minutes"] * 60
        self.n_workers = self._workers()
        self.roles = {r["Company name"]: (r.get("_role", ""), r.get("_dest", "")) for r in companies.reading_list()}

    def _find_destination(self):
        """LIVE mode without a typed path: look for the synced folder on this PC (a rehearsal only remembers it for this run)."""
        root = settings.destination_root() if settings.destination_mode() == "live" else ""
        if settings.destination_mode() != "live" or (root and os.path.isdir(root)):
            return
        if not settings.run_bool("Find the destination folder automatically", True):
            return
        from . import autofind
        self.log(("The LIVE folder was not found on this PC (" + root + ")" if root else "No destination path is set")
                 + " - looking for the synced folder on this PC ...")
        found, hits = autofind.discover(remember_it=not self.shadow)
        if found:
            self.log(f"Destination folder found automatically: {found}")
            if len(hits) > 1:
                self.summary.warnings.append("several folders match; the first was used: " + " | ".join(hits[:4]))
        else:
            self.log("The synced folder was not found.")

    def _workers(self):
        if self.workers_override is not None:
            return max(1, int(self.workers_override))
        raw = settings.run_value("Workers", "auto").lower()
        if raw.isdigit():
            return max(1, int(raw))
        return max(1, min(4, memory.cpu_count() - 1))

    def _pool(self):
        """The reader processes. Started with 'spawn' on every system (the same as Windows), and recycled by THIS program every
        'Recycle worker after documents' documents per reader. (The built-in max_tasks_per_child option is not used: it can deadlock in Python 3.12.)"""
        if self.pool is None:
            self.pool = ProcessPoolExecutor(max_workers=self.n_workers, mp_context=multiprocessing.get_context("spawn"),
                                            initializer=worker.init_worker, initargs=(self.snap, ocr.config(), paths.home()))
            self.pool_tasks = 0
        return self.pool

    def _kill_pool(self):
        """Get rid of the reader processes at once (after a problem)."""
        pool, self.pool = self.pool, None
        if pool is None:
            return
        procs = list(getattr(pool, "_processes", {}).values())
        try:
            pool.shutdown(wait=False, cancel_futures=True)
        except Exception:
            pass
        for p in procs:
            try:
                p.terminate()
            except Exception:
                pass

    def _recycle_pool_if_due(self):
        """Fresh reader processes after N documents each: memory is given back to Windows."""
        every = settings.run_int("Recycle worker after documents", 50)
        if self.pool is not None and every > 0 and self.pool_tasks >= every * self.n_workers:
            pool, self.pool = self.pool, None
            pool.shutdown(wait=True)

    def _teardown(self):
        if self.pool is not None:
            self.pool.shutdown(wait=False, cancel_futures=True)
        if self.conn is not None:
            self.conn.commit()
            self.conn.close()

    # ------------------------------------------------------------------------ the run
    def run(self):
        try:
            self._setup()
        except (ConfigError, SettingsError) as e:
            self.log("")
            self.log("CANNOT START: " + str(e), stamp=False)
            self.summary.stop_reason = "setup incomplete"
            return self.summary
        lock = None if self.shadow else RunLock(paths.p("data", "run.lock"), self.conf["max_minutes"] + 15)
        try:
            if lock:
                lock.__enter__()
                if lock.took_over_from is not None:
                    self.summary.warnings.append(f"the previous run did not finish (its process {lock.took_over_from or '?'} is gone); this run carries on from what the database holds")
            self._body()
        except LockBusy as e:
            self.log(f"Not started: {e}")
            self.summary.stop_reason = "another run in progress"
            return self.summary
        finally:
            if lock and lock.held:
                lock.__exit__(None, None, None)
            self._teardown()
        self.summary.peak_mb = memory.peak_mb()
        text = self.summary.text()
        self.log("")
        for line in text.splitlines():
            self.log(line, stamp=False)
        if self.shadow:
            path = exports.write_shadow_report(self.report_rows, text, dict(
                blocked=self.summary.blocked_words, new_suppliers=self.summary.new_suppliers, skipped=self.skipped_sample,
                counters=dict(self.summary.c), mailboxes=self.summary.mailbox_data, notes=self.summary.notes + self.summary.warnings))
            self.log(f"\nShadow report: {path}", stamp=False)
        if getattr(self.log, "path", None):
            self.log(f"Log file: {self.log.path}", stamp=False)
        return self.summary

    def _body(self):
        mode = "SHADOW (nothing is changed)" if self.shadow else "LIVE"
        self.log(f"Run started - {mode} - destination mode: {self.router.mode} - root: {self.router.root or '(not set)'}")
        if self.subject_filter:
            note = (f"TEST FILTER: only e-mails whose subject contains '{self.subject_filter}' are considered; "
                    "all other e-mails are left alone and the catch-up window is NOT advanced.")
            self.log(note)
            self.summary.notes.append(note)
        if not self.shadow:
            self.router.cleanup_scratch()
            self._apply_fixes()
            self._apply_site_choices()
            self._recheck_waiting()
            self._month_housekeeping()
        rows = companies.reading_list()                             # the companies + the Fleet card / Statements mailboxes (never office@, never On hold)
        off = [r["Company name"] for r in settings.rows("Companies")
               if not companies._active(r) and r["Text to look for"] != "zz test only"]
        if off:                                                   # switched off in Setup step 2: say so in every run
            self.summary.warnings.append(f"{len(off)} compan{'y is' if len(off) == 1 else 'ies are'} switched OFF and NOT read: "
                                         + ", ".join(off) + " (tick them again in Setup step 2)")
        held = [sp["label"] for sp in companies.special_rows() if sp["role"] == "on hold"]
        if held:
            self.summary.mailbox_lines.append("  " + ", ".join(held) + ": ON HOLD - not read (SpecialMailboxes sheet)")
        if self.only:
            want = settings.norm(self.only)
            rows = [r for r in rows if any(want in settings.norm(n) for n in companies.mailbox_names(r)) or want in settings.norm(r["Company name"])]
        for row in rows:
            if self._should_stop():
                break
            self._mailbox(row)
        if not self.shadow:
            self._finish_live()

    # ------------------------------------------------------------------------ stop conditions
    def _should_stop(self):
        if time.time() > self.deadline:
            self.summary.stop_reason = self.summary.stop_reason or f"time limit of {self.conf['max_minutes']} minutes reached"
            return True
        free = memory.free_ram_mb()
        if free is not None and free < self.conf["stop_ram"]:
            self.summary.stop_reason = f"free memory is only {free} MB"
            return True
        return False

    # ------------------------------------------------------------------------ one mailbox
    def _mailbox(self, row):
        label, names = row["Company name"], companies.mailbox_names(row)
        if not names:
            self.summary.mailbox_lines.append(f"  {label}: no mailbox name in the Companies sheet - skipped")
            self.log(f"{label}: no mailbox name in the Companies sheet - skipped")
            return
        never = companies.never_names()
        try:
            store = self.gw.resolve_store(names, exclude=never)    # a mailbox set to Never read is never reached by another name
        except LookupError:
            try:
                other = self.gw.resolve_store(names)
            except LookupError:
                other = None
            if other is not None:
                self.summary.mailbox_lines.append(f"  {label}: its Outlook mailbox {other} is set to Never read - not read")
                self.log(f"{label}: its Outlook mailbox {other} is set to Never read - not read")
                return
            self.summary.mailbox_lines.append(f"  {label}: mailbox not found in Outlook ({names[0]}) - skipped")
            self.log(f"{label}: mailbox not found in Outlook ({names[0]})")
            return
        last = None if self.shadow else db.last_success(self.conn, label)
        start, end, kind = compute_window(self.now, self.from_date, self.to_date, last, self.conf["default_days"],
                                          self.conf["max_days"], self.conf["overlap"])
        try:
            fuel_box = self.roles.get(label, ("", ""))[0] in ("fuel invoices", "read all")     # Inbox + its sub-folders
            infos, truncated = (self.gw.scan(store, label, start, end, self.conf["max_emails"], subfolders=True) if fuel_box
                                else self.gw.scan(store, label, start, end, self.conf["max_emails"]))
        except Exception as e:
            self.summary.mailbox_lines.append(f"  {label}: could not read the mailbox ({e})")
            self.log(f"{label}: could not read the mailbox: {e}")
            return
        if self.subject_filter:
            infos = [i for i in infos if self.subject_filter in i.subject.lower()]
        self.summary.c["emails scanned"] += len(infos)
        self.summary.c["mailboxes read"] += 1
        if not infos:                                              # nothing new: no noise in the summary
            self.summary.mailbox_data.append((label, 0, f"{start:%Y-%m-%d %H:%M}", f"{end:%Y-%m-%d %H:%M}", kind, 0, 0, 0, ""))
            self.summary.c["quiet mailboxes"] += 1
            if not self.shadow and not self.subject_filter:
                db.record_mailbox_run(self.conn, label, start, end, not truncated)
                self.conn.commit()
            return
        wanted, n_blocked, n_skipped = [], 0, 0
        for n, info in enumerate(infos):
            if n % 50 == 0 and self._should_stop():
                break
            key = db.email_key(label, info.entry_id)
            if key in self.seen:
                outcome, attempts = self.seen[key]
                if not (outcome == "error" and attempts < MAX_EMAIL_ATTEMPTS):
                    self.summary.c["already handled"] += 1
                    n_skipped += 1
                    continue
            if info.done and self.conf["skip_ticked"]:               # ticked before (by the bot, an older copy of it, or a person): never read again
                self.summary.c["already ticked"] += 1
                n_skipped += 1
                continue
            sender = doctypes.sender_row(info.sender_name, info.sender_email)
            fwre = bool(FWRE.match(info.subject))
            word = doctypes.blocked_word(info.subject)
            if word:
                self._blocked(info, word, fwre)
                n_blocked += 1
            elif (doctypes.subject_matches(info.subject) or doctypes.always_download(sender) or (fwre and self.conf["fwre_download"])
                  or fuel_box):                                  # the fuel mailbox holds only fuel invoices: every e-mail with a PDF is read
                wanted.append(info)
            elif self._order_name_probe(info):                    # no keyword in the subject, but a file called ...order...: look at the page
                self.probe_ids.add(info.entry_id)
                wanted.append(info)
            elif fwre and self.conf["red_all"]:
                self._category_only(info)
            elif self.shadow and len(self.skipped_sample) < 200:
                self._sample_skipped(info)
        done_before = self.summary.c["documents"]
        self.log(f"{label}: {len(infos)} e-mails found, {len(wanted)} with documents to read" + (f", {n_blocked} blocked" if n_blocked else ""))
        for i in range(0, len(wanted), self.conf["batch"]):
            if self._should_stop():
                break
            self.log(f"{label}: reading e-mails {i + 1} to {min(i + self.conf['batch'], len(wanted))} of {len(wanted)} ({self.summary.c['documents'] - done_before} documents done)")
            self._chunk(wanted[i:i + self.conf["batch"]])
        complete = not truncated and not self.summary.stop_reason and not self.retry_pending[label]
        if not self.shadow and not self.subject_filter:           # a filtered test must never move the real catch-up window
            db.record_mailbox_run(self.conn, label, start, end, complete)
            self.conn.commit()
        self.summary.mailbox_data.append((label, len(infos), f"{start:%Y-%m-%d %H:%M}", f"{end:%Y-%m-%d %H:%M}", kind, n_blocked,
                                          len(wanted), self.summary.c["documents"] - done_before, "" if complete else "window kept open"))
        self.summary.mailbox_lines.append(
            f"  {label}: {len(infos)} e-mails ({start:%Y-%m-%d %H:%M} to {end:%Y-%m-%d %H:%M}, {kind})  "
            f"blocked {n_blocked}, already done {n_skipped}, document e-mails {len(wanted)}, documents {self.summary.c['documents'] - done_before}"
            + ("  [window kept open - repeated next run]" if not complete else ""))
        self.log(self.summary.mailbox_lines[-1].strip())

    # ------------------------------------------------------------------------ Outlook actions
    def _outlook(self, what, info, description, *args):
        self.would[info.entry_id].append(description)
        if self.shadow:
            return True
        try:
            getattr(self.gw, what)(info, *args)
            return True
        except Exception as e:
            self.summary.c["outlook errors"] += 1
            self.summary.first_outlook_error = self.summary.first_outlook_error or f"{what}: {str(e)[:400]}"
            if self.summary.c["outlook errors"] <= 3:
                self.log(f"could not update Outlook ({what}) for '{info.subject[:50]}': {e}")
            return False

    def _record_email(self, info, outcome, detail=""):
        if not self.shadow:
            self.changed = True
            db.record_email(self.conn, info.mailbox, info.entry_id, info.subject, info.received, outcome, detail)
            self.conn.commit()

    def _red_category(self, info):
        self._outlook("flag_red", info, "red flag")

    def _blocked(self, info, word, fwre):
        """Overdue / urgent / reminder ...: never download. Red flag (and red category for FW/RE)."""
        self.summary.c["blocked"] += 1
        self.summary.blocked_words[word] += 1
        if self.conf["flag_blocked_no_att"] or self._has_documents(info):
            self._outlook("flag_red", info, "red flag")
        if fwre:
            self._red_category(info)
        self._record_email(info, "blocked", word)
        self._report_email(info, f"NOT downloaded (blocked word '{word}')")

    def _category_only(self, info):
        self.summary.c["category-only"] += 1
        self._red_category(info)
        self._record_email(info, "category-only", "FW/RE")
        self._report_email(info, "red flag only")

    def _has_documents(self, info):
        try:
            return any(a.kind == 1 and os.path.splitext(a.name)[1].lower() in DOC_EXT for a in self.gw.attachments(info))
        except Exception:
            return True

    def _sample_skipped(self, info):
        try:
            atts = [a.name for a in self.gw.attachments(info) if not a.hidden and a.kind == 1]
        except Exception:
            return
        if atts:
            self.skipped_sample.append([info.mailbox, str(info.received)[:16], info.subject, ", ".join(atts[:6])])

    # ------------------------------------------------------------------------ a batch of document e-mails
    def _chunk(self, infos):
        jobs, state = [], {}
        scratch = paths.ensure("Scratch")
        for info in infos:
            if self.conf["scan_body"]:
                try:
                    word = doctypes.blocked_word(self.gw.body_start(info))
                except Exception:
                    word = None
                if word:
                    self._blocked(info, f"{word} (in the e-mail body)", bool(FWRE.match(info.subject)))
                    continue
            key = db.email_key(info.mailbox, info.entry_id)
            try:
                atts = self.gw.attachments(info)
            except Exception as e:
                self._email_error(info, f"cannot list attachments: {e}")
                continue
            files = [a for a in atts if a.kind == 1 and not a.hidden and os.path.splitext(a.name)[1].lower() in DOC_EXT]
            images = [a for a in atts if a.kind == 1 and not a.hidden and os.path.splitext(a.name)[1].lower() in IMG_EXT
                      and a.size >= self.conf["image_kb"] * 1024]
            if not files:
                if images:
                    self.summary.c["image-only"] += 1
                    self._outlook("add_category", info, f"orange category '{self.conf['img_cat']}'", self.conf["img_cat"], "orange")
                    if FWRE.match(info.subject):
                        self._red_category(info)
                    self._record_email(info, "image-only", ", ".join(a.name for a in images[:4]))
                    self._report_email(info, "image-only: orange category, not read")
                else:
                    self.summary.c["no-attachment"] += 1
                    if FWRE.match(info.subject):
                        self._red_category(info)
                    self._record_email(info, "no-attachment")
                    self._report_email(info, "no attachment: left untouched")
                continue
            state[key] = {"info": info, "handled": 0, "failed": 0}
            tag = naming.clean(info.entry_id[-8:], keep_spaces=False) or "mail"
            for a in files[: self.conf["max_atts"]]:
                path = os.path.join(scratch, f"{tag}_{a.index}_{naming.clean(a.name, keep_spaces=False) or 'file'}")
                try:
                    self.gw.save_attachment(info, a.index, path)
                except Exception as e:
                    state[key]["failed"] += 1
                    self.log(f"could not save '{a.name}': {e}")
                    continue
                jobs.append({"job_id": len(jobs), "path": path, "info": info, "att": a,
                             "oversized": a.size > self.conf["max_mb"] * 1048576,
                             "ctx": {"subject": info.subject, "sender_name": info.sender_name,
                                     "sender_email": info.sender_email, "mailbox": info.mailbox,
                                     "is_forward": bool(FWRE.match(info.subject)), "orig_name": a.name,
                                     "to_name": info.to_name, "extra_known": sorted(self.known)[:300]}})
        results = self._analyse(jobs)
        pairs = sorted(zip(jobs, results), key=lambda p: p[0]["path"].lower().endswith((".xlsx", ".xlsm")))     # PDFs first: a spreadsheet is filed beside its invoice
        for job, res in pairs:
            st = state[db.email_key(job["info"].mailbox, job["info"].entry_id)]
            try:
                self._handle(job, res, st)
            except Exception as e:                                # keep going; the e-mail is retried next run
                st["failed"] += 1
                self.log(f"problem with '{job['att'].name}': {type(e).__name__}: {e}")
        self.conn.commit()
        for st in state.values():
            self._finish_email(st)

    def _analyse(self, jobs):
        """Read the files - in parallel when there are several and enough memory."""
        results = self._analyse_raw(jobs)
        peaks = [r.get("peak_mb") for r in results if r.get("peak_mb")]
        if peaks:
            self.summary.worker_peak_mb = max(peaks + [self.summary.worker_peak_mb or 0])
        return results

    def _analyse_raw(self, jobs):
        if not jobs:
            return []
        light = [{"job_id": j["job_id"], "path": j["path"], "ctx": j["ctx"]} for j in jobs]
        free = memory.free_ram_mb()
        parallel = self.n_workers > 1 and len(jobs) > 1 and not (free is not None and free < self.conf["low_ram"])
        results = [None] * len(light)
        if parallel:
            try:
                futures = [self._pool().submit(worker.process_job, j) for j in light]
                for i, fut in enumerate(futures):
                    results[i] = fut.result(timeout=self.parallel_timeout)         # a stuck reader can never freeze the run
                self.pool_tasks += len(light)
                self._recycle_pool_if_due()
                return results
            except Exception as e:                                                  # includes a timeout and a crashed reader process
                self.log(f"parallel reading problem ({type(e).__name__}); the rest of this batch is read one by one")
                self.summary.warnings.append(f"parallel reading had a problem ({type(e).__name__}) and was restarted; nothing was lost")
                self._kill_pool()
        for i, j in enumerate(light):                                               # anything not read yet is read one at a time
            if results[i] is None:
                results[i] = self._read_one(j)
        return results

    def _read_one(self, j):
        """Read one file with a time limit, so a damaged file can never freeze the run. (With an explicit worker count - tests,
        command line - the file is read right here.)"""
        if self.workers_override is not None:
            return worker.process_job(j)
        t = time.time()
        try:
            return self._pool().submit(worker.process_job, j).result(timeout=self.parallel_timeout)
        except Exception as e:                                                      # too slow, or the reader process died
            self._kill_pool()
            res = worker.extractor._blank()
            res["fatal"] = True
            res["problems"].append(f"the file took too long or crashed the reader ({type(e).__name__}); skipped: a person should look at it")
            res.update(job_id=j["job_id"], seconds=round(time.time() - t, 2), peak_mb=0)
            self.log(f"gave up on a file after {time.time() - t:.0f} s: {os.path.basename(j['path'])}")
            self.summary.warnings.append("a damaged or very slow file was skipped (it is listed as unrecognised)")
            return res

    # ------------------------------------------------------------------------ deciding and writing
    @staticmethod
    def _fields_from_result(job, res):
        info = job["info"]
        date = res.get("invoice_date")
        return dict(doc_type=res["doc_type"], supplier=res["supplier"], short=res["short"], folder_override=res["folder_override"],
                    supplier_source=res["supplier_source"], number=res["invoice_number"], date=date,
                    orig_name=job["att"].name, ext=os.path.splitext(job["att"].name)[1].lower(), received=info.received,
                    mailbox=info.mailbox, doc_company=res["doc_company"], site_company=res["site_company"],
                    sender_company=(companies.for_sender(info.sender_name, info.sender_email) or {}).get("Company name", ""),
                    site=res["site"], site_source=res["site_source"],
                    company_folder="", company_source="", fatal=res["fatal"], problems=list(res["problems"]),
                    oversized=job["oversized"])

    @staticmethod
    def _fields_from_row(d):
        record, short, override = settings.canonical_supplier(d["supplier"] or "")
        received = datetime.datetime.fromisoformat(str(d["received"])) if d["received"] else datetime.datetime.now()
        return dict(doc_type=d["doc_type"], supplier=d["supplier"] or "", short=short, folder_override=override,
                    supplier_source="stored", number=d["number"], date=parse_date(d["doc_date"]),
                    orig_name=d["orig_name"] or "", ext=os.path.splitext(d["orig_name"] or "")[1].lower() or ".pdf",
                    received=received, mailbox=d["mailbox"], doc_company="",
                    site_company=((sites.site_row(d["site"]) or {}).get("company", "") if d["site"] else ""),
                    site=d["site"] or settings.UNASSIGNED, site_source=d["site_source"] or "",
                    company_folder=d["company_folder"] or "", company_source=d["company_source"] or "",
                    fatal=False, problems=[], oversized=False)

    def _decide_and_place(self, f, src):
        """The routing rules. Returns {'status','reason','dest_dir','dest_name','company_folder','company_source','warnings'}."""
        rt = self.router
        out = dict(status="", reason="", dest_dir="", dest_name="", company_folder=f["company_folder"],
                   company_source=f["company_source"], warnings=[], remarks=[])
        rule = doctypes.rule(f["doc_type"]) if f["doc_type"] else None
        ext = f["ext"] or ".pdf"
        shown_date = naming.format_date(f["date"] or f["received"].date(), self.conf["name_date"])
        raw_name = (f"{naming.format_date(f['received'], self.conf['name_date'])}_"
                    f"{naming.clean(os.path.splitext(f['orig_name'])[0], keep_spaces=False) or 'file'}{ext}")

        # an unreadable or failed document is still named "Supplier_Number" when the supplier and the number were read
        held_name = raw_name
        if rule and rule["route"] == "company" and f["supplier"] and f["number"] and not f["fatal"] and not f["oversized"]:
            held_name = naming.build(rule["pattern"], f["supplier"], f["short"], f["number"], shown_date, rule["tag"], rule["type"], f["orig_name"], ext)

        month = self._holding_month(f)

        def put(parts, name, status, reason=""):
            path, _how = rt.hold(src, parts, name)
            out.update(status=status, reason=reason, dest_dir=os.path.dirname(path), dest_name=os.path.basename(path))
            return out

        if f["oversized"]:
            return put((UNRECOGNISED_DIR, month), raw_name, db.UNRECOGNISED, f"file is larger than {self.conf['max_mb']} MB; not read")
        if f["fatal"]:
            return put((UNRECOGNISED_DIR, month), raw_name, db.UNRECOGNISED, f["problems"][0] if f["problems"] else "unreadable")
        if rule is None:
            return put((UNRECOGNISED_DIR, month), raw_name, db.UNRECOGNISED,
                       "; ".join(f["problems"][:2]) or "document type not recognised")
        site_known = f.get("site") if f.get("site") and f.get("site") != settings.UNASSIGNED else ""
        if rule["route"] == "mailbox" or self._special_statement(f):  # Fleet card / Statements mailbox: <its folder>\<SITE>\<date>_STM
            return self._place_special(f, src, rule, site_known, shown_date, ext, out, put)
        name = naming.build(rule["pattern"], f["supplier"], f["short"], f["number"] or "", shown_date, rule["tag"],
                            rule["type"], f["orig_name"], ext, site=site_known)
        if "{site}" in rule["pattern"] and not site_known:         # statement / order confirmation without a site: flagged in To_Fix
            out["remarks"].append(settings.MISSING_SITE)
        if rule["route"] == "holding":                             # its own holding folder, e.g. _Statements\October 2026\<COMPANY>
            parts = [rule["subfolder"] or "_Holding"]
            for token in doctypes.layout_parts(rule["layout"]):
                if token == "{MONTH}":
                    parts.append(month)
                elif token == "{COMPANY}":
                    folder = self._holding_company(f, out)
                    if not folder:                                 # filed one level up, and said so
                        out["remarks"].append("COMPANY NOT FOUND")
                        break
                    parts.append(folder)
                elif token == "{SITE}":
                    if not site_known:
                        if settings.MISSING_SITE not in out["remarks"]:
                            out["remarks"].append(settings.MISSING_SITE)
                        break
                    parts.append(sites.folder_for(site_known))
                else:
                    parts.append(token)
            if "{COMPANY}" not in rule["layout"] and site_known and not out["company_folder"]:
                self._holding_company(f, out)                      # for the record only
            return put(tuple(parts), name, db.PARKED)
        if rule["route"] == "unrecognised":                        # parked on purpose: account summaries, activity reports ...
            return put((UNRECOGNISED_DIR, month, rule["subfolder"]), name, db.PARKED)
        if f["problems"]:
            return put((UNRECOGNISED_DIR, month), held_name, db.UNRECOGNISED, "; ".join(f["problems"][:3]))
        if not f["supplier"]:
            return put((UNRECOGNISED_DIR, month), raw_name, db.UNRECOGNISED, "supplier not recognised")
        if (self.conf["hold_first"] and f["supplier_source"] in ("printed", "logo", "sender name")
                and settings.supplier_key(f["supplier"]) not in self.known):
            return put((UNASSIGNED_DIR,), name, db.UNASSIGNED, f"first time this supplier name is seen: '{f['supplier']}' - confirm it")
        # --- a ship-to address that is not a known site: the document WAITS (the bill-to company and site are never used)
        if (f.get("site") or settings.UNASSIGNED) == settings.UNASSIGNED and sites.ship_hold(f.get("site_source")) \
                and f.get("company_source") != "manual" and not f["folder_override"]:
            return put((UNASSIGNED_DIR,), name, db.UNASSIGNED,
                       f"{f['site_source']} - add the site to the Sites sheet (the next run files it by itself) or choose the Site folder in To_Fix.xlsx")
        # --- which company, which site: the layout (setting 'Destination pattern') says what is needed
        needs = rt.needs()
        folder = f["folder_override"] or f["company_folder"]
        if not folder:
            if f.get("company_row"):                                    # already chosen in step 1 (company first, then site)
                row, source, warns = f["company_row"], f["company_pick"], f.get("company_notes", [])
            else:
                row, source, warns = companies.resolve(f["doc_company"], f["site_company"], f["mailbox"],
                                                       sender_company=f.get("sender_company", ""))
            out["warnings"] += warns
            if row:
                folder, out["company_folder"], out["company_source"] = row["Folder name"], row["Folder name"], source
        else:
            out["company_folder"] = out["company_folder"] or folder
            out["company_source"] = out["company_source"] or ("supplier folder" if f["folder_override"] else "manual")
        if "company" in needs and not folder:
            return put((UNASSIGNED_DIR,), name, db.UNASSIGNED, "company folder could not be determined (mailbox not in the Companies sheet)")
        site_folder, skip_site = "", False
        if "site" in needs:
            site = f.get("site") or settings.UNASSIGNED
            if site == settings.UNASSIGNED:
                why = f.get("site_source") or "not found"
                if self.conf["site_fallback"] and "company" in needs and rt.site_is_last():
                    skip_site = True                                  # filed one level up, in the company folder, and said so
                    out["remarks"].append("SITE NOT FOUND")
                    out["warnings"].append(f"the site could not be determined ({why}); filed in the company folder")
                else:
                    return put((UNASSIGNED_DIR,), name, db.UNASSIGNED, f"site could not be determined ({why}) - choose the site folder in To_Fix.xlsx")
            else:
                site_folder = sites.folder_for(site)
        when, month_case, month_target, month_note = self._folder_month(f)
        if month_note:
            out["remarks"].append(month_note)
        out["month_case"], out["month_target"] = month_case, month_target
        try:
            target = rt.prepare_dir(when, company=folder, site=site_folder, skip_site=skip_site)
            path, how = rt.place(src, target, name)
        except PlacementError as e:
            put((UNASSIGNED_DIR,), name, db.UNASSIGNED, e.reason)
            out["attempts"] = 1
            return out
        out.update(status=db.FILED, dest_dir=os.path.dirname(path), dest_name=os.path.basename(path))
        if how == "renamed":
            out["warnings"].append("a different file with this name already existed; saved with a suffix")
        return out

    def _holding_company(self, f, out):
        """The company folder for a holding document (statements by company). Same rules as invoices: the document, the site's
        company, the sender, the mailbox (a special mailbox is never a clue). Returns the folder name or ''."""
        if f.get("company_folder"):
            out["company_folder"], out["company_source"] = f["company_folder"], f.get("company_source") or "stored"
            return f["company_folder"]
        if f.get("company_row"):
            row, source = f["company_row"], f.get("company_pick", "")
        else:
            row, source, warns = companies.resolve(f.get("doc_company", ""), f.get("site_company", ""), f["mailbox"],
                                                   sender_company=f.get("sender_company", ""))
            out["warnings"] += warns
        if not row:
            return ""
        out["company_folder"], out["company_source"] = row["Folder name"], source
        return row["Folder name"]

    def _place_special(self, f, src, rule, site_known, shown_date, ext, out, put):
        """Fleet Card statements (Fleet card mailbox) and statements of the Statements mailbox: <Destination folder>\\<SITE>\\<date>_STM.
        The site folder is created only if it is missing. No site = the document waits in _Unassigned (never a guessed folder)."""
        role, dest = self.roles.get(f["mailbox"], ("", ""))
        dest = f.get("special_dest") or dest or self._role_dest(   # 5.19.0: a fuel invoice / fleet statement from another mailbox
            {"Fleet Card Statement": "fleet card", "Fuel Invoice": "fuel invoices"}.get(rule["type"], "-"))   # (also when retried later)
        if rule["type"] == "Fuel Invoice":                       # <Destination folder>\<COMPANY>\<INVOICE NUMBER>_<SITE>.pdf
            name = naming.build("{number}_{site}", f["supplier"], f["short"], f["number"] or "", shown_date, "", rule["type"], f["orig_name"],
                                ext, site=site_known)
            if not f.get("company_folder") or not site_known:
                return put((UNASSIGNED_DIR,), name, db.UNASSIGNED,
                           "company or site not found in the Customer of this fuel invoice - choose the Company folder in To_Fix.xlsx")
            out["company_folder"], out["company_source"] = f["company_folder"], f.get("company_source", "")
            try:
                target = self.router.special_dir(dest, f["company_folder"])
                path, how = self.router.place(src, target, name)
            except PlacementError as e:
                put((UNASSIGNED_DIR,), name, db.UNASSIGNED, e.reason)
                out["attempts"] = 1
                return out
            out.update(status=db.FILED, dest_dir=os.path.dirname(path), dest_name=os.path.basename(path))
            return out
        name = naming.build("{date}_STM", f["supplier"], f["short"], f["number"] or "", shown_date, "STM", rule["type"], f["orig_name"], ext,
                            site=site_known)
        if not site_known:
            return put((UNASSIGNED_DIR,), name, db.UNASSIGNED,
                       f"site not found for this {rule['type'].lower()} - choose the Site folder in To_Fix.xlsx")
        if site_known and not out["company_folder"]:
            self._holding_company(f, out)                          # for the record only
        try:
            target = self.router.special_dir(dest, sites.folder_for(site_known))
            path, how = self.router.place(src, target, name)
        except PlacementError as e:
            put((UNASSIGNED_DIR,), name, db.UNASSIGNED, e.reason)
            out["attempts"] = 1
            return out
        out.update(status=db.FILED, dest_dir=os.path.dirname(path), dest_name=os.path.basename(path))
        if how == "renamed":
            out["warnings"].append("a different file with this name already existed; saved with a suffix")
        return out

    def _holding_month(self, f):
        """Month sub-folder for statements, order confirmations and _Unrecognised: 'October 2026' (see router.holding_month)."""
        from .router import holding_month
        return holding_month(f["received"], f["date"], self.conf["month_window"])

    def _folder_month(self, f):
        """Which month folder a document goes to, and why.  Returns (date for the folder, case, target month 'YYYY-MM', note for Remarks).
        case: ''  = the month it arrived in (or the invoice is dated that month)
              past   = an invoice dated 1..N months before the arrival month -> its own, older month folder
              far    = dated more than N months back or ahead -> the current (arrival) month folder
              future = dated 1..N months AFTER the arrival month -> waits in the current month folder; moved when its month starts"""
        arrival, inv, mode = f["received"], f["date"], self.conf["month_follow"]
        if mode == "invoice":
            return (inv or arrival), "", "", ""
        if mode != "window" or inv is None:
            return arrival, "", "", ""
        limit = self.conf["month_window"]
        diff = (arrival.year * 12 + arrival.month) - (inv.year * 12 + inv.month)
        cur = f"{router_months(arrival)} {arrival.year}"
        own = f"{router_months(inv)} {inv.year}"
        shown = naming.format_date(inv, self.conf["name_date"])
        if diff == 0:
            return arrival, "", "", ""
        if 1 <= diff <= limit:
            return inv, "past", "", f"INVOICE DATED {own}: filed in the {own} folder"
        if diff > limit:
            return arrival, "far", "", f"INVOICE DATE {shown} IS MORE THAN {limit} MONTHS OLD: filed in the {cur} folder"
        if diff < -limit:
            return arrival, "far", "", f"INVOICE DATE {shown} IS MORE THAN {limit} MONTHS AHEAD: filed in the {cur} folder"
        return arrival, "future", f"{inv.year:04d}-{inv.month:02d}", \
            f"FUTURE-DATED {own}: kept in the {cur} folder until the {own} folder exists, then moved there"

    def _select_company_and_site(self, f, res, job):
        """Step 1: the COMPANY (named on the document, else the sender, else the mailbox).
        Step 2: the SITE, looked for among the sites of that company only. A company named on the document is trusted: if none of its
        sites fits, the document waits for a person. A company that only comes from the sender / mailbox is a weaker clue: if none of its
        sites fits, the other companies' sites are tried too (and the difference is noted)."""
        if not self.conf["site_in_company"]:
            return
        info = job["info"]
        order = [o.strip() for o in settings.run_value("Company source order", "document,site,sender,mailbox").split(",")
                 if o.strip() and o.strip() != "site"]                     # the site comes AFTER the company here
        row, source, notes = companies.resolve(f["doc_company"], "", f["mailbox"], order=order,
                                               sender_company=f.get("sender_company", ""))
        notes = list(notes)

        def look(only):
            return sites.match(res.get("site_text", ""), subject=info.subject, blocks=res.get("site_blocks", []),
                               supplier=f["supplier"], to_name=info.to_name, only_company=only)

        first = look(None)                                   # a ship-to / deliver-to address (or a configured account number) decides BOTH the site and the company
        if sites.ship_hold(first["source"]):                 # a ship-to address is on the page but is not a known site: never fall back to the bill-to
            f["site"], f["site_source"], f["site_company"] = settings.UNASSIGNED, first["source"], ""
            ship = self._ship_text(res)
            self._resolve_unknown_ship_to(f, ship)
            return
        if first.get("ship_company"):                        # the ship-to names the company but none of its sites: that company, no site
            owner = companies.by_name(first["company"])
            if row and owner and owner["Company name"] != row["Company name"]:
                notes.append(f"the {source} points to {row['Folder name']}, but the ship-to address names {owner['Folder name']}: filed there")
            f["site"], f["site_source"], f["site_company"] = settings.UNASSIGNED, first["source"], first["company"]
            f["company_row"], f["company_pick"], f["company_notes"] = owner, "ship-to address", notes
            return
        if first["site"] != settings.UNASSIGNED and (first["source"].startswith((sites.SHIP_SOURCE, sites.SOLD_SOURCE))
                                                     or first["source"] == "account number"):
            owner = companies.by_name(first["company"])
            if owner:
                if row and owner["Company name"] != row["Company name"]:
                    notes.append(f"the {source} points to {row['Folder name']}, but the site {first['site']} belongs to {owner['Folder name']} ({first['source']}): filed there")
                f["site"], f["site_source"], f["site_company"] = first["site"], first["source"], first.get("company", "")
                f["company_row"], f["company_pick"], f["company_notes"] = owner, "site", notes
                return
        if row:
            found = look(row["Company name"])
            if found["site"] != settings.UNASSIGNED:
                found["source"] += f" (within {row['Folder name']})"
            elif source == "document":
                found["source"] = f"no site of {row['Folder name']} fits the document ({found['source']})"
            else:
                other = look(None)
                if other["site"] != settings.UNASSIGNED:
                    owner = companies.by_name(other["company"])
                    notes.append(f"the {source} points to {row['Folder name']}, but the site {other['site']} belongs to {owner['Folder name'] if owner else other['company']}")
                    row, source, found = owner or row, "site", other
        else:
            found = look(None)
            row, source = companies.by_name(found["company"]), "site"
        f["site"], f["site_source"], f["site_company"] = found["site"], found["source"], found.get("company", "")
        if row:
            f["company_row"], f["company_pick"], f["company_notes"] = row, source, notes

    @staticmethod
    def _ship_text(res):
        """The deliver-to address as printed: the ship-to blocks, or (no ship-to on the page) the sold-to blocks."""
        blocks = res.get("site_blocks", [])
        ship = sites.split_blocks([b for b in blocks if str(b).startswith(sites.SHIP_TAG)])[0]
        return " || ".join(ship or sites.sold_blocks(blocks))

    def _resolve_unknown_ship_to(self, f, ship_text):
        """A ship-to address that is not in the Sites sheet: (1) a site staff chose before for this supplier's deliver-to number /
        address, (2) a site FOLDER that exists in the month folder and is written in the ship-to address. Never the bill-to.
        Returns True when the site (and its company) is now known."""
        if not ship_text:
            return False
        hit = None
        if settings.run_bool("Learn sites from staff fixes", True) and self.conn is not None:
            got = db.learned_site(self.conn, "ship-to", settings.supplier_key(f.get("supplier") or ""), sites.ship_keys(ship_text))
            if got:
                hit = (got[0], got[1], f"learned from a staff fix ({got[2][:10]})")
        if hit is None and settings.run_bool("Look for the site in the SharePoint folders", True) and self.router is not None \
                and {"company", "site"} <= self.router.needs():
            when = self._folder_month(f)[0] if f.get("received") else self.now
            pair = sites.folder_in_ship_text(ship_text, sites.site_folders_in(self.router.month_dir(when)))
            if pair:
                hit = (pair[1], pair[0], "ship-to address = SharePoint folder")
        if hit is None:
            return False
        site, company_folder, how = hit
        row = sites.site_row(site) or sites.find_site(site)
        site_name = row["name"] if row else site
        owner = companies.by_folder(company_folder) or (companies.by_name(row["company"]) if row and row.get("company") else None)
        if owner is None:
            return False
        f["site"], f["site_source"], f["site_company"] = site_name, how, owner["Company name"]
        f["company_row"], f["company_pick"], f["company_notes"] = owner, "site", []
        f["company_folder"], f["company_source"] = owner["Folder name"], "site"
        return True

    @staticmethod
    def _fleet_site_by_name(f, res):
        """A Fleet Card statement: exactly one site NAMED at the top ('MOBIL MOE', 'Pearl Energy Tomago') decides the site and its
        company. The address under it is the owner's postal address (Moe -> Kirk St, Tomago -> Hexham) and is not used."""
        top = "\n".join([l for l in str(res.get("site_text", "")).splitlines() if l.strip()][:12])
        hits = sites._name_hits(sites.active_sites(), settings.words(top))
        if len(hits) != 1:
            return False
        name = next(iter(hits))
        row = sites.site_row(name)
        owner = companies.by_name(row["company"]) if row else None
        if not owner:
            return False
        f["site"], f["site_source"], f["site_company"] = name, "merchant name on the Fleet Card statement", row["company"]
        f["company_row"], f["company_pick"], f["company_notes"] = owner, "site", []
        return True

    def _role_dest(self, role):
        """The Destination folder of the special mailbox with this role (SpecialMailboxes sheet), or '' when there is none."""
        for sp in companies.special_rows():
            if sp["role"] == role and sp["dest"]:
                return sp["dest"]
        return ""

    def _read_fuel_invoice(self, job, f, res, any_mailbox=False):
        """A PDF in the Fuel invoices mailbox: if it is a Pearl Investment fuel invoice, its Customer decides the company and
        the site (never the Reference); the invoice number names the file and finds duplicates.
        any_mailbox (5.19.0): the same, in any other mailbox, but only for a page that names Pearl Investment Family Trust."""
        from . import fuel
        got = fuel.parse_if_fuel(job["path"]) if any_mailbox else fuel.parse(job["path"])
        if not got:
            return
        rows = companies.find_in_text(got["company_text"], header_lines=3)
        f.update(doc_type="Fuel Invoice", supplier=got["supplier"], short="", supplier_source="fuel invoice", number=got["number"],
                 date=got["date"] or f["date"], site=got["site_text"], site_source="customer on the invoice",
                 reference=got["reference"], ship_override=got["customer"], fatal=False, problems=[])
        res.update(doc_type="Fuel Invoice", supplier=got["supplier"], invoice_number=got["number"], invoice_date=got["date"],
                   site=got["site_text"], site_source="customer on the invoice", problems=[], fatal=False)
        if rows:
            f["company_folder"], f["company_source"] = rows[0][0]["Folder name"], "customer on the invoice"
        else:                                                # not one of the companies (e.g. SKM DIESEL PTY LTD): its own folder, named as printed
            from .router import folder_clean
            f["company_folder"] = folder_clean(got["company_text"])
            f["company_source"] = "customer on the invoice (not a company in the Companies sheet - own fuel folder)"

    def _special_statement(self, f):
        """A statement in the Statements mailbox (mstatement@): goes to that mailbox's own folder by site."""
        return f.get("doc_type") == "Statement" and self.roles.get(f.get("mailbox"), ("", ""))[0] == "statements"

    def _order_name_probe(self, info):
        if not self.conf["order_probe"]:
            return False
        try:
            return any(a.kind == 1 and not a.hidden and re.search(r"order", a.name, re.IGNORECASE)
                       and os.path.splitext(a.name)[1].lower() in DOC_EXT for a in self.gw.attachments(info))
        except Exception:
            return False

    def _handle(self, job, res, st):
        info, att = job["info"], job["att"]
        f = self._fields_from_result(job, res)
        if info.entry_id in self.probe_ids and f["doc_type"] != "Order Confirmation":
            self.summary.c["not an order confirmation (ignored)"] += 1       # opened only because of its file name: nothing is filed or recorded as a document
            try:
                os.remove(job["path"])
            except OSError:
                pass
            self._record_email(info, "not-an-order", att.name)
            return
        role = self.roles.get(info.mailbox, ("", ""))[0]
        if role == "fuel invoices" and f["ext"] == ".pdf" and not f["oversized"]:
            self._read_fuel_invoice(job, f, res)
        elif f["ext"] == ".pdf" and not f["oversized"] and self._role_dest("fuel invoices"):
            self._read_fuel_invoice(job, f, res, any_mailbox=True)   # 5.19.0: a fuel invoice that came to another mailbox
            if f["doc_type"] == "Fuel Invoice":
                f["special_dest"] = self._role_dest("fuel invoices")
        if f["doc_type"] == "Fuel Invoice" and role != "fuel invoices" and not f.get("special_dest"):
            f["doc_type"] = res["doc_type"] = "Invoice"               # no Fuel invoices mailbox in the settings: an ordinary invoice
        if f["doc_type"] == "Fleet Card Statement" and role != "fleet card":
            if self._role_dest("fleet card"):                         # 5.19.0: filed with the other Fleet Card statements, whatever the mailbox
                f["special_dest"] = self._role_dest("fleet card")
            else:
                f["doc_type"] = res["doc_type"] = "Other" if doctypes.rule("Other") else f["doc_type"]
        rule0 = doctypes.rule(f["doc_type"]) if f["doc_type"] else None
        needs_place = f["doc_type"] != "Fuel Invoice" and rule0 and (rule0["route"] in ("company", "mailbox") or (rule0["route"] == "holding" and
                                 any(t in rule0["layout"] for t in ("{COMPANY}", "{SITE}"))) or self._special_statement(f))
        if needs_place and f["doc_type"] == "Fleet Card Statement" and not f["fatal"] and not f["problems"] \
                and self._fleet_site_by_name(f, res):
            needs_place = False                                    # 5.20.0: the merchant NAME decides (the address is the owner's postal address)
            res["site"], res["site_source"] = f["site"], f["site_source"]
        if needs_place and not f["fatal"] and not f["problems"]:
            self._select_company_and_site(f, res, job)
            res["site"], res["site_source"] = f["site"], f["site_source"]
        merchant = sites.merchant_number(res.get("site_text", "")) if f["doc_type"] == "Fleet Card Statement" else ""
        if merchant:                                             # a Fleet Card statement: its merchant number is remembered for its site
            skey, mkey = settings.supplier_key(f["supplier"] or ""), "merchant:" + merchant
            if (f.get("site") or settings.UNASSIGNED) == settings.UNASSIGNED:
                got = db.learned_site(self.conn, "ship-to", skey, [mkey]) if self.conn is not None else None
                if got and (sites.site_row(got[0]) or sites.find_site(got[0])):
                    f["site"], f["site_source"] = got[0], f"merchant number {merchant} (learned {got[2][:10]})"
                    res["site"], res["site_source"] = f["site"], f["site_source"]
            elif not self.shadow and settings.run_bool("Learn sites from staff fixes", True):
                db.learn(self.conn, "ship-to", skey, mkey, f["site"], example=f"merchant number {merchant}", by="bot (site name on the statement)")
            f["ship_override"] = mkey
        if f["doc_type"] == "Fleet Card Statement":            # the Fleet Card record shows the statement's TOTAL PAYMENT DUE
            m = re.search(r"TOTAL\s+PAYMENT\s+DUE\s*\$?\s*([\d,]+\.\d{2})", res.get("site_text", ""), re.IGNORECASE)
            if m:
                res["amount"] = float(m.group(1).replace(",", ""))
        # a spreadsheet that came with an invoice and is not a document of its own is filed BESIDE that invoice
        primary = st.get("primary")
        out = None
        if primary and f["ext"] in (".xlsx", ".xlsm") and self._is_breakdown(res):
            out = self._place_breakdown(job, f, res, primary)
        if out is None:
            # duplicates: the same content, or the same supplier + type + number, already handled
            if not f["fatal"]:
                rule = doctypes.rule(f["doc_type"]) if f["doc_type"] else None
                by_number = f["number"] if (rule and rule["route"] in ("company", "mailbox") and "number" in rule["required"]) else None
                dup = None if self.shadow else db.find_duplicate(self.conn, f["supplier"], f["doc_type"], by_number, res["file_hash"])
                if not dup and not self.shadow and f["doc_type"] == "Fleet Card Statement" and f.get("date") \
                        and str(f.get("ship_override") or "").startswith("merchant:"):   # the same statement from two mailboxes
                    dup = db.find_fleet_duplicate(self.conn, f["ship_override"], f["date"].isoformat())
                if dup:
                    self.summary.c["duplicates"] += 1
                    st["handled"] += 1
                    os.remove(job["path"])
                    if not self.shadow:
                        self._note_duplicate(dup, info)
                    self._report_doc(job, res, None, "duplicate of " + dup["ref"])
                    return
            out = self._decide_and_place(f, job["path"])
        is_new = (f["supplier_source"] in ("printed", "logo", "sender name") and f["supplier"]
                  and settings.supplier_key(f["supplier"]) not in self.known)
        if is_new:
            self.summary.new_suppliers.append((f["supplier"], f["supplier_source"]))
            self.known.add(settings.supplier_key(f["supplier"]))
        warnings = "; ".join(res["warnings"] + out["warnings"])
        row = dict(status=out["status"], route=out["status"], reason=out["reason"], supplier=f["supplier"],
                   supplier_source=f["supplier_source"], doc_type=f["doc_type"] or "", number=f["number"] or "",
                   doc_date=f["date"].isoformat() if f["date"] else "", amount=res["amount"], gst=res["gst"],
                   ex_amount=res.get("ex_amount"), remarks="; ".join(x for x in [res.get("remarks", "")] + out.get("remarks", []) if x),
                   company_folder=out["company_folder"], company_source=out["company_source"], site=res["site"],
                   site_source=res["site_source"], mailbox=info.mailbox, entry_id=info.entry_id, subject=info.subject,
                   sender=f"{info.sender_name} <{info.sender_email}>", received=info.received.isoformat(sep=" "),
                   orig_name=att.name, file_hash=res["file_hash"], dest_dir=out["dest_dir"], dest_name=out["dest_name"],
                   current_path=os.path.join(out["dest_dir"], out["dest_name"]), warnings=warnings, method=res["method"],
                   page_count=res["page_count"], ocr_used=1 if res["ocr_used"] else 0, attempts=out.get("attempts", 0),
                   filed_at=db.now() if out["status"] == db.FILED else None,
                   month_case=out.get("month_case", "") if out["status"] == db.FILED else "",
                   month_target=out.get("month_target", "") if out["status"] == db.FILED else "",
                   reference=f.get("reference", ""),
                   ship_to=f.get("ship_override") or self._ship_text(res)[:600])
        ref = db.insert_doc(self.conn, row)
        self.changed = True
        if not self.shadow:                                      # the last e-mail of each mailbox in this run: marked in the Record
            last = self.run_last.get(info.mailbox)
            if last is None or info.received >= last[0]:
                self.run_last[info.mailbox] = (info.received, ref)
        row["archive_path"] = ""
        if self.shadow:
            os.remove(job["path"])
        else:
            arch = self.router.archive(job["path"], ref, att.name, self.conf["archive_days"])
            db.update_doc(self.conn, ref, archive_path=arch)
        self.summary.c["documents"] += 1
        self.summary.c[out["status"]] += 1
        if out["status"] in (db.UNRECOGNISED, db.UNASSIGNED) and out["reason"]:
            for part in out["reason"].split("; "):
                self.summary.reasons[_reason_key(part)] += 1
        st["handled"] += 1
        if out["status"] == db.FILED:
            self.filed_now.append(dict(row, ref=ref))
            if f["ext"] == ".pdf" and f["doc_type"] != "Breakdown" and "primary" not in st:
                st["primary"] = dict(dir=out["dest_dir"], name=out["dest_name"], supplier=f["supplier"], number=f["number"],
                                     date=f["date"], site=res["site"], company_folder=out["company_folder"], ref=ref)
        self.touched_months.add(self._month_of(f))
        self._report_doc(job, res, out, out["status"] + (f": {out['reason']}" if out["reason"] else ""))

    @staticmethod
    def _is_breakdown(res):
        """A spreadsheet that is not an invoice in its own right: nothing could be read as a document, or no amounts were found."""
        return bool(res["fatal"] or res["problems"] or not res["doc_type"]
                    or (res["amount"] is None and res["gst"] is None and res["ex_amount"] is None))

    def _place_breakdown(self, job, f, res, primary):
        """File the spreadsheet next to its invoice, named like it ('<invoice name>_breakdown.xlsx')."""
        stem = os.path.splitext(primary["name"])[0]
        try:
            path, _how = self.router.place(job["path"], primary["dir"], f"{stem}_breakdown{f['ext']}")
        except PlacementError:
            return None                                          # could not be placed there: it is handled like any other document
        f.update(doc_type="Breakdown", supplier=primary["supplier"], number=primary["number"], date=primary["date"])
        res.update(doc_type="Breakdown", supplier=primary["supplier"], invoice_number=primary["number"], amount=None, gst=None,
                   ex_amount=None, site=primary["site"], site_source="with its invoice", problems=[], warnings=[],
                   remarks=f"breakdown of {primary['name']}")
        return dict(status=db.FILED, reason="", dest_dir=os.path.dirname(path), dest_name=os.path.basename(path),
                    company_folder=primary["company_folder"], company_source="with its invoice", warnings=[], remarks=[])

    def _note_duplicate(self, dup, info):
        """The same invoice arrived again (a reprint, a re-send): no second copy is filed; the original's Remarks says so."""
        try:
            d = db.get_doc(self.conn, dup["ref"]) or {}
            note = f"duplicate arrived {naming.format_date(info.received, self.conf['name_date'])}, not filed again"
            old = d.get("remarks") or ""
            if note not in old:
                db.update_doc(self.conn, dup["ref"], remarks="; ".join(x for x in (old, note) if x))
                self.changed = True
        except Exception as e:
            self.log(f"could not note the duplicate on {dup.get('ref')}: {e}")

    def _month_of(self, f):
        """The (year, month) of the record this document belongs to - same basis as the SharePoint month folder."""
        d = (f["date"] or f["received"].date()) if self.conf["month_basis"] == "invoice" else f["received"]
        return (d.year, d.month)

    def _finish_email(self, st):
        info = st["info"]
        if st["failed"] > 0:                                   # even one lost attachment: do not tick, retry the e-mail later
            self._email_error(info, f"{st['failed']} attachment(s) could not be saved or read")
        elif st["handled"] > 0:
            self.summary.c["emails downloaded"] += 1
            if FWRE.match(info.subject):                       # a forward / reply: a person looks at it - UNREAD, no tick, red category
                self.summary.c["forwards left unread"] += 1
                self._outlook("mark_unread", info, "leave UNREAD (no tick)")
                self._red_category(info)
            else:
                self._outlook("mark_done", info, "mark as read + flag complete (tick)")
            self._record_email(info, "downloaded", f"{st['handled']} file(s)")
            self._report_email(info, None)

    def _email_error(self, info, why):
        self.summary.c["email errors"] += 1
        key = db.email_key(info.mailbox, info.entry_id)
        attempts = self.seen.get(key, ("", 0))[1] + 1
        self.seen[key] = ("error", attempts)
        if attempts < MAX_EMAIL_ATTEMPTS:
            self.retry_pending[info.mailbox] += 1              # keeps this mailbox's window open so the e-mail is seen again
        self._record_email(info, "gave-up" if attempts >= MAX_EMAIL_ATTEMPTS else "error", why)
        self.log(f"e-mail '{info.subject[:50]}': {why} (attempt {attempts})")

    # ------------------------------------------------------------------------ shadow report rows
    def _report_email(self, info, headline):
        if self.shadow and headline is not None:
            self.report_rows.append([info.mailbox, str(info.received)[:16], info.subject,
                                     "; ".join(self.would.get(info.entry_id, [])) or "nothing", "", "", "", "", None,
                                     None, None, None, "", "", "", "", "", "", headline])

    def _report_doc(self, job, res, out, result):
        if not self.shadow:
            return
        info = job["info"]
        out = out or {}
        self.report_rows.append([info.mailbox, str(info.received)[:16], info.subject,
                                 "; ".join(self.would.get(info.entry_id, [])) + ("" if FWRE.match(info.subject) else "; mark as read + flag complete (tick)"),
                                 job["att"].name, res["doc_type"] or "", res["supplier"], res["invoice_number"] or "",
                                 res["invoice_date"], res["amount"], res["gst"], res.get("ex_amount"), out.get("company_folder", ""), res["site"],
                                 out.get("dest_dir", ""), out.get("dest_name", ""), (out.get("status") or "duplicate"),
                                 "; ".join(x for x in [res.get("remarks", "")] + out.get("remarks", []) if x),
                                 "; ".join(([out.get("reason", "")] if out.get("reason") else []) + res["problems"] + res["warnings"])])

    # ------------------------------------------------------------------------ fixes and waiting files (live)
    def _apply_fixes(self):
        for lid in exports.read_forget():                        # To_Fix > Learned sites: 'Forget (type YES)'
            db.forget(self.conn, lid)
            self.changed = True
            self.summary.c["learned sites forgotten"] += 1
        conflicts = exports.to_fix_conflicts()
        if conflicts:
            self.summary.warnings.append("OneDrive made a conflict copy of To_Fix.xlsx (" + ", ".join(conflicts[:3]) + "): fixes typed in a copy are NOT "
                                         "read. Copy them into To_Fix.xlsx, then delete the copy.")
        edits = exports.read_to_fix()
        who = os.environ.get("USERNAME") or os.environ.get("USER") or "staff"
        for e in edits:
            d = db.get_doc(self.conn, e["ref"])
            if not d or d["status"] not in (db.UNASSIGNED, db.UNRECOGNISED):
                continue
            changes = {}
            for key, column in (("company_folder", "company_folder"), ("supplier", "supplier"),
                                ("number", "number"), ("doc_type", "doc_type")):
                if key in e and e[key] != (d.get(column) or ""):
                    db.add_correction(self.conn, d["ref"], column, d.get(column), e[key], who)
                    changes[column] = e[key]
            if "company_folder" in changes:
                changes["company_source"] = "manual"
            if e.get("site_folder"):                                  # a site folder chosen in To_Fix
                chosen = sites.find_site(e["site_folder"])
                name = chosen["name"] if chosen else e["site_folder"]
                if name != (d.get("site") or ""):
                    db.add_correction(self.conn, d["ref"], "site", d.get("site"), name, who)
                    changes["site"], changes["site_source"] = name, "manual"
                    if d.get("ship_to") and settings.run_bool("Learn sites from staff fixes", True):
                        owner = companies.by_name(chosen["company"]) if chosen and chosen.get("company") else companies.by_folder(e.get("company_folder", ""))
                        for key in sites.ship_keys(d["ship_to"]):          # the next document with this deliver-to number / address goes there by itself
                            db.learn(self.conn, "ship-to", settings.supplier_key(d.get("supplier") or ""), key, name,
                                     owner["Folder name"] if owner else "", example=d["ship_to"], by=who)
                        self.summary.c["sites learned"] += 1
            if not changes:
                continue
            db.update_doc(self.conn, d["ref"], fixed_by=who, **changes)
            self.changed = True
            self.touched_months.add(self._month_of(self._fields_from_row(db.get_doc(self.conn, d["ref"]))))   # its month's Record must be rewritten
            self.summary.c["fixes"] += 1
            self._reroute(db.get_doc(self.conn, d["ref"]), human=True)
        self.conn.commit()

    def _recheck_waiting(self):
        for d in db.waiting_docs(self.conn, self.conf["max_wait"]):
            if self._should_stop():
                break
            self._reroute(d, human=False)
        self.conn.commit()

    def _reroute(self, d, human):
        src = d["current_path"]
        if not src or not os.path.exists(src):
            db.update_doc(self.conn, d["ref"], reason="the file is missing from its holding folder", attempts=(d["attempts"] or 0) + 1)
            return
        f = self._fields_from_row(d)
        if not human and d["company_source"] != "manual":
            f["company_folder"] = d["company_folder"] or ""
        if not human and sites.ship_hold(d["site_source"]) and (d.get("ship_to") or ""):
            self._recheck_ship_to(d, f)
        elif not human and str(d.get("ship_to") or "").startswith("merchant:") and f["site"] == settings.UNASSIGNED:
            got = db.learned_site(self.conn, "ship-to", settings.supplier_key(d["supplier"] or ""), [d["ship_to"]])
            if got:
                f["site"], f["site_source"] = got[0], f"{d['ship_to']} (learned {got[2][:10]})"
                db.update_doc(self.conn, d["ref"], site=f["site"], site_source=f["site_source"])
        out = self._decide_and_place(f, src)
        if out["status"] in (db.FILED, db.PARKED):
            if os.path.abspath(os.path.join(out["dest_dir"], out["dest_name"])) != os.path.abspath(src):
                try:
                    os.remove(src)
                except OSError:
                    pass
            remarks = "; ".join(x for x in [d.get("remarks") or ""] + [r for r in out.get("remarks", []) if r not in (d.get("remarks") or "")] if x)
            db.update_doc(self.conn, d["ref"], status=out["status"], route=out["status"], reason="",
                          dest_dir=out["dest_dir"], dest_name=out["dest_name"],
                          current_path=os.path.join(out["dest_dir"], out["dest_name"]),
                          company_folder=out["company_folder"], company_source=out["company_source"],
                          filed_at=db.now() if out["status"] == db.FILED else None, remarks=remarks,
                          month_case=out.get("month_case", "") if out["status"] == db.FILED else "",
                          month_target=out.get("month_target", "") if out["status"] == db.FILED else "")
            self.changed = True
            self.summary.c["released"] += 1
            self.touched_months.add(self._month_of(f))
            if out["status"] == db.FILED:
                self.filed_now.append(dict(db.get_doc(self.conn, d["ref"]), site=d["site"]))
        else:
            if out["reason"] != (d["reason"] or "") or out["status"] != d["status"]:
                self.changed = True                                           # a waiting file only shows up again if its reason changed
                self.touched_months.add(self._month_of(f))
            db.update_doc(self.conn, d["ref"], reason=out["reason"], attempts=(d["attempts"] or 0) + 1,
                          status=out["status"], dest_dir=out["dest_dir"], dest_name=out["dest_name"],
                          current_path=os.path.join(out["dest_dir"], out["dest_name"]))
            if os.path.abspath(os.path.join(out["dest_dir"], out["dest_name"])) != os.path.abspath(src):
                try:
                    os.remove(src)
                except OSError:
                    pass

    def _recheck_ship_to(self, d, f):
        """A document that waits because its ship-to address was not a known site: look again (the Sites sheet may have the site now).
        When the site is found, the site AND its company decide, exactly as for a new document."""
        blocks = [sites.SHIP_TAG + b for b in str(d["ship_to"]).split(" || ") if b]
        m = sites.match("", blocks=blocks, supplier=d["supplier"] or "")
        if m["site"] == settings.UNASSIGNED and self._resolve_unknown_ship_to(f, d["ship_to"]):
            db.update_doc(self.conn, d["ref"], site=f["site"], site_source=f["site_source"])
            return
        if m["site"] == settings.UNASSIGNED:
            if m["source"] != d["site_source"]:
                f["site_source"] = m["source"]
                db.update_doc(self.conn, d["ref"], site_source=m["source"])
            return
        owner = companies.by_name(m["company"])
        f["site"], f["site_source"] = m["site"], m["source"] + " (site added later)"
        if owner:
            f["company_folder"], f["company_source"] = owner["Folder name"], "site"
        db.update_doc(self.conn, d["ref"], site=f["site"], site_source=f["site_source"])

    # ------------------------------------------------------------------------ sites chosen in To_Fix > New sites
    def _apply_site_choices(self):
        """A person typed the site for a document that was filed in the company folder (SITE NOT FOUND): move the file into the site's folder.
        A site that is not on the list is never created: the document stays and Remarks says so."""
        who = os.environ.get("USERNAME") or os.environ.get("USER") or "staff"
        for ref, typed in exports.read_new_sites():
            d = db.get_doc(self.conn, ref)
            if not d or d["status"] != db.FILED or "SITE NOT FOUND" not in (d.get("remarks") or ""):
                continue
            chosen = sites.find_site(typed)
            if not chosen:
                note = f"SITE NAME NOT ON LIST: {typed}"
                if note not in (d.get("remarks") or ""):
                    db.update_doc(self.conn, ref, remarks="; ".join(x for x in (d.get("remarks") or "", note) if x))
                    self.changed = True
                    self.touched_months.add(self._month_of(self._fields_from_row(d)))
                continue
            try:
                new_dir = os.path.join(d["dest_dir"], sites.folder_for(chosen["name"]))
                same = router_mod._same_folder_other_spelling(d["dest_dir"], sites.folder_for(chosen["name"]))
                if same:
                    new_dir = os.path.join(d["dest_dir"], same)
                if not os.path.isdir(new_dir) and not self.router.create["site"] and self.router.mode != "test":
                    raise PlacementError(f"site folder does not exist: {new_dir}")
                new_path = self.router.move(d["current_path"], new_dir, d["dest_name"])
            except PlacementError as e:
                note = f"COULD NOT MOVE TO SITE: {e.reason}"
                if note not in (d.get("remarks") or ""):
                    db.update_doc(self.conn, ref, remarks="; ".join(x for x in (d.get("remarks") or "", note) if x))
                    self.changed = True
                    self.summary.warnings.append(f"{ref}: {note}")
                continue
            remarks = "; ".join(x for x in (d.get("remarks") or "").split("; ") if x and x != "SITE NOT FOUND" and not x.startswith(("SITE NAME NOT ON LIST", "COULD NOT MOVE TO SITE")))
            db.add_correction(self.conn, ref, "site", d.get("site"), chosen["name"], who)
            db.update_doc(self.conn, ref, site=chosen["name"], site_source="manual (New sites)", remarks=remarks, fixed_by=who,
                          dest_dir=os.path.dirname(new_path), dest_name=os.path.basename(new_path), current_path=new_path)
            self.changed = True
            self.summary.c["fixes"] += 1
            self.touched_months.add(self._month_of(self._fields_from_row(db.get_doc(self.conn, ref))))
        self.conn.commit()

    # ------------------------------------------------------------------------ month folders: build the next month, move waiting invoices
    def _month_housekeeping(self):
        """In the last two days of a month: copy the structure of the latest month to the next month's folder. On the first run of a new month
        build its folder if it is missing. Then move the future-dated invoices that waited in an earlier month into their own month."""
        rt = self.router
        if not rt.starts_with_month() or not (os.path.isdir(rt.root) or rt.mode == "test"):
            return
        today = self.now.date()
        if self.conf["make_trees"]:
            this_month = today.replace(day=1)
            next_month = (this_month + datetime.timedelta(days=32)).replace(day=1)
            wanted = []
            if not os.path.isdir(rt.month_dir(this_month)):
                wanted.append(this_month)                                      # the first run of the month, and nobody built the folder
            if (next_month - today).days <= 2:
                wanted.append(next_month)                                      # the last two days of the month (also tops up a folder that exists)
            for month in wanted:
                source = rt.latest_month_dir_before(month)
                if not source:
                    if rt.mode != "live":
                        continue
                    self.summary.warnings.append(f"no earlier month folder to copy the structure from, so {os.path.basename(rt.month_dir(month))} was not built")
                    continue
                try:
                    made = rt.copy_tree(source, rt.month_dir(month))
                except OSError as e:
                    self.summary.warnings.append(f"could not build {os.path.basename(rt.month_dir(month))}: {e}")
                    continue
                if made:
                    self.summary.notes.append(f"Built the folders of {os.path.basename(rt.month_dir(month))} ({made} folder(s), copied from {os.path.basename(source)}).")
                    self.log(self.summary.notes[-1])
        self._move_waiting_future_invoices()

    def _move_waiting_future_invoices(self):
        rt = self.router
        today = self.now.date()
        for d in [dict(r) for r in self.conn.execute("SELECT * FROM docs WHERE status=? AND month_case='future' ORDER BY id", (db.FILED,))]:
            try:
                year, month = (int(x) for x in d["month_target"].split("-"))
                target_month = datetime.date(year, month, 1)
            except (ValueError, AttributeError):
                continue
            if not os.path.isdir(rt.month_dir(target_month)):
                continue                                                       # its month's folder does not exist yet: keep waiting
            try:
                parts = os.path.relpath(d["dest_dir"], rt.root).split(os.sep)
                new_dir = os.path.join(rt.root, month_folder_clean(target_month), *parts[1:])
                new_path = rt.move(d["current_path"], new_dir, d["dest_name"])
            except (PlacementError, ValueError, OSError) as e:
                reason = getattr(e, "reason", str(e))
                note = f"COULD NOT MOVE TO {router_months(target_month)} {year}: {reason}"
                if note not in (d.get("remarks") or ""):
                    db.update_doc(self.conn, d["ref"], remarks="; ".join(x for x in (d.get("remarks") or "", note) if x))
                    self.summary.warnings.append(f"{d['ref']}: {note} - a person must move it")
                    self.changed = True
                    self.touched_months.add(self._month_of(self._fields_from_row(d)))
                continue
            moved = f"MOVED TO {router_months(target_month)} {year} on {naming.format_date(today, self.conf['name_date'])}"
            db.update_doc(self.conn, d["ref"], month_case="moved", moved_on=today.isoformat(), remarks=moved,
                          dest_dir=os.path.dirname(new_path), dest_name=os.path.basename(new_path), current_path=new_path)
            self.changed = True
            self.summary.c["future invoices moved"] += 1
            self.touched_months.add(self._month_of(self._fields_from_row(db.get_doc(self.conn, d["ref"]))))       # the month it arrived in (orange row)
            self.touched_months.add((year, month))                                                                  # the month it moved to (light-blue row)
        self.conn.commit()

    # ------------------------------------------------------------------------ the end of a live run
    def _finish_live(self):
        s = self.summary
        warn = self.router.month_warning()
        if warn:
            s.warnings.append(warn)
        # Excel files: rewritten only when something they show has changed (or they are missing) - an idle run writes nothing
        written, current = [], (self.now.year, self.now.month)
        if self.run_last:                                        # move the 'last e-mail of the run' markers to this run
            import json
            stamp = self.now.strftime("%d-%m-%Y %H:%M")
            old = db.run_markers(self.conn)
            new = {ref: f"LAST E-MAIL OF RUN {stamp} ({mb})" for mb, (_rcv, ref) in self.run_last.items()}
            for ref in set(old) | set(new):
                d = db.get_doc(self.conn, ref)
                if d:
                    self.touched_months.add(self._month_of(self._fields_from_row(d)))
            db.set_meta(self.conn, "run_markers", json.dumps(new))
            self.conn.commit()
        months = set(self.touched_months)
        if not os.path.exists(exports.record_path(*current)):
            months.add(current)
        for (year, month) in sorted(months):
            try:
                exports.write_monthly(self.conn, year, month, self.conf["month_basis"])
                written.append(f"Record_{year:04d}-{month:02d}")
            except Exception as e:
                s.warnings.append(f"could not write the monthly record {year}-{month:02d}: {e}")
        if self.changed or not os.path.exists(os.path.join(exports.reports_dir(), "To_Fix.xlsx")):
            try:
                exports.write_to_fix(self.conn, s.new_suppliers)
                written.append("To_Fix")
            except Exception as e:
                s.warnings.append(f"could not write To_Fix.xlsx: {e}")
        try:                                                     # fuel invoices: new rows at the bottom of the month's list
            files, added, busy = exports.append_fuel_list(self.conn)
            if added:
                written.append(f"PI_Invoice_list (+{added})")
            for b in busy:
                s.warnings.append(f"{b} is open in Excel: its new fuel invoices are added at the next run (close the file)")
        except Exception as e:
            s.warnings.append(f"could not update the fuel invoice list: {e}")
        fleet_months = {(int(str(d["created_at"])[:4]), int(str(d["created_at"])[5:7])) for d in self.conn.execute(
            "SELECT created_at FROM docs WHERE doc_type='Fleet Card Statement' AND updated_at>=?", (self.started_iso,))}
        for (year, month) in sorted(fleet_months):
            try:
                exports.write_fleet_record(self.conn, year, month)
                written.append(f"FleetCard_Record_{exports.MONTH_ABBR[month - 1]}_{year}")
            except Exception as e:
                s.warnings.append(f"could not write the Fleet Card record {year}-{month:02d}: {e}")
        s.notes.append("Excel files updated: " + (", ".join(written) if written else "none (nothing new)"))
        if self.conf["path_mode"] in ("draft", "send"):
            n = notify.path_emails(self.gw, self.filed_now, self.conf["path_mode"], self.log)
            s.notes.append(f"Path e-mails ({self.conf['path_mode']}): {n}")
        if settings.run_bool("Quarterly records archive", True):
            try:
                made, count = quarterly.make_due(self.conn, self.now.date(), "invoice" if self.conf["month_basis"] == "invoice" else "arrival")
                if made:
                    s.notes.append(f"Quarterly records archive written: {made} ({count} documents)")
                for path, n in quarterly.make_features(self.conn, *quarterly.last_finished(self.now.date())):
                    s.notes.append(f"Quarterly archive written: {path} ({n} rows)")
            except Exception as e:
                s.warnings.append(f"quarterly records archive failed (it is tried again at the next run): {e}")
        self.router.cleanup_archive(self.conf["archive_days"])
        backup_dir = settings.run_value("Backup folder") or paths.p("data", "backups")
        try:
            db.backup(self.conn, backup_dir, settings.run_int("Backup keep days", 14))
        except Exception as e:
            s.warnings.append(f"database backup failed: {e}")
        waiting = db.counts_by_status(self.conn)
        s.notes.append(f"Waiting right now: {waiting.get(db.UNASSIGNED, 0)} in _Unassigned, "
                       f"{waiting.get(db.UNRECOGNISED, 0)} in _Unrecognised (see Reports\\To_Fix.xlsx)")
        db.record_run(self.conn, f"{self.now:%Y-%m-%d %H:%M:%S}", "live", s.text())
