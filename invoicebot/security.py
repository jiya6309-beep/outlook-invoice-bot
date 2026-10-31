"""The security run: the weekly alarm reports from the Security mailbox -> the fortnightly STORE OPEN & CLOSE Excel.

Separate from the invoices: its own button, its own schedule, its own settings file (Security_Settings.xlsx). It reads
ONLY the mailbox with Role 'Security' (SpecialMailboxes sheet) and never touches an invoice.

The rules (agreed with the company):
  OPEN   = the first disarm of the day. A disarm up to 3.5 h before opening counts for that day; earlier = the day before.
  CLOSE  = the last arm of the day. An arm up to 4 h after closing still counts; later = CLOSE is MISSING.
  OPEN label yellow   opened more than 10 minutes late, or no disarm.
  CLOSE label yellow  armed even 1 minute before closing, or no arm.
  Both labels yellow  the site sent no report that week (NO REPORT), or it is closed that day.
  Timing and the times themselves are never highlighted. Seconds are ignored. Times are used as printed (site local time).

Files:  <Destination folder>\\<MONTH YEAR>\\Downloaded Reports\\<report>.pdf
        <Destination folder>\\<MONTH YEAR>\\SITE_DATA_<MON>_<YEAR>_<dd-mm>_TO_<dd-mm>.xlsx   (two weeks; rebuilt each run)
The events are kept in the bot's database, so a late arm that appears in the next week's report is still counted."""
import datetime
import hashlib
import os
import shutil
import time

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from . import companies, db, paths, security_reader, security_settings as ss, settings
from .runlock import LockBusy, RunLock

SCHEMA = """
CREATE TABLE IF NOT EXISTS security_events (
    client TEXT, at TEXT, kind TEXT, line TEXT, file TEXT, PRIMARY KEY(client, at, kind));
CREATE TABLE IF NOT EXISTS security_reports (
    client TEXT, week TEXT, file TEXT, entry_id TEXT, saved_as TEXT, received TEXT, PRIMARY KEY(client, week));
"""
MONTHS = ["JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE", "JULY", "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER"]
LABEL = "SECURITY"                                            # the mailbox label used for its e-mail records / catch-up window


def ensure_schema(conn):
    conn.executescript(SCHEMA)


# ---------------------------------------------------------------------------------------------- fortnights
def monday(day):
    return day - datetime.timedelta(days=day.weekday())


def fortnight(day, anchor):
    """(first Monday, last Sunday, label month date) of the two-week Excel that holds this day."""
    k = (monday(day) - anchor).days // 14
    start = anchor + datetime.timedelta(days=14 * k)
    end = start + datetime.timedelta(days=13)
    return start, end, start + datetime.timedelta(days=7)    # the file is named by the month of its SECOND Monday


def month_folder(label_day):
    return f"{MONTHS[label_day.month - 1]} {label_day.year}"


def file_name(start, end, label_day):
    return f"SITE_DATA_{MONTHS[label_day.month - 1][:3]}_{label_day.year}_{start:%d-%m}_TO_{end:%d-%m}.xlsx"


def anchor_of(conf):
    try:
        return monday(datetime.date.fromisoformat(str(conf["run"].get("Fortnight starts on", "2026-10-05")).strip()[:10]))
    except ValueError:
        raise ss.SecurityError("'Fortnight starts on' in Security_Settings.xlsx (sheet Run) must be a date like 2026-10-05.")


# ---------------------------------------------------------------------------------------------- the day rules
def _events_for(conn, client, start, end):
    s = (start - datetime.timedelta(days=2)).isoformat()
    e = (end + datetime.timedelta(days=2)).isoformat()
    return [(datetime.datetime.fromisoformat(r[0]), r[1], r[2]) for r in conn.execute(
        "SELECT at, kind, line FROM security_events WHERE client=? AND at>=? AND at<? ORDER BY at", (client, s, e))]


def day_results(conn, site, start, end, rules, covered):
    """{date: result} for every day start..end. result = {'state': 'ok'|'closed'|'24/7'|'no report'|'not yet',
    'open','close' (minutes), 'in','out' (datetime or None), 'in_line','out_line', 'open_flag','close_flag', ...}."""
    allow, arm_h, dis_h = rules
    days = [start + datetime.timedelta(days=i) for i in range((end - start).days + 1)]

    def window(d):
        h = site["hours"].get(d.weekday())
        if not h:
            return None
        o, c = h
        base = datetime.datetime.combine(d, datetime.time())
        open_dt = base + datetime.timedelta(minutes=o)
        close_dt = base + datetime.timedelta(minutes=c if c > o else c + 1440)
        return open_dt, close_dt

    span = [start + datetime.timedelta(days=i) for i in range(-2, (end - start).days + 3)]
    wins = {d: window(d) for d in span}
    assigned = {d: {"disarm": [], "arm": []} for d in span}
    for at, kind, line in _events_for(conn, site["client"], start, end):
        owner = None
        for d in span:                                      # the latest day whose 'opening - 3.5 h' is not after the event
            w = wins[d]
            if w and w[0] - datetime.timedelta(hours=dis_h) <= at:
                owner = d
        if owner is None:
            continue
        if kind == "arm" and at > wins[owner][1] + datetime.timedelta(hours=arm_h):
            continue                                         # armed too late: CLOSE is missing for that day
        assigned[owner][kind].append((at, line))
    out = {}
    for d in days:
        wk = monday(d)
        if wk not in covered:
            out[d] = {"state": "not yet"}
            continue
        h = site["hours"].get(d.weekday())
        if site["always_open"]:
            out[d] = {"state": "24/7"}
            continue
        if not covered[wk]:
            out[d] = {"state": "no report", "open": h[0] if h else None, "close": h[1] if h else None}
            continue
        if not h:
            out[d] = {"state": "closed"}
            continue
        o, c = h
        dis, arm = sorted(assigned[d]["disarm"]), sorted(assigned[d]["arm"])
        r = {"state": "ok", "open": o, "close": c, "in": dis[0][0] if dis else None, "in_line": dis[0][1] if dis else "",
             "out": arm[-1][0] if arm else None, "out_line": arm[-1][1] if arm else ""}
        base = datetime.datetime.combine(d, datetime.time())
        if r["in"]:
            diff = int((r["in"] - (base + datetime.timedelta(minutes=o))).total_seconds() // 60)
            r["open_diff"] = diff
            r["open_remark"] = "EARLY" if diff < 0 else ("ON TIME" if diff == 0 else "LATE")
            r["open_flag"] = diff > allow
        else:
            r["open_remark"], r["open_flag"] = "MISSING", True
        if r["out"]:
            close_dt = base + datetime.timedelta(minutes=c if c > o else c + 1440)
            diff = int((r["out"] - close_dt).total_seconds() // 60)
            r["close_diff"] = diff
            r["close_remark"] = "EARLY" if diff < 0 else ("ON TIME" if diff == 0 else "LATE")
            r["close_flag"] = diff < 0
        else:
            r["close_remark"], r["close_flag"] = "MISSING", True
        out[d] = r
    return out


def coverage(conn, sites, start, end):
    """{week Monday: {client: True}} for weeks of which at least one report was read; a site without a report that week is
    marked False there (-> NO REPORT)."""
    out = {}
    for row in conn.execute("SELECT DISTINCT week FROM security_reports WHERE week>=? AND week<=?", (start.isoformat(), end.isoformat())):
        wk = datetime.date.fromisoformat(row[0])
        have = {r[0] for r in conn.execute("SELECT client FROM security_reports WHERE week=?", (row[0],))}
        out[wk] = {s["client"]: s["client"] in have for s in sites}
    return out


# ---------------------------------------------------------------------------------------------- the Excel
THIN = Side(style="thin", color="000000")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BOLD = Font(bold=True)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
YELLOW = PatternFill("solid", fgColor="FFFF00")
RED = Font(bold=True, color="C00000")
NAVY = Font(bold=True, color="002060")


def _hm(minutes):
    m = abs(int(minutes))
    return f"{m // 60}:{m % 60:02d}"


def _clock(dt):
    return ss.show_time(dt.hour * 60 + dt.minute) if dt else ""


def write_excel(conn, conf, start, end, label_day, folder):
    """Writes the fortnight's SITE_DATA_... file (rebuilt completely each time - nothing in it is typed by staff)."""
    rules = (ss.number(conf["run"], "Late opening allowance (minutes)", 10), ss.number(conf["run"], "Late arming window (hours)", 4),
             ss.number(conf["run"], "Early disarm window (hours)", 3.5))
    sites = [s for s in conf["sites"] if s["active"]]
    cover = {wk: {c: v for c, v in d.items()} for wk, d in coverage(conn, sites, start, end).items()}
    days = [start + datetime.timedelta(days=i) for i in range(14)]
    wb = Workbook()
    ws = wb.active
    ws.title = "SITE DATA"
    # columns: A company, B no, C site, D client, E address, F labels, then per week 14 day columns + 1 gap
    col_of = {}
    c = 7
    for i, d in enumerate(days):
        if i == 7:
            c += 1                                           # a narrow gap column between the two weeks
        col_of[d] = c
        c += 2
    last = c - 1

    def put(r, col, v=None, font=BOLD, fill=None, fmt=None):
        cell = ws.cell(r, col, v)
        cell.border, cell.alignment, cell.font = BOX, CENTER, font
        if fill:
            cell.fill = fill
        if fmt:
            cell.number_format = fmt
        return cell

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last)
    ws.cell(1, 1, "STORE OPEN & CLOSE DATA").font = Font(bold=True, size=14)
    ws.cell(1, 1).alignment = CENTER
    label = (f"{MONTHS[start.month - 1][:3]} - {MONTHS[end.month - 1][:3]} {end.year}" if start.month != end.month
             else f"{MONTHS[start.month - 1][:3]} {start.year}")
    for r in range(2, 5):
        for col in range(1, 4):
            put(r, col)
    ws.merge_cells("A2:C4")
    ws.cell(2, 1, f"{label}\n{start:%d-%m} TO {end:%d-%m}")
    put(2, 4, "Client No")
    put(2, 5, "ADDRESS")
    for r in (3, 4):
        put(r, 4)
        put(r, 5)
    for r, t in ((2, "Date"), (3, "Day"), (4, "Week")):
        put(r, 6, t)
    for d in days:
        col = col_of[d]
        for r, v in ((2, d.day), (3, d.strftime("%a").lower())):
            put(r, col, v)
            put(r, col + 1)
            ws.merge_cells(start_row=r, start_column=col, end_row=r, end_column=col + 1)
    for wk in (days[0], days[7]):
        a, b = col_of[wk], col_of[wk + datetime.timedelta(days=6)] + 1
        for col in range(a, b + 1):
            put(4, col)
        ws.merge_cells(start_row=4, start_column=a, end_row=4, end_column=b)
        ws.cell(4, a, f"WEEK-{(wk.day - 1) // 7 + 1}")
    row = 6
    stats = {"ok": 0, "open_late": 0, "close_early": 0, "missing": 0, "no_report": 0}
    for site in sites:
        res = day_results(conn, site, start, end, rules, {wk: v.get(site["client"], False) for wk, v in cover.items()})
        for col, v in ((1, site["company"]), (2, site["no"]), (3, site["name"]), (4, site["client"]), (5, site["address"])):
            for r in range(row, row + 6):
                put(r, col)
            ws.cell(row, col, v)
            ws.merge_cells(start_row=row, start_column=col, end_row=row + 5, end_column=col)
        for k, t in enumerate(["Timing", "", "EXPECTED", "ACTUAL", "HOURS", "Remark"]):
            put(row + k, 6, t)
        for d in days:
            col, r = col_of[d], res[d]
            h = site["hours"].get(d.weekday())
            put(row, col, ss.timing_text(*h) if h and not site["always_open"] else ("24/7" if site["always_open"] else "CLOSED"))
            put(row, col + 1)
            ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + 1)
            o_lab, c_lab = put(row + 1, col, "OPEN"), put(row + 1, col + 1, "CLOSE")
            for k in range(2, 6):
                put(row + k, col, font=Font())
                put(row + k, col + 1, font=Font())
            state = r["state"]
            if state in ("not yet", "24/7"):
                continue
            if h:
                ws.cell(row + 2, col, ss.show_time(h[0]))
                ws.cell(row + 2, col + 1, ss.show_time(h[1]))
            if state in ("no report", "closed"):
                text = "NO REPORT" if state == "no report" else "CLOSED"
                for cc, lab in ((col, o_lab), (col + 1, c_lab)):
                    ws.cell(row + 5, cc, text).font = RED
                    lab.fill = YELLOW
                    lab.comment = Comment("No report from this site this week" if state == "no report" else
                                          "The site is closed this day (no hours in Security_Settings.xlsx)", "Bot")
                stats["no_report" if state == "no report" else "missing"] += 1
                continue
            for cc, lab, key, side in ((col, o_lab, "in", "open"), (col + 1, c_lab, "out", "close")):
                val = r[key]
                act = ws.cell(row + 3, cc, _clock(val))
                if val:
                    act.comment = Comment(r[f"{key}_line"], "Bot")
                    ws.cell(row + 4, cc, _hm(r[f"{side}_diff"])).font = NAVY
                ws.cell(row + 5, cc, r[f"{side}_remark"]).font = RED
                if r[f"{side}_flag"]:
                    lab.fill = YELLOW
                    why = ("No disarm for this day" if side == "open" and not val else "No arm for this day (or armed more than "
                           f"{rules[1]:g} h after closing)" if not val else f"Opened {_hm(r['open_diff'])} late" if side == "open"
                           else f"Closed {_hm(r['close_diff'])} early")
                    lab.comment = Comment(why, "Bot")
            if r["open_flag"] or r["close_flag"]:
                if not r["in"] or not r["out"]:
                    stats["missing"] += 1
                if r["in"] and r["open_flag"]:
                    stats["open_late"] += 1
                if r["out"] and r["close_flag"]:
                    stats["close_early"] += 1
            else:
                stats["ok"] += 1
        row += 7                                             # one empty row between sites
    widths = {1: 22, 2: 6, 3: 26, 4: 10, 5: 30, 6: 11}
    for col in range(1, last + 1):
        ws.column_dimensions[get_column_letter(col)].width = widths.get(col, 11)
    ws.column_dimensions[get_column_letter(col_of[days[7]] - 1)].width = 2
    ws.freeze_panes = "G6"
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    os.makedirs(folder, exist_ok=True)
    target = os.path.join(folder, file_name(start, end, label_day))
    try:
        wb.save(target)
    except PermissionError:
        base, ext = os.path.splitext(target)
        target = f"{base}_new{ext}"
        wb.save(target)
    return target, stats


# ---------------------------------------------------------------------------------------------- the run
def security_mailbox():
    for sp in companies.special_rows():
        if sp["role"] == "security":
            return sp
    return None


class SecurityRun:
    """One security run. shadow=True: a rehearsal - reads the e-mails and reports what it would do, writes nothing."""

    def __init__(self, gateway, shadow=False, from_date=None, to_date=None, log=print, wait_minutes=30):
        self.gw, self.shadow, self.log = gateway, shadow, log
        self.from_date, self.to_date = from_date, to_date
        self.wait_minutes = wait_minutes
        self.lines, self.warnings = [], []
        self.c = {"emails looked at": 0, "reports read": 0, "sites read": 0, "skipped (no arm/disarm)": 0,
                  "skipped (Client No not in the list)": 0, "already done": 0}
        self.unknown = set()
        self.files = []

    def run(self):
        lock = None
        if not self.shadow:
            lock = RunLock(paths.p("data", "run.lock"), 60)
            deadline = time.time() + self.wait_minutes * 60
            while True:                                      # an invoice run in progress: wait for it, never run at the same time
                try:
                    lock.__enter__()
                    break
                except LockBusy:
                    if time.time() > deadline:
                        self.warnings.append("another run kept the lock too long - the security run did not start")
                        return self
                    self.log("Waiting for the invoice run to finish ...")
                    time.sleep(20)
        try:
            self._run()
        finally:
            if lock is not None:
                lock.__exit__(None, None, None)
        return self

    def _run(self):
        conf = self.conf = ss.load()
        sp = security_mailbox()
        if sp is None:
            raise ss.SecurityError("No mailbox has the Role 'Security' in Bot_Settings.xlsx (sheet SpecialMailboxes).")
        dest = sp["dest"]
        if not dest:
            raise ss.SecurityError("The Security mailbox has no Destination folder (SpecialMailboxes sheet).")
        if not settings.is_test_mode():
            settings.check_output_root(dest, "Security reports")
        elif not self.shadow:
            os.makedirs(dest, exist_ok=True)
        anchor = anchor_of(conf)
        known = {s["client"]: s for s in conf["sites"]}
        conn = db.connect(paths.p("data", "bot.db"))
        ensure_schema(conn)
        try:
            store = self.gw.resolve_store(sp["names"])
            now = datetime.datetime.now()
            if self.from_date:
                start = datetime.datetime.combine(self.from_date, datetime.time())
                end = datetime.datetime.combine(self.to_date or self.from_date, datetime.time(23, 59, 59))
            else:
                last = db.last_success(conn, LABEL)
                start = (last - datetime.timedelta(hours=1)) if last else now - datetime.timedelta(days=8)
                end = now
            mails, _trunc = self.gw.scan(store, LABEL, start, end, 2000)
            touched = set()
            done = {k for k, (outcome, _n) in db.seen_emails(conn).items() if k.startswith(LABEL + "|") and outcome == "security"}
            for info in sorted(mails, key=lambda m: m.received):
                self.c["emails looked at"] += 1
                if info.done or db.email_key(LABEL, info.entry_id) in done:
                    self.c["already done"] += 1
                    continue
                used = self._mail(conn, info, known, anchor, dest, touched)
                if used and not self.shadow:
                    try:
                        self.gw.mark_done(info)                  # read + tick
                    except Exception as e:
                        self.warnings.append(f"could not tick '{info.subject}': {e}")
                    db.record_email(conn, LABEL, info.entry_id, info.subject, info.received, "security")
                conn.commit()
            if not self.shadow:
                if not self.from_date:
                    db.record_mailbox_run(conn, LABEL, start, end, True)
                for (fs, fe, label_day) in sorted(touched):
                    folder = os.path.join(dest, month_folder(label_day))
                    path, stats = write_excel(conn, conf, fs, fe, label_day, folder)
                    self.files.append((path, stats))
                conn.commit()
        finally:
            conn.close()

    def _mail(self, conn, info, known, anchor, dest, touched):
        used = False
        scratch = paths.ensure("Scratch")
        for att in self.gw.attachments(info):
            if att.hidden or not att.name.lower().endswith(".pdf"):
                continue
            tmp = os.path.join(scratch, f"sec_{hashlib.sha1((info.entry_id + att.name).encode()).hexdigest()[:10]}.pdf")
            try:
                self.gw.save_attachment(info, att.index, tmp)
                sections = security_reader.parse(tmp)
            except Exception as e:
                self.warnings.append(f"could not read {att.name}: {e}")
                continue
            self.c["reports read"] += 1
            kept = False
            for sec in sections:
                if not sec["events"]:
                    self.c["skipped (no arm/disarm)"] += 1
                    continue
                if sec["client"] not in known:
                    self.c["skipped (Client No not in the list)"] += 1
                    self.unknown.add(f"{sec['client'] or '?'} {sec['name']}".strip())
                    continue
                week = monday(sec["start"]) if sec["start"] else monday(min(e[0] for e in sec["events"]).date())
                fs, fe, label_day = fortnight(week, anchor)
                touched.add((fs, fe, label_day))
                kept = used = True
                self.c["sites read"] += 1
                if self.shadow:
                    self.lines.append(f"{sec['client']} {known[sec['client']]['name']}: week {week:%d-%m} - {len(sec['events'])} arm/disarm events")
                    continue
                for at, kind, line in sec["events"]:
                    conn.execute("INSERT OR IGNORE INTO security_events(client, at, kind, line, file) VALUES(?,?,?,?,?)",
                                 (sec["client"], at.isoformat(sep=" "), kind, line, att.name))
                conn.execute("INSERT OR REPLACE INTO security_reports(client, week, file, entry_id, received) VALUES(?,?,?,?,?)",
                             (sec["client"], week.isoformat(), att.name, info.entry_id, info.received.isoformat(sep=" ")))
                folder = os.path.join(dest, month_folder(label_day), "Downloaded Reports")
                if str(self.conf["run"].get("Keep downloaded reports", "Yes")).lower().startswith("y") and not os.path.exists(
                        os.path.join(folder, att.name)):
                    os.makedirs(folder, exist_ok=True)
                    shutil.copy2(tmp, os.path.join(folder, att.name))
            if kept and not self.shadow:
                arch = paths.ensure("Archive", datetime.date.today().isoformat())
                try:
                    shutil.copy2(tmp, os.path.join(arch, "SEC_" + att.name))
                except OSError:
                    pass
            try:
                os.remove(tmp)
            except OSError:
                pass
        return used

    def summary(self):
        out = ["SECURITY RUN" + (" (rehearsal - nothing changed)" if self.shadow else "")]
        out += [f"  {k}: {v}" for k, v in self.c.items() if v]
        if self.unknown:
            out.append("  Client No not in Security_Settings.xlsx (add them if they are real sites): " + ", ".join(sorted(self.unknown)))
        for path, st in self.files:
            out.append(f"  Excel: {path}")
            out.append(f"     days OK {st['ok']}, opened late {st['open_late']}, closed early {st['close_early']}, "
                       f"missing {st['missing']}, no report {st['no_report']}")
        out += self.lines
        out += [f"  WARNING: {w}" for w in self.warnings]
        return "\n".join(out)
