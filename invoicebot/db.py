"""The bot's memory: one SQLite file on the bot PC (data/bot.db).

SQLite is built into Python, needs no server, writes thousands of rows per second and is safe against
crashes. Excel files for people are generated FROM this database (see exports.py).

    docs          one row per document the bot has handled
    emails        every e-mail the bot has acted on (so nothing is processed twice)
    mailbox_runs  the end of the last successful window per mailbox (drives the automatic catch-up)
    runs          one row per run, with its summary
    corrections   what staff changed in To_Fix.xlsx, and when
"""
import datetime
import gzip
import os
import shutil
import sqlite3

SCHEMA_VERSION = 1

SCHEMA = """
CREATE TABLE IF NOT EXISTS docs (
    id INTEGER PRIMARY KEY AUTOINCREMENT, ref TEXT UNIQUE,
    status TEXT, route TEXT, reason TEXT,
    supplier TEXT, supplier_source TEXT, doc_type TEXT, number TEXT, doc_date TEXT, amount REAL, gst REAL, ex_amount REAL, remarks TEXT,
    company_folder TEXT, company_source TEXT, site TEXT, site_source TEXT,
    mailbox TEXT, entry_id TEXT, subject TEXT, sender TEXT, received TEXT, orig_name TEXT,
    file_hash TEXT, dest_dir TEXT, dest_name TEXT, current_path TEXT, archive_path TEXT,
    warnings TEXT, method TEXT, page_count INTEGER, ocr_used INTEGER, attempts INTEGER DEFAULT 0,
    created_at TEXT, updated_at TEXT, filed_at TEXT, fixed_by TEXT, month_case TEXT, month_target TEXT, moved_on TEXT, ship_to TEXT);
CREATE INDEX IF NOT EXISTS ix_docs_hash ON docs(file_hash);
CREATE INDEX IF NOT EXISTS ix_docs_key ON docs(supplier, doc_type, number);
CREATE INDEX IF NOT EXISTS ix_docs_status ON docs(status);
CREATE INDEX IF NOT EXISTS ix_docs_received ON docs(received);
CREATE TABLE IF NOT EXISTS emails (
    key TEXT PRIMARY KEY, mailbox TEXT, entry_id TEXT, subject TEXT, received TEXT,
    outcome TEXT, detail TEXT, attempts INTEGER DEFAULT 0, processed_at TEXT);
CREATE TABLE IF NOT EXISTS mailbox_runs (
    id INTEGER PRIMARY KEY AUTOINCREMENT, mailbox TEXT, window_from TEXT, window_to TEXT, ok INTEGER, finished TEXT);
CREATE TABLE IF NOT EXISTS runs (
    id INTEGER PRIMARY KEY AUTOINCREMENT, started TEXT, finished TEXT, mode TEXT, summary TEXT);
CREATE TABLE IF NOT EXISTS corrections (
    id INTEGER PRIMARY KEY AUTOINCREMENT, ref TEXT, field TEXT, old TEXT, new TEXT, by TEXT, at TEXT);
CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT);
CREATE TABLE IF NOT EXISTS learned (
    id INTEGER PRIMARY KEY AUTOINCREMENT, kind TEXT, supplier TEXT, key TEXT, site TEXT, company_folder TEXT,
    example TEXT, by TEXT, at TEXT, UNIQUE(kind, supplier, key));
"""

# statuses of a document
FILED, PARKED, UNASSIGNED, UNRECOGNISED = "Filed", "Parked", "Unassigned", "Unrecognised"
HANDLED_OUTCOMES = ("downloaded", "blocked", "image-only", "no-attachment", "category-only", "gave-up")


def now():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def connect(path):
    if path != ":memory:":
        os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    conn = sqlite3.connect(path, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
    conn.executescript(SCHEMA)
    have = {r[1] for r in conn.execute("PRAGMA table_info(docs)")}
    for column, kind in (("ex_amount", "REAL"), ("remarks", "TEXT"), ("month_case", "TEXT"), ("month_target", "TEXT"), ("moved_on", "TEXT"), ("ship_to", "TEXT"), ("reference", "TEXT"), ("listed_at", "TEXT")):              # a database made by an earlier version gets the new columns
        if column not in have:
            conn.execute(f"ALTER TABLE docs ADD COLUMN {column} {kind}")
    conn.execute("INSERT OR IGNORE INTO meta(key, value) VALUES('schema_version', ?)", (str(SCHEMA_VERSION),))
    conn.commit()
    return conn


# ---------------------------------------------------------------------------------- e-mails

def email_key(mailbox, entry_id):
    return f"{mailbox}|{entry_id}"


def seen_emails(conn):
    """{key: (outcome, attempts)} for every e-mail the bot has acted on."""
    return {r["key"]: (r["outcome"], r["attempts"]) for r in conn.execute("SELECT key, outcome, attempts FROM emails")}


def record_email(conn, mailbox, entry_id, subject, received, outcome, detail=""):
    key = email_key(mailbox, entry_id)
    row = conn.execute("SELECT attempts FROM emails WHERE key=?", (key,)).fetchone()
    attempts = (row["attempts"] + 1) if row else 1
    conn.execute(
        "INSERT INTO emails(key, mailbox, entry_id, subject, received, outcome, detail, attempts, processed_at) "
        "VALUES(?,?,?,?,?,?,?,?,?) ON CONFLICT(key) DO UPDATE SET outcome=excluded.outcome, "
        "detail=excluded.detail, attempts=excluded.attempts, processed_at=excluded.processed_at",
        (key, mailbox, entry_id, subject, str(received), outcome, detail, attempts, now()))


# ---------------------------------------------------------------------------------- documents

DOC_FIELDS = ["status", "route", "reason", "supplier", "supplier_source", "doc_type", "number", "doc_date", "amount",
              "gst", "ex_amount", "remarks", "company_folder", "company_source", "site", "site_source", "mailbox", "entry_id", "subject",
              "sender", "received", "orig_name", "file_hash", "dest_dir", "dest_name", "current_path",
              "archive_path", "warnings", "method", "page_count", "ocr_used", "attempts", "filed_at", "fixed_by",
              "month_case", "month_target", "moved_on", "ship_to", "reference"]


def insert_doc(conn, d):
    """Insert a document; returns its Ref (R000001...)."""
    values = {k: d.get(k) for k in DOC_FIELDS}
    values["created_at"] = values["updated_at"] = now()
    cols = ", ".join(values)
    cur = conn.execute(f"INSERT INTO docs({cols}) VALUES({', '.join('?' for _ in values)})", list(values.values()))
    ref = f"R{cur.lastrowid:06d}"
    conn.execute("UPDATE docs SET ref=? WHERE id=?", (ref, cur.lastrowid))
    return ref


def update_doc(conn, ref, **fields):
    fields["updated_at"] = now()
    sets = ", ".join(f"{k}=?" for k in fields)
    conn.execute(f"UPDATE docs SET {sets} WHERE ref=?", list(fields.values()) + [ref])


def get_doc(conn, ref):
    row = conn.execute("SELECT * FROM docs WHERE ref=?", (ref,)).fetchone()
    return dict(row) if row else None


def find_duplicate(conn, supplier, doc_type, number, file_hash):
    """The earlier document with the same content, or the same supplier + type + number."""
    if file_hash:
        row = conn.execute("SELECT * FROM docs WHERE file_hash=? LIMIT 1", (file_hash,)).fetchone()
        if row:
            return dict(row)
    if number and supplier and doc_type:
        row = conn.execute("SELECT * FROM docs WHERE supplier=? AND doc_type=? AND number=? LIMIT 1",
                           (supplier, doc_type, str(number))).fetchone()
        if row:
            return dict(row)
    return None


def find_fleet_duplicate(conn, merchant_key, doc_date):
    """An earlier Fleet Card statement with the same merchant number and report issue date (5.19.0: one statement can come to
    mstatement@ and to a company mailbox)."""
    row = conn.execute("SELECT * FROM docs WHERE doc_type='Fleet Card Statement' AND ship_to=? AND doc_date=? LIMIT 1",
                       (merchant_key, doc_date)).fetchone()
    return dict(row) if row else None


def waiting_docs(conn, max_attempts=50):
    return [dict(r) for r in conn.execute(
        "SELECT * FROM docs WHERE status=? AND (attempts<? OR site_source LIKE 'SHIP-TO %') ORDER BY id", (UNASSIGNED, max_attempts))]
        # a document that waits for its ship-to SITE to be added to the Sites sheet keeps being retried: it may take days


def docs_by_status(conn, *statuses):
    marks = ", ".join("?" for _ in statuses)
    return [dict(r) for r in conn.execute(f"SELECT * FROM docs WHERE status IN ({marks}) ORDER BY id", statuses)]


def known_suppliers(conn):
    """Supplier names the bot has already met (used to recognise a logo)."""
    return [r[0] for r in conn.execute(
        "SELECT DISTINCT supplier FROM docs WHERE supplier IS NOT NULL AND supplier<>'' "
        "AND supplier_source IN ('table','printed','sender','logo')")]


def between_clause(start, end, basis="invoice"):
    """(WHERE text, parameters) for the documents of a period: by the date on the document or by the e-mail's arrival."""
    col = "doc_date" if basis == "invoice" else "received"
    return f"substr({col},1,10) >= ? AND substr({col},1,10) <= ?", (str(start)[:10], str(end)[:10])


def month_clause(year, month, basis="arrival"):
    prefix = f"{year:04d}-{month:02d}"
    col = "COALESCE(NULLIF(doc_date,''), received)" if basis == "invoice" else "received"
    return f"substr({col},1,7)=?", (prefix,)


class DocSource:
    """Documents selected by a WHERE clause. They are streamed from the database, never all held in memory."""

    def __init__(self, conn, where, params):
        self.conn, self.where, self.params = conn, where, tuple(params)

    def count(self):
        return self.conn.execute(f"SELECT COUNT(*) FROM docs WHERE {self.where}", self.params).fetchone()[0]

    def rows(self):
        for r in self.conn.execute(f"SELECT * FROM docs WHERE {self.where} ORDER BY id", self.params):
            yield dict(r)

    def attention(self):
        for r in self.conn.execute(f"SELECT * FROM docs WHERE ({self.where}) AND status IN (?, ?) ORDER BY id", self.params + (UNASSIGNED, UNRECOGNISED)):
            yield dict(r)

    def early(self):
        return iter(())


class MonthSource(DocSource):
    """One month's record. Besides its own documents it lists, after them, the documents that ARRIVED EARLY: future-dated invoices that
    waited in an earlier month's folder and were moved into this month's folder. They are shown (light blue) but are not counted
    anywhere else: the document stays in the record of the month it arrived."""

    def __init__(self, conn, year, month, basis="arrival"):
        super().__init__(conn, *month_clause(year, month, basis))
        self.target = f"{year:04d}-{month:02d}"

    def early(self):
        for r in self.conn.execute("SELECT * FROM docs WHERE month_case='moved' AND month_target=? ORDER BY id", (self.target,)):
            yield dict(r)

    def count(self):
        return super().count() + self.conn.execute("SELECT COUNT(*) FROM docs WHERE month_case='moved' AND month_target=?", (self.target,)).fetchone()[0]


class ListSource:
    """The same interface over a plain list (small cases and tests)."""

    def __init__(self, rows):
        self._rows = list(rows)

    def count(self):
        return len(self._rows)

    def rows(self):
        return iter(self._rows)

    def attention(self):
        return (d for d in self._rows if d["status"] in (UNASSIGNED, UNRECOGNISED))

    def early(self):
        return iter(())


def docs_between(conn, start, end, basis="invoice"):
    return list(DocSource(conn, *between_clause(start, end, basis)).rows())


def docs_in_month(conn, year, month, basis="arrival"):
    """Documents of a month: by the e-mail's arrival (default) or by the date printed on the document."""
    return list(DocSource(conn, *month_clause(year, month, basis)).rows())


def corrections_between(conn, start, end):
    """The corrections made through To_Fix between two dates (inclusive)."""
    return [dict(r) for r in conn.execute(
        "SELECT * FROM corrections WHERE substr(at,1,10) >= ? AND substr(at,1,10) <= ? ORDER BY id", (str(start)[:10], str(end)[:10]))]


def add_correction(conn, ref, field, old, new, by):
    conn.execute("INSERT INTO corrections(ref, field, old, new, by, at) VALUES(?,?,?,?,?,?)",
                 (ref, field, str(old), str(new), by, now()))


# ---------------------------------------------------------------------------------- runs / windows

def last_success(conn, mailbox):
    """End of the last successful window for this mailbox (datetime) or None."""
    row = conn.execute("SELECT window_to FROM mailbox_runs WHERE mailbox=? AND ok=1 ORDER BY id DESC LIMIT 1",
                       (mailbox,)).fetchone()
    return datetime.datetime.fromisoformat(row["window_to"]) if row else None


def record_mailbox_run(conn, mailbox, start, end, ok):
    conn.execute("INSERT INTO mailbox_runs(mailbox, window_from, window_to, ok, finished) VALUES(?,?,?,?,?)",
                 (mailbox, start.isoformat(), end.isoformat(), 1 if ok else 0, now()))


def record_run(conn, started, mode, summary_text):
    conn.execute("INSERT INTO runs(started, finished, mode, summary) VALUES(?,?,?,?)",
                 (started, now(), mode, summary_text))


def counts_by_status(conn):
    return {r["status"]: r["n"] for r in conn.execute("SELECT status, COUNT(*) n FROM docs GROUP BY status")}


# ---------------------------------------------------------------------------------- backup and restore

BACKUP_PREFIX, BACKUP_SUFFIX = "bot_", ".db.gz"


def _backup_date(name):
    for suffix in (".db.gz", ".db"):                              # older versions made plain .db backups
        if name.startswith(BACKUP_PREFIX) and name.endswith(suffix):
            try:
                return datetime.datetime.strptime(name[len(BACKUP_PREFIX):len(BACKUP_PREFIX) + 8], "%Y%m%d").date()
            except ValueError:
                return None
    return None


def list_backups(folder):
    """[(path, date, size)] newest first, compressed and plain."""
    out = []
    if os.path.isdir(folder):
        for name in os.listdir(folder):
            d = _backup_date(name)
            if d:
                path = os.path.join(folder, name)
                out.append((path, d, os.path.getsize(path)))
    return sorted(out, key=lambda x: x[1], reverse=True)


def backup(conn, folder, keep_days=14):
    """One COMPRESSED backup per day (about 4 times smaller than the database). It is checked before it is kept:
    SQLite's integrity check on the copy, and a full read-back of the compressed file. Backups older than keep_days are removed."""
    os.makedirs(folder, exist_ok=True)
    cutoff = datetime.date.today() - datetime.timedelta(days=keep_days)
    for path, d, _size in list_backups(folder):                       # clean up on every run, not only when a backup is made
        if d < cutoff:
            try:
                os.remove(path)
            except OSError:
                pass
    today = datetime.date.today().strftime("%Y%m%d")
    target = os.path.join(folder, f"{BACKUP_PREFIX}{today}{BACKUP_SUFFIX}")
    if os.path.exists(target) or os.path.exists(os.path.join(folder, f"{BACKUP_PREFIX}{today}.db")):
        return None
    raw, part = target + ".tmp", target + ".part"
    try:
        dest = sqlite3.connect(raw)
        with dest:
            conn.backup(dest)
        dest.close()
        check = sqlite3.connect(raw)
        verdict = check.execute("PRAGMA integrity_check").fetchone()[0]
        check.close()
        if verdict != "ok":
            raise RuntimeError(f"the copy failed SQLite's integrity check ({verdict})")
        with open(raw, "rb") as src, gzip.open(part, "wb", compresslevel=6) as gz:
            shutil.copyfileobj(src, gz)
        size_back = 0
        with gzip.open(part, "rb") as gz:                             # reading to the end also checks the gzip checksum
            for chunk in iter(lambda: gz.read(1 << 20), b""):
                size_back += len(chunk)
        if size_back != os.path.getsize(raw):
            raise RuntimeError("the compressed backup does not read back to the same size")
        os.replace(part, target)
    finally:
        for leftover in (raw, part):
            if os.path.exists(leftover):
                try:
                    os.remove(leftover)
                except OSError:
                    pass
    return target


def restore(backup_file, db_path):
    """Put a backup back as the live database. The current database is kept next to it as bot.db.before_restore_<time>.
    The bot must not be running. Returns the path of the kept old database (or '')."""
    raw = db_path + ".restoring"
    try:
        if backup_file.endswith(".gz"):
            with gzip.open(backup_file, "rb") as gz, open(raw, "wb") as out:
                shutil.copyfileobj(gz, out)
        else:
            shutil.copyfile(backup_file, raw)
        check = sqlite3.connect(raw)
        verdict = check.execute("PRAGMA integrity_check").fetchone()[0]
        check.close()
        if verdict != "ok":
            raise RuntimeError(f"this backup is damaged ({verdict}); nothing was changed")
        kept = ""
        if os.path.exists(db_path):
            kept = f"{db_path}.before_restore_{datetime.datetime.now():%Y%m%d_%H%M%S}"
            os.replace(db_path, kept)
        for ext in ("-wal", "-shm"):
            if os.path.exists(db_path + ext):
                os.remove(db_path + ext)
        os.replace(raw, db_path)
        return kept
    finally:
        if os.path.exists(raw):
            os.remove(raw)


# ------------------------------------------------------------------ sites learned from staff fixes
def learn(conn, kind, supplier, key, site, company_folder="", example="", by="staff"):
    """Remember that this supplier's deliver-to number / address (kind 'ship-to') or Fleet Card merchant number (kind 'merchant')
    means this site. A later choice for the same key replaces the earlier one."""
    if not key or not site:
        return
    conn.execute("INSERT INTO learned(kind, supplier, key, site, company_folder, example, by, at) VALUES(?,?,?,?,?,?,?,?) "
                 "ON CONFLICT(kind, supplier, key) DO UPDATE SET site=excluded.site, company_folder=excluded.company_folder, "
                 "example=excluded.example, by=excluded.by, at=excluded.at",
                 (kind, supplier or "", key, site, company_folder or "", (example or "")[:200], by, now()))


def learned_site(conn, kind, supplier, keys):
    """The learned (site, company folder, learned on) for the first key that is known, else None."""
    for key in keys:
        row = conn.execute("SELECT site, company_folder, at FROM learned WHERE kind=? AND supplier=? AND key=?",
                           (kind, supplier or "", key)).fetchone()
        if row:
            return row[0], row[1], row[2]
    return None


def learned_all(conn):
    return [dict(zip(("id", "kind", "supplier", "key", "site", "company_folder", "example", "by", "at"), r))
            for r in conn.execute("SELECT id, kind, supplier, key, site, company_folder, example, by, at FROM learned ORDER BY supplier, site")]


def forget(conn, learned_id):
    conn.execute("DELETE FROM learned WHERE id=?", (learned_id,))


def get_meta(conn, key, default=""):
    row = conn.execute("SELECT value FROM meta WHERE key=?", (key,)).fetchone()
    return row[0] if row else default


def set_meta(conn, key, value):
    conn.execute("INSERT INTO meta(key, value) VALUES(?, ?) ON CONFLICT(key) DO UPDATE SET value=excluded.value", (key, value))


def run_markers(conn):
    """{ref: 'LAST E-MAIL OF RUN ...'} for the last e-mail of each mailbox in the latest run that downloaded documents."""
    import json
    try:
        return json.loads(get_meta(conn, "run_markers", "{}") or "{}")
    except ValueError:
        return {}
