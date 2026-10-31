"""Shared helpers for the tests: an isolated BOT_HOME per test and small builders."""
import datetime
import os
import shutil
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)
DATA = os.path.join(ROOT, "tests", "data")

from openpyxl import load_workbook  # noqa: E402

from invoicebot import settings  # noqa: E402
from invoicebot.fake_outlook import FakeMail, FakeOutlook  # noqa: E402

ONE = "One@pearlenergy.com.au"
COLYTON = "pearl.colyton@pearlenergy.com.au"
NOW = datetime.datetime(2026, 10, 3, 14, 0, 0)
# 5.15.0: LIVE creates every missing folder like TEST by default. Tests of the older, careful option switch it back on with this.
CAREFUL_LIVE = {"Live_mode_creates_missing_folders_like_test_mode": "No", "Create_site_folder_in_live_mode": "No"}          # a fixed 'now' so the tests do not depend on the date


class Isolated:
    """Mixin for unittest.TestCase: every test gets its own empty BOT_HOME and default settings."""

    def setUp(self):
        self.home = tempfile.mkdtemp(prefix="botTest_")
        self._old = os.environ.get("BOT_HOME")
        os.environ["BOT_HOME"] = self.home
        self._env = {k: os.environ.get(k) for k in ("HOME", "USERPROFILE", "OneDrive", "OneDriveCommercial", "OneDriveConsumer")}
        self.profile = os.path.join(self.home, "profile")                      # the bot's 'user folder' is inside the test
        os.makedirs(self.profile)
        os.environ["HOME"] = os.environ["USERPROFILE"] = self.profile
        for k in ("OneDrive", "OneDriveCommercial", "OneDriveConsumer"):
            os.environ.pop(k, None)
        from invoicebot import autofind
        autofind.session_root = ""
        settings.SETTINGS_FILE_OVERRIDE = None
        settings.reset()
        settings.get()                                           # creates the default Bot_Settings.xlsx

    def tearDown(self):
        for k, v in self._env.items():
            if v is None:
                os.environ.pop(k, None)
            else:
                os.environ[k] = v
        if self._old is None:
            os.environ.pop("BOT_HOME", None)
        else:
            os.environ["BOT_HOME"] = self._old
        settings.reset()
        shutil.rmtree(self.home, ignore_errors=True)

    # ---- settings edits (what staff do in Excel)
    def set_run(self, **kv):
        wb = load_workbook(settings.settings_file())
        ws = wb["Run"]
        wanted = {k.replace("_", " ").lower(): str(v) for k, v in kv.items()}
        for r in range(2, ws.max_row + 1):
            key = str(ws.cell(r, 1).value or "").lower()
            if key in wanted:
                ws.cell(r, 2).value = wanted.pop(key)
        assert not wanted, f"unknown Run settings: {list(wanted)}"
        wb.save(settings.settings_file())
        wb.close()
        settings.reset()

    def add_rows(self, sheet, rows):
        wb = load_workbook(settings.settings_file())
        for row in rows:
            wb[sheet].append(row)
        wb.save(settings.settings_file())
        wb.close()
        settings.reset()

    def docs_under(self, *parts):
        """files_under without the LIVE Reports / HOLDINGS folders (5.15.0: they sit inside the EMAIL INVOICES folder)."""
        return [f for f in self.files_under(*parts) if not f.startswith(("Reports/", "HOLDINGS/"))]

    def path(self, *parts):
        return os.path.join(self.home, *parts)

    def files_under(self, *parts):
        base = self.path(*parts)
        out = []
        for folder, _dirs, names in os.walk(base):
            for n in names:
                out.append(os.path.relpath(os.path.join(folder, n), base).replace("\\", "/"))
        return sorted(out)


def write(path, data):
    """Write a small file (text or bytes) and close it."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "wb" if isinstance(data, bytes) else "w") as f:
        f.write(data)


def read_bytes(path):
    with open(path, "rb") as f:
        return f.read()


def gateway(*mailboxes):
    """A simulated Outlook where every company mailbox exists (most are empty)."""
    from invoicebot import companies
    gw = FakeOutlook()
    for row in companies.all_active():
        if companies.mailbox_names(row):
            gw.add_mailbox(companies.mailbox_names(row)[0])
    for m in mailboxes:
        gw.add_mailbox(m)
    return gw


def mail(subject, files=(), hours_ago=2, now=NOW, **kw):
    atts = [{"name": os.path.basename(f), "path": f if os.path.isabs(f) else os.path.join(DATA, f)} for f in files]
    return FakeMail(subject, now - datetime.timedelta(hours=hours_ago), attachments=atts, **kw)


def run(gw, **kw):
    from invoicebot.pipeline import Pipeline
    kw.setdefault("workers", 1)
    kw.setdefault("now", NOW)
    kw.setdefault("log", lambda *a, **k: None)
    return Pipeline(gw, **kw).run()


def connect_db(home):
    from invoicebot import db
    return db.connect(os.path.join(home, "data", "bot.db"))
