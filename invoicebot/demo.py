"""Demo: the whole bot against a simulated Outlook with sample e-mails (no Outlook needed)."""
import datetime
import json
import os
import shutil

from . import paths
from .fake_outlook import FakeMail, FakeOutlook


def build(now=None):
    now = now or datetime.datetime.now().replace(microsecond=0)
    folder = os.path.join(paths.project_dir(), "demo_data")
    with open(os.path.join(folder, "emails.json"), encoding="utf-8") as f:
        emails = json.load(f)
    gw = FakeOutlook()
    from . import companies
    for row in companies.all_active():                      # every company mailbox exists (most are empty in the demo)
        gw.add_mailbox(companies.mailbox_names(row)[0])
    for e in emails:
        sender = e.get("sender", "")
        name, _, addr = sender.partition("<")
        mail = FakeMail(e["subject"], now - datetime.timedelta(hours=e["hours_ago"]),
                        attachments=[{"name": a["file"], "path": os.path.join(folder, a["file"])} for a in e["attachments"]],
                        sender_name=name.strip(), sender_email=addr.rstrip(">").strip(), to_name=e.get("to", ""))
        mail.expect = e.get("expect", "")
        gw.add(e["mailbox"], mail)
    return gw


def reset_workspace(name="demo_run"):
    """The demo / .msg test works in its own folder so it can never touch real data."""
    home = os.path.join(paths.project_dir(), name)
    shutil.rmtree(home, ignore_errors=True)
    os.makedirs(home)
    os.environ["BOT_HOME"] = home
    return home


def describe(gw):
    """After a run: what happened to every sample e-mail (read / flag / categories)."""
    names = {0: "-", 1: "TICK", 2: "RED FLAG"}
    rows = []
    for box, mails in gw.boxes.items():
        for m in sorted(mails, key=lambda x: x.received, reverse=True):
            rows.append((m.subject[:52], "unread" if m.unread else "read", names.get(m.flag, str(m.flag)),
                         ", ".join(m.categories) or "-", getattr(m, "expect", "")))
    return rows
