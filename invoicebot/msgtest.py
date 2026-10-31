"""Run the bot on a FOLDER OF SAVED .msg E-MAILS instead of Outlook (for testing with real e-mails).

Each .msg becomes an e-mail in a simulated Outlook. Nothing in your real Outlook is touched. The company mailbox is chosen
from the e-mail's recipient name (for example "Pearl Energy Colyton"); other e-mails use the mailbox you give."""
import datetime
import os

from . import companies, msgfile
from .fake_outlook import FakeMail, FakeOutlook


def build(folder, default_mailbox=None):
    """Returns (gateway, from_date, to_date, count, problems)."""
    files = sorted(os.path.join(folder, f) for f in os.listdir(folder) if f.lower().endswith(".msg"))
    gw = FakeOutlook()
    for row in companies.all_active():
        gw.add_mailbox(companies.mailbox_names(row)[0])
    fallback = (companies.by_name(default_mailbox) or companies.for_mailbox(default_mailbox)) if default_mailbox else None
    fallback = fallback or companies.all_active()[0]
    problems, dates, count = [], [], 0
    for path in files:
        try:
            m = msgfile.load(path)
        except (msgfile.MsgError, OSError) as e:
            problems.append((os.path.basename(path), str(e)))
            continue
        received = (m["received"] or datetime.datetime.fromtimestamp(os.path.getmtime(path))).replace(microsecond=0)
        row = None
        for part in m["to"].replace(",", ";").split(";"):
            row = companies.for_mailbox(part.strip()) if part.strip() else None
            if row:
                break
        row = row or fallback
        atts = [{"name": a["filename"] or "attachment", "data": a["data"] or b"", "hidden": a["hidden"],
                 "kind": 1 if a["method"] == 1 else 5} for a in m["attachments"]]
        mail = FakeMail(m["subject"], received, attachments=atts, sender_name=m["sender_name"],
                        sender_email=m["sender_email"], to_name=m["to"], body=m["body"])
        mail.expect = os.path.basename(path)[:60]
        gw.add(companies.mailbox_names(row)[0], mail)
        dates.append(received.date())
        count += 1
    if not dates:
        return gw, None, None, 0, problems
    return gw, min(dates), max(dates), count, problems
