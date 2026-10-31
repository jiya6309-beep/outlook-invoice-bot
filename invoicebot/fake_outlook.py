"""A simulated Outlook with the same methods as ClassicOutlook.

Used by the tests, the benchmark and the demo (python run_bot.py --demo), so the whole bot can be
exercised on any PC without Outlook. It also records every change made to an e-mail, which is how the
tests check the flagging rules."""
import itertools
import os

from .outlook_io import AttInfo, MailInfo, OL_FLAG_COMPLETE, OL_FLAG_MARKED, OL_FLAG_NONE, _n
from . import settings

_ids = itertools.count(1)


class FakeMail:
    def __init__(self, subject, received, attachments=None, sender_name="", sender_email="", to_name="", body="",
                 unread=True, flag=OL_FLAG_NONE, categories=None):
        self.entry_id = f"FAKE{next(_ids):08d}"
        self.subject, self.received = subject, received
        self.sender_name, self.sender_email, self.to_name, self.body = sender_name, sender_email, to_name, body
        self.unread, self.flag, self.categories = unread, flag, list(categories or [])
        # attachments: list of dict(name=, path=  or  data=, hidden=False, kind=1)
        self.attachments = attachments or []


class FakeOutlook:
    def __init__(self):
        self.boxes = {}                 # mailbox name -> [FakeMail]
        self.drafts, self.sent = [], []
        self.read_only = False          # True = every change raises (simulates missing edit permission)
        self.save_failures = set()      # (entry_id, attachment name) that fail to save
        self.scan_calls = 0

    def add_mailbox(self, name):
        self.boxes.setdefault(name, [])
        return name

    def add(self, mailbox, mail, folder=""):
        mail.folder = folder                             # '' = the Inbox, otherwise a sub-folder of the Inbox
        self.boxes.setdefault(mailbox, []).append(mail)
        return mail

    def list_mailboxes(self):
        return list(self.boxes)

    def resolve_store(self, names, exclude=()):
        """Like the real Outlook: an exact name first, then a lenient one (one name inside the other); never an excluded mailbox."""
        boxes = [b for b in self.boxes if _n(b) not in exclude]
        for want in names:
            for box in boxes:
                if _n(box) == _n(want):
                    return box
        for want in names:
            w = _n(want)
            for box in boxes:
                if w and _n(box) and (w in _n(box) or _n(box) in w):
                    return box
        raise LookupError("mailbox not found in Outlook: " + " | ".join(names))

    def _mail(self, info):
        return next(m for m in self.boxes[info.store] if m.entry_id == info.entry_id)

    def scan(self, store, label, start, end, limit, subfolders=False):
        self.scan_calls += 1
        mails = sorted((m for m in self.boxes[store] if subfolders or not getattr(m, "folder", "")),
                       key=lambda m: m.received, reverse=True)
        out, truncated = [], False
        for n, m in enumerate(mails, start=1):
            if n > limit:
                truncated = True
                break
            if m.received < start:
                break
            if m.received <= end:
                out.append(MailInfo(mailbox=label, store=store, entry_id=m.entry_id, subject=m.subject,
                                    received=m.received, sender_name=m.sender_name, sender_email=m.sender_email,
                                    to_name=m.to_name, folder=getattr(m, "folder", ""),
                                    done=m.flag == OL_FLAG_COMPLETE or "Processed by bot" in m.categories or (m.flag == OL_FLAG_MARKED and settings.run_bool("Skip e-mails already flagged in Outlook", True))))
        return out, truncated

    def attachments(self, info):
        out = []
        for i, a in enumerate(self._mail(info).attachments, start=1):
            size = a.get("size")
            if size is None:
                size = os.path.getsize(a["path"]) if a.get("path") else len(a.get("data", b""))
            out.append(AttInfo(index=i, name=a["name"], size=size, hidden=a.get("hidden", False), kind=a.get("kind", 1)))
        return out

    def save_attachment(self, info, index, path):
        a = self._mail(info).attachments[index - 1]
        if (info.entry_id, a["name"]) in self.save_failures:
            raise OSError("simulated save failure")
        if a.get("path"):
            with open(a["path"], "rb") as src, open(path, "wb") as dst:
                dst.write(src.read())
        else:
            with open(path, "wb") as dst:
                dst.write(a["data"])

    def body_start(self, info, chars=600):
        return self._mail(info).body[:chars]

    def _check(self):
        if self.read_only:
            raise PermissionError("simulated: no edit permission on this mailbox")

    def mark_done(self, info):
        self._check()
        m = self._mail(info)
        m.unread = False
        if m.flag == OL_FLAG_NONE:
            m.flag = OL_FLAG_COMPLETE

    def mark_unread(self, info):
        self._check()
        self._mail(info).unread = True

    def flag_red(self, info):
        self._check()
        m = self._mail(info)
        if m.flag == OL_FLAG_NONE:
            m.flag = OL_FLAG_MARKED

    def add_category(self, info, name, color="red"):
        self._check()
        m = self._mail(info)
        if name not in m.categories:
            m.categories.append(name)

    def create_mail(self, to, subject, body, send=False):
        (self.sent if send else self.drafts).append({"to": to, "subject": subject, "body": body})
