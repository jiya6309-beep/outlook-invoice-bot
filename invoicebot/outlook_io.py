"""The only file that talks to Outlook.

    ClassicOutlook   the real thing: classic desktop Outlook on Windows, through COM (pywin32).
                     "New Outlook" and Outlook on the web do NOT offer COM and cannot be used.
    (fake_outlook.py has a simulated Outlook with the same methods, used by the tests and the demo.)

The pipeline never touches COM directly, so everything else can be tested without Outlook.
Nothing here moves or deletes an e-mail. The only changes ever made are: mark as read, set the
follow-up flag (only if the e-mail has no flag yet), and add a category.
"""
import datetime
from dataclasses import dataclass

import re

OL_MAIL_ITEM = 43
OL_FLAG_NONE, OL_FLAG_COMPLETE, OL_FLAG_MARKED = 0, 1, 2
OL_MARK_COMPLETE = 5                      # OlMarkInterval.olMarkComplete (for MarkAsTask)
OL_BY_VALUE = 1
CATEGORY_COLORS = {"red": 1, "orange": 2, "yellow": 4, "green": 5, "blue": 8}
PR_ATTACHMENT_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"


@dataclass
class MailInfo:
    mailbox: str            # label used inside the bot (the company name)
    store: str              # the real mailbox name in Outlook
    entry_id: str
    subject: str
    received: datetime.datetime
    sender_name: str = ""
    sender_email: str = ""
    to_name: str = ""
    folder: str = ""        # '' = the Inbox; otherwise the sub-folder of the Inbox it was found in (fuel mailbox)
    done: bool = False      # already ticked (flag complete) or given the 'Processed by bot' category: handled before


@dataclass
class AttInfo:
    index: int
    name: str
    size: int
    hidden: bool = False    # inline pictures (signature logos) are hidden attachments
    kind: int = OL_BY_VALUE


def _n(text):
    return re.sub(r"[^a-z0-9]", "", str(text).lower())


def _naive(dt):
    return datetime.datetime(dt.year, dt.month, dt.day, dt.hour, dt.minute, dt.second)


class ClassicOutlook:
    """Gateway to classic desktop Outlook (must be open and signed in on this PC)."""

    def __init__(self):
        import win32com.client                      # pywin32; only needed on the bot PC
        # Dispatch() silently STARTS a hidden Outlook when none is open, which made a closed Outlook look "found".
        # So first ask for an Outlook that is already running, and stop with a clear message if there is none.
        get_running = getattr(win32com.client, "GetActiveObject", None)
        if get_running is not None:
            try:
                self.app = get_running("Outlook.Application")
            except Exception:
                raise RuntimeError("Outlook is not open. Open classic Outlook, wait until your mailboxes show, then try again")
        else:
            self.app = win32com.client.Dispatch("Outlook.Application")
        self.ns = self.app.GetNamespace("MAPI")
        self._inbox_cache = {}

    # ------------------------------------------------------------------ finding mailboxes
    def list_mailboxes(self):
        return [str(f.Name) for f in self.ns.Folders]

    def _find_store(self, names, exclude=()):
        """exclude = squashed names of mailboxes that must never be picked (Never read), not even by the lenient match."""
        stores = [f for f in self.ns.Folders if _n(f.Name) not in exclude]
        for want in names:                                           # exact match first
            for f in stores:
                if _n(f.Name) == _n(want):
                    return f
        for want in names:                                           # then a lenient one
            w = _n(want)
            for f in stores:
                n = _n(f.Name)
                if w and n and (w in n or n in w):
                    return f
        raise LookupError("mailbox not found in Outlook: " + " | ".join(names))

    def _inbox(self, store):
        if store not in self._inbox_cache and str(store).startswith("shared:"):
            inbox = self._shared_inbox(str(store)[7:])
            if inbox is None:
                raise LookupError(f"shared mailbox {store[7:]} cannot be opened")
            self._inbox_cache[store] = inbox
        if store not in self._inbox_cache:
            folder = self._find_store([store])
            inbox = next((s for s in folder.Folders if _n(s.Name) == "inbox"), None)
            if inbox is None:
                raise LookupError(f"no Inbox folder in mailbox {store}")
            self._inbox_cache[store] = inbox
        return self._inbox_cache[store]

    def resolve_store(self, names, exclude=()):
        """The real Outlook name for a mailbox given alternatives, or raises LookupError.
        A shared mailbox that Outlook does not list as its own mailbox (no automapping) is opened by its e-mail address through
        the signed-in account (for example pifuelinvoice@ through office@): its key is then 'shared:<address>'."""
        try:
            return str(self._find_store(names, exclude).Name)
        except LookupError:
            for want in names:
                if "@" in want and _n(want) not in exclude:
                    inbox = self._shared_inbox(want)
                    if inbox is not None:
                        key = "shared:" + want.strip().lower()
                        self._inbox_cache[key] = inbox
                        return key
            raise

    def _shared_inbox(self, address):
        """The Inbox of a shared mailbox opened by address, or None when this account has no access to it."""
        try:
            rec = self.ns.CreateRecipient(address.strip())
            if not rec.Resolve():
                return None
            return self.ns.GetSharedDefaultFolder(rec, 6)          # 6 = olFolderInbox
        except Exception:
            return None

    # ------------------------------------------------------------------ reading
    def scan(self, store, label, start, end, limit, subfolders=False):
        """E-mails received between start and end, newest first. Returns (list, truncated).
        subfolders=True: also every folder UNDER the Inbox (P I Fuel Invoice: Outlook rules move each company's fuel invoices
        into its own sub-folder - Colyton, Gopinath, P1 ... - and the rest stays in the Inbox)."""
        inbox = self._inbox(store)
        folders = [(inbox, "")]
        if subfolders:
            stack = [(inbox, "")]
            while stack:
                parent, path = stack.pop()
                try:
                    for sub in parent.Folders:
                        name = (path + "\\" if path else "") + str(sub.Name)
                        folders.append((sub, name))
                        stack.append((sub, name))
                except Exception:
                    pass
        out, truncated = [], False
        for folder, fname in folders:
            got, cut = self._scan_folder(folder, fname, store, label, start, end, limit)
            out += got
            truncated = truncated or cut
        out.sort(key=lambda m: m.received, reverse=True)
        return out, truncated

    def _scan_folder(self, folder, fname, store, label, start, end, limit):
        items = folder.Items
        items.Sort("[ReceivedTime]", True)
        out, count = [], 0
        item = items.GetFirst()
        truncated = False
        while item is not None:
            count += 1
            if count > limit:
                truncated = True
                break
            try:
                if item.Class == OL_MAIL_ITEM:
                    received = _naive(item.ReceivedTime)
                    if received < start:
                        break                                         # sorted newest first: we are done
                    if received <= end:
                        email = ""
                        try:
                            if str(item.SenderEmailType).upper() == "SMTP":
                                email = str(item.SenderEmailAddress or "")
                        except Exception:
                            pass
                        out.append(MailInfo(mailbox=label, store=store, entry_id=str(item.EntryID),
                                            subject=str(item.Subject or ""), received=received,
                                            sender_name=str(item.SenderName or ""), sender_email=email,
                                            to_name=str(item.To or ""), folder=fname, done=self._is_done(item)))
            except Exception:
                pass                                                  # one unreadable item must not stop the run
            item = items.GetNext()
        return out, truncated

    @staticmethod
    def _is_done(item):
        """True when the e-mail carries the bot's tick (flag complete) or its fallback category: it was handled before."""
        try:
            fs = int(item.FlagStatus)
            if fs == OL_FLAG_COMPLETE:
                return True
            if fs == OL_FLAG_MARKED:                                  # red flag (the bot's own, or a person's)
                from . import settings
                if settings.run_bool("Skip e-mails already flagged in Outlook", True):
                    return True
        except Exception:
            pass
        try:
            from . import settings
            name = (settings.run_value("Tick fallback category", "Processed by bot") or "").strip().lower()
            if name and name != "none":
                return name in [c.strip().lower() for c in re.split(r"[;,]", str(item.Categories or "")) if c.strip()]
        except Exception:
            pass
        return False

    def _item(self, info):
        return self.ns.GetItemFromID(info.entry_id, self._inbox(info.store).StoreID)

    def attachments(self, info):
        item = self._item(info)
        out = []
        for i in range(1, int(item.Attachments.Count) + 1):
            a = item.Attachments.Item(i)
            hidden = False
            try:
                hidden = bool(a.PropertyAccessor.GetProperty(PR_ATTACHMENT_HIDDEN))
            except Exception:
                pass
            out.append(AttInfo(index=i, name=str(a.FileName or ""), size=int(a.Size or 0), hidden=hidden,
                               kind=int(a.Type)))
        return out

    def save_attachment(self, info, index, path):
        self._item(info).Attachments.Item(index).SaveAsFile(path)

    def body_start(self, info, chars=600):
        return str(self._item(info).Body or "")[:chars]

    # ------------------------------------------------------------------ the only changes made to e-mails
    # Every change is CHECKED: the e-mail is fetched again from Outlook after saving and the value is compared, so a change that
    # Outlook accepted but did not keep is noticed. When something cannot be done, the error says exactly what Outlook answered.

    def _clear_flag(self, item):
        try:
            item.ClearTaskFlag()
        except Exception:
            item.FlagStatus = OL_FLAG_NONE

    def mark_unread(self, info):
        """Leave the e-mail UNREAD (forwards and replies: a person should look at them). Never ticks. Checked by reading it back."""
        item = self._item(info)
        item.UnRead = True
        item.Save()
        if not bool(self._item(info).UnRead):
            raise RuntimeError("Outlook accepted 'unread' but the e-mail came back as read")

    def mark_done(self, info):
        """Read + flag complete (the tick). An existing follow-up flag is never overwritten.
        The tick is tried three ways (Outlook versions and mailbox types differ); if none works, the e-mail gets the 'fallback
        category' (setting 'Tick fallback category') so staff can still see which e-mails the bot handled."""
        item = self._item(info)
        item.UnRead = False
        wants_tick = int(item.FlagStatus) == OL_FLAG_NONE
        problems = []
        if wants_tick:
            def by_flag(it):
                it.FlagStatus = OL_FLAG_COMPLETE

            def by_task(it):
                it.MarkAsTask(OL_MARK_COMPLETE)

            def by_flag_and_date(it):
                it.FlagStatus = OL_FLAG_COMPLETE
                it.TaskCompletedDate = datetime.date.today()

            for label, action in (("FlagStatus", by_flag), ("MarkAsTask", by_task), ("FlagStatus+date", by_flag_and_date)):
                try:
                    action(item)
                    item.Save()
                    seen = int(self._item(info).FlagStatus)                         # a fresh copy: what Outlook really kept
                    if seen == OL_FLAG_COMPLETE:
                        return
                    problems.append(f"{label}: Outlook accepted it, but the flag was {seen} when read back")
                except Exception as e:
                    problems.append(f"{label}: {e}")
                item = self._item(info)
        try:
            item = self._item(info)
            item.UnRead = False
            item.Save()                                                             # at least the read status must stick
        except Exception as e:
            problems.append(f"read status: {e}")
        if problems:
            from . import settings
            fallback = settings.run_value("Tick fallback category", "Processed by bot")
            note = ""
            if fallback and fallback.strip().lower() != "none":
                try:
                    self.add_category(info, fallback, "green")
                    note = f"; the category '{fallback}' was added instead"
                except Exception as e:
                    note = f"; the category '{fallback}' could not be added either ({e})"
            raise RuntimeError("the tick could not be set (" + " | ".join(problems) + ")" + note)

    def flag_red(self, info):
        item = self._item(info)
        if int(item.FlagStatus) != OL_FLAG_NONE:
            return
        problems = []
        for label, action in (("FlagStatus", lambda it: setattr(it, "FlagStatus", OL_FLAG_MARKED)), ("MarkAsTask", lambda it: it.MarkAsTask(0))):
            try:
                action(item)
                item.Save()
                seen = int(self._item(info).FlagStatus)
                if seen == OL_FLAG_MARKED:
                    return
                problems.append(f"{label}: Outlook accepted it, but the flag was {seen} when read back")
            except Exception as e:
                problems.append(f"{label}: {e}")
            item = self._item(info)
        raise RuntimeError("the red flag could not be set (" + " | ".join(problems) + ")")

    def add_category(self, info, name, color="red"):
        try:                                                           # make sure the category exists in the colour list
            self.ns.Categories.Item(name)
        except Exception:
            try:
                self.ns.Categories.Add(name, CATEGORY_COLORS.get(color, 1))
            except Exception:
                pass                                                   # it can still be given to the e-mail by name
        item = self._item(info)
        current = [c.strip() for c in re.split(r"[;,]", str(item.Categories or "")) if c.strip()]
        if name not in current:
            item.Categories = ", ".join(current + [name])
            item.Save()
            seen = [c.strip() for c in re.split(r"[;,]", str(self._item(info).Categories or "")) if c.strip()]
            if name not in seen:
                raise RuntimeError(f"the category '{name}' was accepted but is not on the e-mail when read back (categories: '{', '.join(seen)}')")

    # ------------------------------------------------------------------ the self-test
    def diagnose_marking(self, mailbox_text, subject_text, days=14, now=None):
        """Tries every kind of marking on ONE e-mail (the newest whose subject contains subject_text), reads each back from Outlook,
        and puts the e-mail back as it was. Returns (lines, ok)."""
        lines, results = [], {}

        def say(text=""):
            lines.append(text)

        def step(name, action, check):
            try:
                action()
                good = bool(check())
                results[name] = good
                say(f"  {'OK  ' if good else 'FAIL'}  {name}" + ("" if good else "   (Outlook accepted it, but it was not there when read back)"))
            except Exception as e:
                results[name] = False
                say(f"  FAIL  {name}   {type(e).__name__}: {str(e)[:220]}")

        for what in ("Version",):
            try:
                say(f"Outlook version: {getattr(self.app, what)}")
            except Exception as e:
                say(f"Outlook version: unknown ({e})")
        store = next((f for f in self.ns.Folders if _n(mailbox_text) and _n(mailbox_text) in _n(f.Name)), None)
        if store is None:
            say(f"No mailbox in this Outlook contains '{mailbox_text}'. Mailboxes here: " + ", ".join(str(f.Name) for f in self.ns.Folders))
            return lines, False
        say(f"Mailbox: {store.Name}")
        try:
            say(f"Mailbox type: {store.Store.ExchangeStoreType} (0 = your own Exchange mailbox, 1 = shared/delegate, 3 = not Exchange)   connection mode: {self.ns.ExchangeConnectionMode}")
        except Exception:
            pass
        inbox = next((s for s in store.Folders if _n(s.Name) == "inbox"), None)
        if inbox is None:
            say("This mailbox has no Inbox folder.")
            return lines, False
        since = (now or datetime.datetime.now()) - datetime.timedelta(days=days)
        items = inbox.Items
        items.Sort("[ReceivedTime]", True)
        target, item = None, items.GetFirst()
        while item is not None:
            try:
                if item.Class == OL_MAIL_ITEM:
                    if _naive(item.ReceivedTime) < since:
                        break
                    if subject_text.lower() in str(item.Subject or "").lower():
                        target = item
                        break
            except Exception:
                pass
            item = items.GetNext()
        if target is None:
            say(f"No e-mail with '{subject_text}' in the subject was found in the last {days} days. Send yourself one and try again.")
            return lines, False
        info = MailInfo(mailbox="TEST", store=str(store.Name), entry_id=str(target.EntryID), subject=str(target.Subject or ""),
                        received=_naive(target.ReceivedTime))
        say(f"Test e-mail: '{info.subject[:70]}'")
        item = self._item(info)
        original = dict(unread=bool(item.UnRead), flag=int(item.FlagStatus), cats=str(item.Categories or ""))
        if original["flag"] != OL_FLAG_NONE:
            say("This e-mail already has a follow-up flag, and the test would have to change it. Pick an e-mail without a flag.")
            return lines, False
        say("")
        say("Trying each kind of change, reading it back from Outlook each time:")

        def fresh():
            return self._item(info)

        def mark(attr, value):
            it = fresh()
            setattr(it, attr, value)
            it.Save()

        step("mark as read", lambda: mark("UnRead", False), lambda: not bool(fresh().UnRead))
        step("tick by setting FlagStatus", lambda: mark("FlagStatus", OL_FLAG_COMPLETE), lambda: int(fresh().FlagStatus) == OL_FLAG_COMPLETE)
        self._reset_flag(info)
        step("tick by MarkAsTask(complete)", lambda: (lambda it: (it.MarkAsTask(OL_MARK_COMPLETE), it.Save()))(fresh()), lambda: int(fresh().FlagStatus) == OL_FLAG_COMPLETE)
        self._reset_flag(info)
        step("tick by FlagStatus + completed date", lambda: (lambda it: (setattr(it, "FlagStatus", OL_FLAG_COMPLETE), setattr(it, "TaskCompletedDate", datetime.date.today()), it.Save()))(fresh()),
             lambda: int(fresh().FlagStatus) == OL_FLAG_COMPLETE)
        self._reset_flag(info)
        step("red flag by setting FlagStatus", lambda: mark("FlagStatus", OL_FLAG_MARKED), lambda: int(fresh().FlagStatus) == OL_FLAG_MARKED)
        self._reset_flag(info)
        test_category = "Bot test category"

        def add_cat():
            try:
                self.ns.Categories.Add(test_category, 1)
            except Exception:
                pass
            it = fresh()
            it.Categories = (original["cats"] + ", " if original["cats"] else "") + test_category
            it.Save()

        step("category added to the e-mail", add_cat, lambda: test_category in str(fresh().Categories or ""))
        say("")
        try:                                                                       # put everything back as it was
            it = fresh()
            it.Categories = original["cats"]
            it.UnRead = original["unread"]
            self._clear_flag(it)
            it.Save()
            back = fresh()
            restored = (bool(back.UnRead) == original["unread"] and int(back.FlagStatus) == OL_FLAG_NONE and str(back.Categories or "") == original["cats"])
            say("The e-mail was put back as it was." if restored else "WARNING: the e-mail could not be put back exactly; check its read status, flag and categories.")
        except Exception as e:
            say(f"WARNING: the e-mail could not be put back ({e}); check its read status, flag and categories.")
        try:
            self.ns.Categories.Remove(test_category)
        except Exception:
            pass
        ways = [n for n in ("tick by setting FlagStatus", "tick by MarkAsTask(complete)", "tick by FlagStatus + completed date") if results.get(n)]
        say("RESULT")
        say(f"  The bot can mark as read: {'yes' if results.get('mark as read') else 'NO'}")
        say(f"  The bot can set the tick: {'yes, by ' + ways[0].replace('tick by ', '') if ways else 'NO - it will add the fallback category instead'}")
        say(f"  The bot can set the red flag: {'yes' if results.get('red flag by setting FlagStatus') else 'NO'}")
        say(f"  The bot can add a category: {'yes' if results.get('category added to the e-mail') else 'NO'}")
        return lines, bool(results) and all(results.get(n) for n in ("mark as read", "category added to the e-mail")) and bool(ways)

    def _reset_flag(self, info):
        try:
            it = self._item(info)
            self._clear_flag(it)
            it.Save()
        except Exception:
            pass

    def create_mail(self, to, subject, body, send=False):
        mail = self.app.CreateItem(0)
        mail.To, mail.Subject, mail.Body = to, subject, body
        mail.Send() if send else mail.Save()
