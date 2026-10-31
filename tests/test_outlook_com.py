"""Tests the REAL Outlook adapter (invoicebot/outlook_io.ClassicOutlook) against a mock of Outlook's documented object model.

This cannot prove how Outlook behaves on your PC, but it does prove that the adapter calls the object model correctly
(property names, 1-based collections, Items.Sort/GetFirst/GetNext, categories, flags, attachments) and that the whole bot
works through it."""
import datetime
import os
import sys
import types
import unittest

from tests.helpers import DATA, NOW, Isolated, run



class Attachment:
    def __init__(self, name, path=None, data=b"x", hidden=False, kind=1):
        self.FileName, self._path, self._data, self._hidden, self.Type = name, path, data, hidden, kind
        self.Size = os.path.getsize(path) if path else len(data)
        self.PropertyAccessor = types.SimpleNamespace(GetProperty=self._prop)

    def _prop(self, tag):
        if "7FFE000B" not in tag:
            raise KeyError(tag)
        return self._hidden

    def SaveAsFile(self, target):
        data = self._data
        if self._path:
            with open(self._path, "rb") as src:
                data = src.read()
        with open(target, "wb") as f:
            f.write(data)


class Attachments:
    def __init__(self, items):
        self._items = items
        self.Count = len(items)

    def Item(self, i):
        if not 1 <= i <= len(self._items):
            raise IndexError("COM collections are 1-based")
        return self._items[i - 1]


class MailItem:
    Class = 43
    _n = 0

    def __init__(self, subject, received, attachments=(), sender="", sender_addr="", to="", body="", flag=0):
        MailItem._n += 1
        self.EntryID = f"00000000ABCD{MailItem._n:08d}"
        self.Subject, self.ReceivedTime, self.SenderName, self.To, self.Body = subject, received, sender, to, body
        self.SenderEmailType, self.SenderEmailAddress = "SMTP", sender_addr
        self.UnRead, self.FlagStatus, self.Categories = True, flag, ""
        self.Attachments = Attachments(list(attachments))
        self.saves = 0

    def Save(self):
        self.saves += 1

    def MarkAsTask(self, interval):                       # what Outlook does for olMarkComplete (5)
        if interval == 5:
            self.FlagStatus = 1

    def ClearTaskFlag(self):
        self.FlagStatus = 0


class DropsDirectFlag(MailItem):
    """Outlook accepts FlagStatus = 1 and even shows it, but throws it away when the item is saved; MarkAsTask is kept."""
    @property
    def FlagStatus(self):
        return self._flag

    @FlagStatus.setter
    def FlagStatus(self, value):
        self._flag, self._direct = value, True

    def Save(self):
        self.saves += 1
        if getattr(self, "_direct", False):
            self._flag, self._direct = 0, False

    def MarkAsTask(self, interval):
        if interval == 5:
            self._flag = 1

    def ClearTaskFlag(self):
        self._flag = 0


class IgnoresCategories(MailItem):
    """Accepts a category without error and does not keep it."""
    @property
    def Categories(self):
        return ""

    @Categories.setter
    def Categories(self, value):
        pass


class IgnoresFlagStatus(MailItem):
    """Some mailboxes accept FlagStatus = 1 without error and then show no tick (the flag stays at 0)."""
    @property
    def FlagStatus(self):
        return self._flag

    @FlagStatus.setter
    def FlagStatus(self, value):
        self._flag = getattr(self, "_flag", 0)                 # accepted, but nothing changes

    def MarkAsTask(self, interval):                            # the task route works on this mailbox
        if interval == 5:
            self._flag = 1

    def ClearTaskFlag(self):
        self._flag = 0


class RefusesEverything(MailItem):
    """Every way of setting the tick is refused."""
    @property
    def FlagStatus(self):
        return 0

    @FlagStatus.setter
    def FlagStatus(self, value):
        if getattr(self, "armed", False):
            raise OSError("the property cannot be set on this item")

    def MarkAsTask(self, interval):
        raise OSError("MarkAsTask is not available")

    TaskCompletedDate = property(lambda self: None, lambda self, v: (_ for _ in ()).throw(OSError("TaskCompletedDate cannot be set")))


class OtherItem:
    Class = 26                      # an appointment: must be ignored


class Items:
    def __init__(self, items):
        self._items, self._pos = list(items), 0

    def Sort(self, prop, descending):
        assert prop == "[ReceivedTime]"
        self._items.sort(key=lambda i: getattr(i, "ReceivedTime", datetime.datetime.min), reverse=descending)

    def GetFirst(self):
        self._pos = 0
        return self._items[0] if self._items else None

    def GetNext(self):
        self._pos += 1
        return self._items[self._pos] if self._pos < len(self._items) else None


class Folder:
    def __init__(self, name, items=(), store_id="STORE1"):
        self.Name, self._items, self.StoreID, self.Folders = name, list(items), store_id, []

    @property
    def Items(self):                 # a new collection on every access, like Outlook
        return Items(self._items)


class Store:
    def __init__(self, name, inbox):
        self.Name, self.Folders = name, [Folder("Sent Items", store_id=inbox.StoreID), inbox]


class Categories:
    def __init__(self):
        self.names = {}

    def Item(self, name):
        if name not in self.names:
            raise KeyError(name)
        return self.names[name]

    def Add(self, name, color):
        self.names[name] = color


class Namespace:
    def __init__(self, stores):
        self.Folders, self.Categories = stores, Categories()

    def GetItemFromID(self, entry_id, store_id):
        for s in self.Folders:
            for f in s.Folders:
                if f.StoreID == store_id:
                    for it in f._items:
                        if getattr(it, "EntryID", None) == entry_id:
                            return it
        raise KeyError(entry_id)


class App:
    def __init__(self, stores):
        self.ns, self.created = Namespace(stores), []

    def GetNamespace(self, kind):
        assert kind == "MAPI"
        return self.ns

    def CreateItem(self, kind):
        assert kind == 0
        item = types.SimpleNamespace(To="", Subject="", Body="", saved=False, sent=False)
        item.Save = lambda: setattr(item, "saved", True)
        item.Send = lambda: setattr(item, "sent", True)
        self.created.append(item)
        return item


def install_mock(stores):
    app = App(stores)
    client = types.ModuleType("win32com.client")
    client.Dispatch = lambda name: app
    pkg = types.ModuleType("win32com")
    pkg.client = client
    sys.modules["win32com"], sys.modules["win32com.client"] = pkg, client
    return app


def att(name):
    return Attachment(name, path=os.path.join(DATA, name))


class ClassicOutlookAdapterTests(Isolated, unittest.TestCase):
    def setUp(self):
        super().setUp()
        self._saved = {k: sys.modules.get(k) for k in ("win32com", "win32com.client")}
        t = lambda h: NOW - datetime.timedelta(hours=h)
        self.items = [
            MailItem("Tax Invoice 778899", t(3), [att("invoice_bidfood.pdf")], "Bidfood", "a@bidfood.example", "One Pearl Energy"),
            MailItem("Overdue reminder", t(5), [att("invoice_other.pdf")]),
            MailItem("FW: Invoice NS-4040", t(8), [att("invoice_company_named.pdf"), Attachment("logo.png", hidden=True, data=b"p" * 60000)]),
            MailItem("Invoice - photo attached", t(10), [Attachment("photo.jpg", data=b"j" * 70000)]),
            MailItem("old mail", t(24 * 9), []),
            OtherItem(),
        ]
        inbox = Folder("Inbox", self.items)
        stores = [Store("One@pearlenergy.com.au", inbox), Store("Public Folders", Folder("Inbox", store_id="S9"))]
        self.app = install_mock(stores)
        from invoicebot.outlook_io import ClassicOutlook
        self.gw = ClassicOutlook()

    def tearDown(self):
        for k, v in self._saved.items():
            sys.modules.pop(k, None)
            if v is not None:
                sys.modules[k] = v
        super().tearDown()

    def scan(self, hours=48, limit=100):
        return self.gw.scan("One@pearlenergy.com.au", "LABEL", NOW - datetime.timedelta(hours=hours), NOW, limit)

    def test_mailbox_names(self):
        self.assertEqual(self.gw.list_mailboxes(), ["One@pearlenergy.com.au", "Public Folders"])
        self.assertEqual(self.gw.resolve_store(["Nope", "one@pearlenergy.com.au"]), "One@pearlenergy.com.au")
        self.assertEqual(self.gw.resolve_store(["One@pearlenergy.com"]), "One@pearlenergy.com.au")      # lenient
        with self.assertRaises(LookupError):
            self.gw.resolve_store(["Missing mailbox"])

    def test_scan_is_newest_first_and_stops_at_the_window_start(self):
        infos, truncated = self.scan(hours=48)
        self.assertEqual([i.subject for i in infos], ["Tax Invoice 778899", "Overdue reminder", "FW: Invoice NS-4040", "Invoice - photo attached"])
        self.assertFalse(truncated)
        self.assertEqual(infos[0].mailbox, "LABEL")
        self.assertEqual(infos[0].sender_email, "a@bidfood.example")
        self.assertEqual(infos[0].to_name, "One Pearl Energy")
        self.assertIsInstance(infos[0].received, datetime.datetime)

    def test_scan_limit_reports_truncation(self):
        infos, truncated = self.scan(hours=24 * 30, limit=2)
        self.assertEqual(len(infos), 2)
        self.assertTrue(truncated)

    def test_attachments(self):
        infos, _ = self.scan()
        atts = self.gw.attachments(infos[2])                      # the FW: e-mail with a pdf and a hidden signature picture
        self.assertEqual([(a.index, a.name, a.hidden) for a in atts], [(1, "invoice_company_named.pdf", False), (2, "logo.png", True)])
        self.assertTrue(all(a.kind == 1 for a in atts))

    def test_save_attachment(self):
        infos, _ = self.scan()
        target = self.path("out.pdf")
        self.gw.save_attachment(infos[0], 1, target)
        self.assertEqual(os.path.getsize(target), os.path.getsize(os.path.join(DATA, "invoice_bidfood.pdf")))

    def test_mark_done_and_flags(self):
        infos, _ = self.scan()
        self.gw.mark_done(infos[0])
        self.assertEqual((self.items[0].UnRead, self.items[0].FlagStatus, self.items[0].saves), (False, 1, 1))
        self.gw.flag_red(infos[1])
        self.assertEqual(self.items[1].FlagStatus, 2)
        self.items[2].FlagStatus = 2                              # a person's own flag
        self.gw.mark_done(infos[2])
        self.gw.flag_red(infos[2])
        self.assertEqual((self.items[2].UnRead, self.items[2].FlagStatus), (False, 2))

    def put_special(self, item):
        self.app.GetNamespace("MAPI").Folders[0].Folders[1]._items.insert(0, item)     # the One mailbox's Inbox
        infos, _ = self.scan()
        return next(i for i in infos if i.entry_id == item.EntryID)

    def test_the_tick_is_set_another_way_when_the_first_is_ignored(self):
        item = IgnoresFlagStatus("Tax Invoice 1", NOW - datetime.timedelta(hours=1), [att("invoice_bidfood.pdf")])
        self.gw.mark_done(self.put_special(item))
        self.assertEqual((item.UnRead, item.FlagStatus), (False, 1))                     # MarkAsTask(olMarkComplete) did it

    def test_when_no_way_works_the_read_status_is_kept_and_the_error_says_why(self):
        item = RefusesEverything("Tax Invoice 2", NOW - datetime.timedelta(hours=1), [att("invoice_bidfood.pdf")])
        item.armed = True
        info = self.put_special(item)
        with self.assertRaises(RuntimeError) as cm:
            self.gw.mark_done(info)
        message = str(cm.exception)
        for way in ("FlagStatus:", "MarkAsTask:", "FlagStatus+date:"):
            self.assertIn(way, message)
        self.assertIn("cannot be set", message)
        self.assertFalse(item.UnRead)                                                    # still marked read
        self.assertGreaterEqual(item.saves, 1)

    def test_a_tick_that_outlook_drops_on_save_is_noticed_and_set_another_way(self):
        item = DropsDirectFlag("Tax Invoice 3", NOW - datetime.timedelta(hours=1), [att("invoice_bidfood.pdf")])
        self.gw.mark_done(self.put_special(item))
        self.assertEqual((item.UnRead, item.FlagStatus), (False, 1))

    def test_when_no_tick_works_a_fallback_category_is_added_and_the_error_says_so(self):
        item = RefusesEverything("Tax Invoice 4", NOW - datetime.timedelta(hours=1), [att("invoice_bidfood.pdf")])
        item.armed = True
        info = self.put_special(item)
        with self.assertRaises(RuntimeError) as cm:
            self.gw.mark_done(info)
        self.assertIn("the category 'Processed by bot' was added instead", str(cm.exception))
        self.assertIn("Processed by bot", item.Categories)
        self.assertFalse(item.UnRead)

    def test_the_fallback_category_can_be_switched_off(self):
        self.set_run(Tick_fallback_category="none")
        item = RefusesEverything("Tax Invoice 5", NOW - datetime.timedelta(hours=1), [att("invoice_bidfood.pdf")])
        item.armed = True
        info = self.put_special(item)
        with self.assertRaises(RuntimeError):
            self.gw.mark_done(info)
        self.assertEqual(item.Categories, "")

    def test_a_category_that_is_not_kept_is_reported(self):
        item = IgnoresCategories("FW: Invoice 6", NOW - datetime.timedelta(hours=1), [att("invoice_bidfood.pdf")])
        info = self.put_special(item)
        with self.assertRaises(RuntimeError) as cm:
            self.gw.add_category(info, "Forwarded / Reply", "red")
        self.assertIn("not on the e-mail when read back", str(cm.exception))

    def selftest(self, subject):
        return self.gw.diagnose_marking("One@pearl", subject, now=NOW)

    def test_the_self_test_reports_what_works_and_puts_the_email_back(self):
        item = MailItem("BOTTEST invoice", NOW - datetime.timedelta(hours=1), [att("invoice_bidfood.pdf")], flag=0)
        item.Categories = "Blue category"
        self.put_special(item)
        lines, ok = self.selftest("BOTTEST")
        text = "\n".join(lines)
        self.assertTrue(ok, text)
        for step in ("mark as read", "tick by setting FlagStatus", "tick by MarkAsTask(complete)", "red flag by setting FlagStatus", "category added to the e-mail"):
            self.assertIn("OK    " + step, text)
        self.assertIn("The bot can set the tick: yes, by setting FlagStatus", text)
        self.assertIn("The e-mail was put back as it was.", text)
        self.assertEqual((item.UnRead, item.FlagStatus, item.Categories), (True, 0, "Blue category"))             # exactly as it was

    def test_the_self_test_shows_which_way_works_when_the_first_does_not(self):
        item = IgnoresFlagStatus("BOTTEST two", NOW - datetime.timedelta(hours=1), [att("invoice_bidfood.pdf")])
        self.put_special(item)
        lines, ok = self.selftest("BOTTEST")
        text = "\n".join(lines)
        self.assertIn("FAIL  tick by setting FlagStatus", text)
        self.assertIn("The bot can set the tick: yes, by MarkAsTask(complete)", text)
        self.assertTrue(ok, text)

    def test_the_self_test_says_so_when_nothing_can_be_set(self):
        item = RefusesEverything("BOTTEST three", NOW - datetime.timedelta(hours=1), [att("invoice_bidfood.pdf")])
        item.armed = True
        self.put_special(item)
        lines, ok = self.selftest("BOTTEST")
        text = "\n".join(lines)
        self.assertFalse(ok)
        self.assertIn("The bot can set the tick: NO - it will add the fallback category instead", text)
        self.assertIn("cannot be set on this item", text)                                                           # Outlook's own words are shown

    def test_the_self_test_will_not_touch_an_email_that_has_a_flag(self):
        item = MailItem("BOTTEST flagged", NOW - datetime.timedelta(hours=1), [att("invoice_bidfood.pdf")], flag=2)
        self.put_special(item)
        lines, ok = self.selftest("BOTTEST")
        self.assertFalse(ok)
        self.assertIn("already has a follow-up flag", "\n".join(lines))
        self.assertEqual((item.FlagStatus, item.UnRead), (2, True))

    def test_the_self_test_explains_when_it_finds_nothing(self):
        self.assertIn("No e-mail with 'NOSUCHWORD'", "\n".join(self.selftest("NOSUCHWORD")[0]))
        self.assertIn("No mailbox in this Outlook contains", "\n".join(self.gw.diagnose_marking("nobody", "BOTTEST", now=NOW)[0]))

    def test_categories(self):
        infos, _ = self.scan()
        self.gw.add_category(infos[2], "Forwarded / Reply", "red")
        self.gw.add_category(infos[2], "Forwarded / Reply", "red")        # idempotent
        self.gw.add_category(infos[2], "Image attachment", "orange")
        self.assertEqual(self.items[2].Categories, "Forwarded / Reply, Image attachment")
        self.assertEqual(self.app.ns.Categories.names, {"Forwarded / Reply": 1, "Image attachment": 2})

    def test_create_mail(self):
        self.gw.create_mail("a@b.com", "Hi", "Body", send=False)
        self.gw.create_mail("c@d.com", "Hi2", "Body2", send=True)
        draft, sent = self.app.created
        self.assertEqual((draft.saved, draft.sent, sent.saved, sent.sent), (True, False, False, True))

    def test_the_whole_bot_through_the_adapter(self):
        s = run(self.gw)
        self.assertEqual(s.c["blocked"], 1)
        self.assertEqual(s.c["Filed"], 2)
        self.assertEqual(s.c["image-only"], 1)
        files = self.files_under("TestOutput")
        self.assertEqual(len(files), 2)
        self.assertTrue(any(f.endswith("Bidfood Australia Ltd_778899.pdf") for f in files))
        self.assertEqual(self.items[0].FlagStatus, 1)             # ticked
        self.assertEqual(self.items[1].FlagStatus, 2)             # overdue: red flag
        self.assertTrue(self.items[1].UnRead)
        self.assertEqual(self.items[2].FlagStatus, 2)           # forward: red flag, no category
        self.assertEqual(self.items[2].Categories, "")
        self.assertEqual(self.items[3].Categories, "Image attachment")
        self.assertEqual(s.c["outlook errors"], 0)
        s2 = run(self.gw)
        self.assertEqual(s2.c["documents"], 0)


if __name__ == "__main__":
    unittest.main()
