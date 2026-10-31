"""Where a file goes, and writing it there safely.

    <root> \\ <MONTH> <YEAR> EMAIL INVOICES \\ <COMPANY FOLDER> \\ Supplier_Number.pdf
    <holding> \\ _Unassigned                      known document, destination not available (yet)
    <holding> \\ _Unrecognised                    could not be read / identified
    <holding> \\ _Unrecognised \\ Statement ...   parked on purpose (statements, account summaries, ...)

Safe writing: copy to  name.part , check the size, then rename. A file is never overwritten.
Everything is a no-op in dry mode (shadow runs), which only reports where files WOULD go."""
import datetime
import os
import re
import shutil
import time

from . import naming, settings

MONTHS = ["JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE", "JULY", "AUGUST", "SEPTEMBER", "OCTOBER",
          "NOVEMBER", "DECEMBER"]
MAX_PATH = 400                  # SharePoint's limit; classic Windows is 260 unless long paths are enabled
UNASSIGNED_DIR = "_Unassigned"
UNRECOGNISED_DIR = "_Unrecognised"


class ConfigError(Exception):
    """The setup is incomplete (for example live mode without a SharePoint path)."""


class PlacementError(Exception):
    def __init__(self, reason, retryable=True):
        super().__init__(reason)
        self.reason, self.retryable = reason, retryable


def month_folder_name(when, pattern=None):
    pattern = pattern or settings.run_value("Month folder pattern", "{MONTH} {YEAR} EMAIL INVOICES")
    name = MONTHS[when.month - 1]
    return (pattern.replace("{MONTH}", name).replace("{Month}", name.title()).replace("{month}", name.lower())
            .replace("{YEAR}", f"{when.year:04d}").replace("{YY}", f"{when.year % 100:02d}")
            .replace("{MM}", f"{when.month:02d}"))


def folder_clean(name):
    return naming.clean_folder(name)


DEFAULT_PATTERN = "{MONTH_FOLDER}\\{COMPANY}"
KNOWN_TOKENS = ("{MONTH_FOLDER}", "{COMPANY}", "{SITE}", "{MONTH}", "{Month}", "{month}", "{YEAR}", "{YY}", "{MM}")


def split_pattern(pattern):
    return [seg.strip() for seg in re.split(r"[\\/]+", pattern or "") if seg.strip()]


def segment_kind(segment):
    """Which kind of folder a part of the layout is: 'site', 'company' or 'month'."""
    return "site" if "{SITE}" in segment else ("company" if "{COMPANY}" in segment else "month")


def holding_month(arrival, doc, window):
    """Month sub-folder of the holding folders: 'October 2026'. The document's own date decides when it is within `window` months of the
    arrival month; otherwise (no date, or far away) the month the e-mail arrived in."""
    d = arrival
    if doc is not None and abs((arrival.year * 12 + arrival.month) - (doc.year * 12 + doc.month)) <= window:
        d = doc
    return f"{MONTHS[d.month - 1].title()} {d.year}"


def _same_folder_other_spelling(parent, name):
    """An existing sub-folder of `parent` whose name is the same apart from capitals, extra spaces and punctuation, else ''.
    Stops LIVE mode from creating a second folder beside the real one ('AMPOL  KELSO' and 'Ampol Kelso')."""
    key = re.sub(r"[^a-z0-9]", "", str(name).lower())
    if not key or not os.path.isdir(parent):
        return ""
    try:
        hits = [d for d in os.listdir(parent) if os.path.isdir(os.path.join(parent, d)) and re.sub(r"[^a-z0-9]", "", d.lower()) == key]
    except OSError:
        return ""
    return hits[0] if len(hits) == 1 else ""


class Router:
    def __init__(self, dry=False, log=print):
        self.dry, self.log = dry, log
        self.mode = settings.destination_mode()
        self.root = settings.destination_root()
        self.holding = settings.holding_root()
        self.pattern = settings.run_value("Destination pattern", DEFAULT_PATTERN) or DEFAULT_PATTERN
        self.segments = split_pattern(self.pattern)
        self.bad = {}               # folder -> reason, remembered for the whole run (one dead NAS cannot stall the run)
        self.checked = set()
        like_test = settings.run_bool("Live mode creates missing folders like test mode", False)     # one switch: LIVE behaves like TEST
        self.create = {"month": like_test or settings.run_bool("Create month folder in live mode", True),
                       "company": like_test or settings.run_bool("Create company folder in live mode", True),
                       "site": like_test or settings.run_bool("Create site folder in live mode", False)}

    # ---------------------------------------------------------------- the folder layout
    def check_config(self):
        if self.mode == "live" and not (self.root and os.path.isdir(self.root)):
            raise ConfigError(
                "LIVE mode, but the SharePoint folder was not found" + (f":\n   {self.root}\n" if self.root else ".\n") +
                "   Nothing was written. Check that OneDrive / SharePoint has synced the EMAIL INVOICES folder on this PC,\n"
                "   or correct 'Live destination root' (Bot_Settings.xlsx, sheet Run - or the panel, Setup 1),\n"
                "   or set 'Destination mode' back to test.")
        if not self.segments:
            raise ConfigError("'Destination pattern' in Bot_Settings.xlsx (sheet Run) is empty. Example:  {SITE}")
        try:
            settings.check_output_root(self.holding, "Holding")
            settings.check_output_root(settings.reports_root(), "Reports")
        except settings.SettingsError as e:
            raise ConfigError(str(e))
        for seg in self.segments:
            left = seg
            for token in KNOWN_TOKENS:
                left = left.replace(token, "")
            if "{" in left or "}" in left:
                raise ConfigError(f"'Destination pattern' (sheet Run) contains an unknown word in '{seg}'. "
                                  "Allowed: {SITE} {COMPANY} {MONTH_FOLDER} {MONTH} {YEAR} {MM}")

    def needs(self):
        """The kinds of folder the layout uses: a subset of {'site', 'company', 'month'}."""
        return {segment_kind(seg) for seg in self.segments}

    def site_is_last(self):
        return bool(self.segments) and segment_kind(self.segments[-1]) == "site"

    def render(self, when, company="", site="", skip_site=False):
        out = []
        for seg in self._segments(skip_site):
            text = seg.replace("{MONTH_FOLDER}", month_folder_name(when)).replace("{COMPANY}", company).replace("{SITE}", site)
            out.append(folder_clean(month_folder_name(when, text) if text else text))
        return out

    def month_dir(self, when):
        return os.path.join(self.root, folder_clean(month_folder_name(when)))

    def company_dir(self, when, folder_name):
        return os.path.join(self.root, *self.render(when, company=folder_name))

    def prepare_company_dir(self, when, folder_name):
        return self.prepare_dir(when, company=folder_name)

    def _segments(self, skip_site=False):
        return [s for s in self.segments if not (skip_site and segment_kind(s) == "site")]

    def prepare_dir(self, when, company="", site="", skip_site=False):
        """The folder for this document, ready to be written to. Raises PlacementError if not usable.
        skip_site: leave the site level out (the document is filed one level up, in the company folder)."""
        parts = self.render(when, company, site, skip_site)
        target = os.path.join(self.root, *parts)
        if target in self.bad:
            raise PlacementError(self.bad[target])
        if len(target) > MAX_PATH - 40:
            raise PlacementError(f"path too long ({len(target)} characters)", retryable=False)
        try:
            if self.mode == "test":
                if not self.dry:
                    os.makedirs(target, exist_ok=True)
            else:
                if not os.path.isdir(self.root):
                    raise PlacementError(f"destination root is not reachable: {self.root}")
                current = self.root
                fixed = []
                for template, part in zip(self._segments(skip_site), parts):
                    current = os.path.join(current, part)
                    if os.path.isdir(current):
                        fixed.append(part)
                        continue
                    same = _same_folder_other_spelling(os.path.dirname(current), part)       # 'AMPOL  KELSO' for 'AMPOL KELSO': use the real one
                    if same:
                        current = os.path.join(os.path.dirname(current), same)
                        fixed.append(same)
                        continue
                    fixed.append(part)
                    kind = segment_kind(template)
                    if not self.create[kind]:
                        raise PlacementError({"month": f"month folder does not exist yet: {current}",
                                              "company": f"company folder does not exist: {current}",
                                              "site": f"site folder does not exist: {current}"}[kind])
                    if not self.dry:
                        os.makedirs(current, exist_ok=True)
                target = os.path.join(self.root, *fixed)
            self._check_writable(target)
        except PlacementError as e:
            self.bad[target] = e.reason
            raise
        except OSError as e:
            self.bad[target] = f"cannot use folder ({type(e).__name__}): {e}"
            raise PlacementError(self.bad[target])
        return target

    def _check_writable(self, folder):
        if self.dry or folder in self.checked:
            return
        probe = os.path.join(folder, f".bot_write_test_{os.getpid()}")
        try:
            with open(probe, "w") as f:
                f.write("ok")
            os.remove(probe)
        except OSError as e:
            raise PlacementError(f"folder is not writable: {e}")
        self.checked.add(folder)

    def month_warning(self, today=None):
        """In live mode, warn in the last 3 days of a month if next month's folder is still missing."""
        today = today or datetime.date.today()
        if self.mode != "live" or not self.root or not self.segments or self.segments[0] != "{MONTH_FOLDER}":
            return None                                    # only a layout that starts with the month folder has a 'next month' to prepare
        nxt = (today.replace(day=28) + datetime.timedelta(days=4)).replace(day=1)
        if (nxt - today).days <= 3 and not os.path.isdir(self.month_dir(nxt)):
            return f"next month's folder does not exist yet: {self.month_dir(nxt)}"
        return None

    # ---------------------------------------------------------------- month folders and moving files
    def starts_with_month(self):
        return bool(self.root) and bool(self.segments) and self.segments[0] == "{MONTH_FOLDER}"

    def latest_month_dir_before(self, when, look_back=6):
        """The newest existing month folder before `when` (to copy the structure from), or None."""
        y, m = when.year, when.month
        for _ in range(look_back):
            m -= 1
            if m == 0:
                y, m = y - 1, 12
            candidate = self.month_dir(datetime.date(y, m, 1))
            if os.path.isdir(candidate):
                return candidate
        return None

    def copy_tree(self, source, target):
        """Create in `target` every folder that `source` has - FOLDERS ONLY, no files. Existing folders are left alone, so it can be repeated.
        Returns how many folders were created."""
        made = 0
        for folder, dirs, _files in os.walk(source):
            dirs[:] = sorted(d for d in dirs if not d.startswith("."))
            rel = os.path.relpath(folder, source)
            here = target if rel == "." else os.path.join(target, rel)
            if not os.path.isdir(here):
                if not self.dry:
                    os.makedirs(here, exist_ok=True)
                made += 1
            for d in dirs:
                sub = os.path.join(here, d)
                if not os.path.isdir(sub):
                    if not self.dry:
                        os.makedirs(sub, exist_ok=True)
                    made += 1
        return made

    def move(self, old_path, new_dir, name=None):
        """Move a filed document into another folder: copied, size-checked, and only then the original removed. Returns the new path.
        Raises PlacementError (not retryable) if the file is gone - someone moved or renamed it - or the copy could not be checked."""
        if not os.path.isfile(old_path):
            raise PlacementError("the file is no longer where the bot put it (it was moved or renamed)", retryable=False)
        if not self.dry:
            os.makedirs(new_dir, exist_ok=True)
        final, how = self.place(old_path, new_dir, name or os.path.basename(old_path))
        if os.path.abspath(final) == os.path.abspath(old_path):
            return final
        if not self.dry:
            try:
                os.remove(old_path)
            except OSError as e:
                raise PlacementError(f"copied, but the original could not be removed ({e})", retryable=False)
        return final

    # ---------------------------------------------------------------- writing files
    def place(self, src, folder, name):
        """Copy src into folder as `name`. Returns (final path, 'copied' | 'identical' | 'renamed')."""
        final = os.path.join(folder, name)
        how = "copied"
        if os.path.exists(final):
            if os.path.getsize(final) == os.path.getsize(src):
                return final, "identical"
            name = naming.unique_name(name, lambda c: os.path.exists(os.path.join(folder, c)))
            final = os.path.join(folder, name)
            how = "renamed"
        if len(final) > MAX_PATH:
            raise PlacementError(f"path too long ({len(final)} characters)", retryable=False)
        if self.dry:
            return final, how
        part = final + ".part"
        try:
            shutil.copyfile(src, part)
            if os.path.getsize(part) != os.path.getsize(src):
                os.remove(part)
                raise PlacementError("copy check failed (size differs)")
            os.replace(part, final)
        except PlacementError:
            raise
        except OSError as e:
            try:
                if os.path.exists(part):
                    os.remove(part)
            except OSError:
                pass
            raise PlacementError(f"could not write the file ({type(e).__name__}): {e}")
        return final, how

    def ensure_holding_folders(self):
        """Creates the holding folders (_Unassigned, _Unrecognised and one for each 'Holding' document type, e.g. _Statements)."""
        names = [UNASSIGNED_DIR, UNRECOGNISED_DIR]
        for r in settings.rows("DocTypes"):
            if r["Route"].strip().lower().startswith("hold") and r["Subfolder"].strip():
                names.append(r["Subfolder"].strip())
        for name in dict.fromkeys(names):
            try:
                os.makedirs(os.path.join(self.holding, folder_clean(name)), exist_ok=True)
            except OSError as e:
                self.log(f"could not create the holding folder '{name}': {e}")

    def holding_dir(self, *parts):
        return self.folder_under(self.holding, parts)

    def folder_under(self, base, parts):
        """base\\part1\\part2 ...: each part is created only if it does not exist yet; an existing folder whose name differs only in
        capitals, spaces or punctuation is used instead of making a second one."""
        current = base
        for part in [folder_clean(p) for p in parts if p]:
            nxt = os.path.join(current, part)
            if not os.path.isdir(nxt):
                same = _same_folder_other_spelling(current, part)
                if same:
                    nxt = os.path.join(current, same)
            current = nxt
        if not self.dry:
            os.makedirs(current, exist_ok=True)
        return current

    def special_dir(self, root, *parts):
        """A special mailbox's own destination (Fleet Card / M STATEMENTS): <root>\\<SITE>. The root must exist (it is a synced
        SharePoint folder); the site folder is created only when it is missing."""
        if not root:
            raise PlacementError("no Destination folder is set for this mailbox (SpecialMailboxes sheet)", retryable=False)
        if not os.path.isdir(root) and settings.is_test_mode() and not self.dry:
            os.makedirs(root, exist_ok=True)                   # TEST: its own folder inside TestOutput
        if not os.path.isdir(root) and not (self.dry and settings.is_test_mode()):
            raise PlacementError(f"destination folder is not reachable (check the SharePoint sync): {root}")
        try:
            folder = self.folder_under(root, parts)
            self._check_writable(folder)
        except OSError as e:
            raise PlacementError(f"cannot use folder ({type(e).__name__}): {e}")
        return folder

    def hold(self, src, parts, name):
        """Put src in a holding folder, e.g. parts = ('_Unrecognised', 'Statement'). Never raises for a missing folder."""
        folder = self.holding_dir(*parts)
        try:
            return self.place(src, folder, name)
        except PlacementError as e:
            raise PlacementError("holding folder problem: " + e.reason)

    # ---------------------------------------------------------------- local scratch / archive
    @staticmethod
    def archive(src, ref, original_name, days):
        """Keep the original file for a few days on the bot PC (days = 0: just delete the scratch copy)."""
        if days <= 0:
            os.remove(src)
            return ""
        folder = os.path.join(settings.archive_root(), datetime.date.today().isoformat())
        os.makedirs(folder, exist_ok=True)
        target = os.path.join(folder, f"{ref}_{naming.clean(original_name) or 'file'}")
        shutil.move(src, target)
        return target

    @staticmethod
    def cleanup_archive(days):
        root = settings.archive_root()
        if not os.path.isdir(root):
            return 0
        cutoff = datetime.date.today() - datetime.timedelta(days=max(days, 0))
        removed = 0
        for name in os.listdir(root):
            try:
                d = datetime.date.fromisoformat(name)
            except ValueError:
                continue                                       # only date-named folders are ever touched
            if d < cutoff or days <= 0:
                shutil.rmtree(os.path.join(root, name), ignore_errors=True)
                removed += 1
        return removed

    @staticmethod
    def cleanup_scratch(older_hours=24):
        root = settings.scratch_root()
        if not os.path.isdir(root):
            return 0
        limit, removed = time.time() - older_hours * 3600, 0
        for name in os.listdir(root):
            path = os.path.join(root, name)
            try:
                if os.path.isfile(path) and os.path.getmtime(path) < limit:
                    os.remove(path)
                    removed += 1
            except OSError:
                pass
        return removed
