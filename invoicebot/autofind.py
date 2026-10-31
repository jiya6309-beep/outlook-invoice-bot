"""Finds the synced SharePoint folder on THIS PC by itself, so nobody has to type a path.

It looks for a folder named  EMAIL INVOICES  whose parent is  OFFICE INDIA  (both names are settings) in C:\\SHAREPOINT and
anywhere under the user's profile / OneDrive folders, for example
    C:\\Users\\jiya\\pearlenergy.com.au\\Patel Group - Documents\\OFFICE INDIA\\EMAIL INVOICES
It only LISTS folder names (it never opens or downloads a file), skips folders that cannot contain it, and stops after a time limit.
The folder found is remembered in data\\destination.txt and checked again at every run."""
import collections
import os
import time

from . import paths, settings

SKIP = {"appdata", "node_modules", ".git", "windows", "program files", "program files (x86)", "programdata", "$recycle.bin",
        "downloads", "music", "videos", "pictures", "3d objects", "saved games", "contacts", "favorites", "links", "searches",
        "ntuser.dat", "microsoft", "__pycache__"}
session_root = ""
EXTRA_ROOTS = [r"C:\SHAREPOINT", r"D:\SHAREPOINT"]     # 5.15.0: SharePoint synced outside the user folder (as on the office PC)            # found during this run (a rehearsal does not write the file)


def cache_file():
    return paths.p("data", "destination.txt")


def cached():
    """The remembered folder if it still exists, else ''."""
    if session_root and os.path.isdir(session_root):
        return session_root
    try:
        with open(cache_file(), encoding="utf-8") as f:
            path = f.read().strip()
    except OSError:
        return ""
    return path if path and os.path.isdir(path) else ""


def remember(path):
    os.makedirs(os.path.dirname(cache_file()), exist_ok=True)
    with open(cache_file(), "w", encoding="utf-8") as f:
        f.write(path)


def search_roots():
    """Where to look: the user profile and the OneDrive folders (duplicates and nested ones removed)."""
    found = [os.path.abspath(r) for r in EXTRA_ROOTS if os.path.isdir(r)]
    for var in ("OneDriveCommercial", "OneDrive", "OneDriveConsumer", "USERPROFILE", "HOME"):
        value = os.environ.get(var)
        if value and os.path.isdir(value):
            found.append(os.path.abspath(value))
    roots = []
    for r in found:
        if not any(r == k or r.startswith(k + os.sep) for k in roots):
            roots = [k for k in roots if not k.startswith(r + os.sep)] + [r]
    return roots


def find(leaf=None, parent=None, max_depth=8, max_seconds=40):
    """Every folder called `leaf` inside a folder called `parent`, shortest path first. [] if none."""
    leaf = (leaf if leaf is not None else settings.run_value("Auto-find folder name", "EMAIL INVOICES")).strip().lower()
    parent = (parent if parent is not None else settings.run_value("Auto-find parent folder", "OFFICE INDIA")).strip().lower()
    if not leaf:
        return []
    deadline = time.time() + max_seconds
    hits = []
    for root in search_roots():
        queue = collections.deque([(root, 0)])
        while queue and time.time() < deadline:
            current, depth = queue.popleft()
            try:
                entries = list(os.scandir(current))
            except OSError:
                continue
            for e in entries:
                try:
                    if not e.is_dir(follow_symlinks=False):
                        continue
                except OSError:
                    continue
                name = e.name.lower()
                if name == leaf and (not parent or os.path.basename(current).lower() == parent):
                    hits.append(e.path)
                elif name not in SKIP and depth + 1 < max_depth:
                    queue.append((e.path, depth + 1))
    return sorted(set(hits), key=lambda p: (len(p), p))


def discover(remember_it=True):
    """Find the folder; remember it. Returns (path or '', all candidates)."""
    global session_root
    hits = find()
    if not hits:
        return "", []
    session_root = hits[0]
    if remember_it:
        remember(hits[0])
    return hits[0], hits
