"""The 13 companies: which mailbox belongs to which company, which SharePoint folder, how documents name them."""
import os

from . import paths, settings


def _active(r):
    return settings.yes(r.get("Active", ""), default=True)


def all_active():
    return [r for r in settings.rows("Companies") if _active(r)]


def mailbox_names(row):
    return [a.strip() for a in str(row["Mailbox (as in Outlook)"]).split("|") if a.strip()]


def for_mailbox(mailbox):
    """The company row for an Outlook mailbox name (exact, then lenient), or None."""
    want = settings.norm(mailbox)
    if not want:
        return None
    for r in all_active():
        if any(settings.norm(n) == want for n in mailbox_names(r)):
            return r
    for r in all_active():
        for n in mailbox_names(r):
            nn = settings.norm(n)
            if nn and (nn in want or want in nn):
                return r
    return None


def find_in_text(text, header_lines=25):
    """Companies named in the document, earliest first: [(row, position)]. Looks mainly at the top."""
    lines = [l for l in str(text).splitlines() if l.strip()]
    top = "\n".join(lines[:header_lines])
    found = {}
    for r in all_active():
        key = settings.norm(r["Text to look for"])
        if not key:
            continue
        pos = settings.norm(top).find(key)
        if pos >= 0:
            found[r["Company name"]] = (r, pos)
    out = sorted(found.values(), key=lambda rp: rp[1])
    return out


def for_sender(sender_name, sender_email=""):
    """The company whose mailbox the e-mail came FROM (a company mailbox that forwards or shares an invoice).
    Strict on purpose: the sender must contain a full mailbox name of the company (8+ letters) or equal its address,
    so a short sender name such as 'One' never matches a company."""
    name, addr = settings.norm(sender_name), settings.norm(sender_email)
    best = None
    for r in all_active():
        for alt in mailbox_names(r) + [r["Folder name"]]:              # staff call a company by its folder name: a sender may carry it
            a = settings.norm(alt)
            if len(a) >= 8 and (a in name or (addr and a == addr)) and (best is None or len(a) > best[0]):
                best = (len(a), r)
    return best[1] if best else None


def by_name(name):
    """The company row for a company name as written in the Companies sheet or on a site list
    ('Pearl Energy No 1 Pty Ltd', 'Central Coast Energy 1 Family Trust' ...). A row with an empty
    'Text to look for' can only be found by its exact company name."""
    key = settings.norm(name)
    if not key:
        return None
    for r in settings.rows("Companies"):
        if settings.norm(r["Company name"]) == key:
            return r
    for r in settings.rows("Companies"):
        text = settings.norm(r["Text to look for"])
        if text and text in key:
            return r
    return None


def resolve(doc_company_name, site_company, mailbox, order=None, sender_company=""):
    """Choose the company (and so the folder). Returns (row or None, source, warnings).
    doc_company_name = a company named on the document; site_company = the company of the matched site;
    sender_company = the company mailbox the e-mail came from; mailbox = the mailbox it arrived in."""
    order = order or [s.strip() for s in settings.run_value("Company source order", "document,site,sender,mailbox").split(",")]
    if not settings.run_bool("Use sender as company clue", True):
        order = [o for o in order if o != "sender"]
    elif "sender" not in order and "mailbox" in order:          # a workbook from before version 5.1.3: add the sender before the mailbox
        order.insert(order.index("mailbox"), "sender")
    candidates, warnings = {}, []
    if sender_company:
        row = by_name(sender_company)
        if row:
            candidates["sender"] = row
    if doc_company_name:
        row = by_name(doc_company_name)
        if row:
            candidates["document"] = row
    if site_company:
        row = by_name(site_company)
        if row:
            candidates["site"] = row
    mb = None if special_for(mailbox) else (by_name(mailbox) or for_mailbox(mailbox))   # a special mailbox (Fleet card, Statements) is never a company clue
    if mb:
        candidates["mailbox"] = mb
    for source in order:
        if source in candidates:
            chosen = candidates[source]
            others = {k: v for k, v in candidates.items() if v["Company name"] != chosen["Company name"]}
            if others:
                warnings.append("company differs by source: " + ", ".join(
                    f"{k}={v['Folder name']}" for k, v in {source: chosen, **others}.items()))
            return chosen, source, warnings
    return None, "", warnings


# ----------------------------------------------------------------------------------------------- special mailboxes
# Mailboxes that are NOT one company (SpecialMailboxes sheet): Fleet card (P I Fuel Invoice), Statements (mstatement@),
# On hold (Security) and Never (office@, serviceinvoices@). Such a mailbox is never used as a company clue.

def _role(value):
    v = str(value or "").strip().lower()
    for role in settings.SPECIAL_ROLES:
        if v.startswith(role.split()[0]):
            return role
    return ""


def special_rows():
    out = []
    for r in settings.rows("SpecialMailboxes"):
        if not settings.yes(r.get("Active", ""), default=True):
            continue
        names = [a.strip() for a in str(r["Mailbox (as in Outlook)"]).split("|") if a.strip()]
        role = _role(r.get("Role"))
        if names and role:
            dest = str(r.get("Destination folder") or "").strip()
            if dest and settings.is_test_mode():            # TEST: never SharePoint - its own folder inside TestOutput
                dest = paths.p("TestOutput", settings.TEST_FOLDERS.get(role, role.upper()))   # made when a file is placed
            out.append({"names": names, "role": role, "dest": dest, "label": names[0]})
    return out


def special_for(mailbox):
    """The special-mailbox entry for an Outlook mailbox name (exact, ignoring capitals and punctuation), or None."""
    key = settings.norm(mailbox)
    if not key:
        return None
    for r in special_rows():
        if any(settings.norm(n) == key for n in r["names"]):
            return r
    return None


def never_names():
    """The squashed names of every mailbox marked Never / On hold (SpecialMailboxes sheet)."""
    return {settings.norm(n) for r in special_rows() if r["role"] in ("never", "on hold") for n in r["names"] if settings.norm(n)}


def is_excluded(row):
    """A Companies row whose mailbox is marked Never / On hold in the SpecialMailboxes sheet is not read."""
    return any((special_for(n) or {}).get("role") in ("never", "on hold") for n in mailbox_names(row))


def reading_list():
    """Every mailbox a run reads: the active companies (minus Never / On hold mailboxes) and the Fleet card / Statements mailboxes.
    A special mailbox is returned as a company-like row: 'Company name' = its label, no folder, plus '_role' and '_dest'."""
    rows = [r for r in all_active() if not is_excluded(r)]
    for sp in special_rows():
        if sp["role"] in ("fleet card", "statements", "fuel invoices", "read all"):
            rows.append({"Company name": sp["label"], "Text to look for": "", "Folder name": "",
                         "Mailbox (as in Outlook)": "|".join(sp["names"]), "Active": "Yes", "_role": sp["role"], "_dest": sp["dest"]})
    return rows


def by_folder(folder):
    """The company row whose SharePoint folder name is this (capitals, spaces and punctuation ignored), or None."""
    key = settings.norm(folder)
    if not key:
        return None
    for r in settings.rows("Companies"):
        if settings.norm(r["Folder name"]) == key:
            return r
    return None
