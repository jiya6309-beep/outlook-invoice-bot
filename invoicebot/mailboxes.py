"""Matching the mailboxes found in Outlook to the 13 companies (so nobody edits the Companies sheet by hand).

    python -m invoicebot.mailboxes        prints the Outlook mailbox names as JSON (used by the control panel)

suggest()  proposes which company each Outlook mailbox belongs to, from the words in its name.
apply()    writes the chosen mapping into the Companies sheet (or adds a TEST row for a personal/test mailbox)."""
import json
import re
import sys

from . import settings

COMMON = {"pearl", "energy", "pty", "ltd", "investment", "investments", "fuels", "com", "au", "the", "and", "no", "mailbox",
          "inbox", "outlook", "email", "gmail", "hotmail", "yahoo", "pearlenergy", "www"}
TEST_TEXT = "zz test only"          # a company row that no document can ever match


def _tokens(text):
    return {t for t in re.findall(r"[a-z0-9]+", str(text).lower()) if t not in COMMON}


def _company_tokens(row):
    return _tokens(" ".join([row["Folder name"], row["Company name"], row["Text to look for"]])) - {"1", "2", "3", "5"} | \
        {t for t in _tokens(row["Folder name"]) if t in {"one", "two", "three", "five"}}


def suggest(outlook_names, company_rows=None):
    """[{'outlook', 'company' (Company name or ''), 'why'}] - a suggestion only when exactly one company fits best."""
    rows = company_rows if company_rows is not None else settings.rows("Companies")
    out = []
    from . import companies
    for name in outlook_names:
        key = settings.norm(name)
        sp = companies.special_for(name)
        if sp:                                               # SpecialMailboxes sheet: never offered as a company
            why = {"never": "NEVER read (SpecialMailboxes sheet)", "on hold": "special mailbox - ON HOLD, not read",
                   "fleet card": "special mailbox - Fleet Card statements", "statements": "special mailbox - statements",
                   "fuel invoices": "special mailbox - fuel invoices (Inbox + its sub-folders)",
                   "security": "special mailbox - security reports (security run only)",
                   "read all": "read every e-mail - the company comes from each invoice"}.get(sp["role"], sp["role"])
            out.append({"outlook": name, "company": "", "why": why, "special": sp["role"]})
            continue
        exact = [r for r in rows if any(settings.norm(a) == key for a in r["Mailbox (as in Outlook)"].split("|") if a.strip())]
        if exact:
            out.append({"outlook": name, "company": exact[0]["Company name"], "why": "already set"})
            continue
        mine = _tokens(name)
        scored = sorted(((len(mine & _company_tokens(r)), r) for r in rows if r["Text to look for"] != TEST_TEXT),
                        key=lambda sr: -sr[0])
        if scored and scored[0][0] > 0 and (len(scored) == 1 or scored[1][0] < scored[0][0]):
            out.append({"outlook": name, "company": scored[0][1]["Company name"], "why": "the name contains: " + ", ".join(sorted(mine & _company_tokens(scored[0][1])))})
        else:
            out.append({"outlook": name, "company": "", "why": "no clear match"})
    return out


def test_folder_name(outlook_name):
    local = re.split(r"[@]", outlook_name)[0]
    return "TEST " + (re.sub(r"[^A-Za-z0-9]+", " ", local).strip().upper() or "MAILBOX")


def apply(items):
    """items = [{'outlook': name, 'action': 'company'|'special'|'test'|'skip', 'company': Company name, 'role': 'fuel invoices' |
    'fleet card' | 'security' | 'never'}]. Companies go to the Companies sheet, special roles to the SpecialMailboxes sheet; a mailbox is
    always in only one of them. Returns a list of messages."""
    messages = []

    singles = [it.get("role") for it in items if it.get("action") == "special" and it.get("role") not in ("never", "read all")]
    twice = sorted({r for r in singles if singles.count(r) > 1})
    if twice:
        raise settings.SettingsError("Only one mailbox can be " + " / ".join(twice) + ". Change one of them, then save.")

    def special_sheet(wb):
        if "SpecialMailboxes" not in wb.sheetnames:
            ws = wb.create_sheet("SpecialMailboxes")
            ws.append(settings.SHEETS["SpecialMailboxes"][0])
        return wb["SpecialMailboxes"]

    def drop_special(sp, name):
        """Remove a mailbox from the SpecialMailboxes sheet. When it is a row's main name, the whole row goes (its other names are
        only addresses of the same mailbox); when it is only an extra name, just that name is removed."""
        for r in range(sp.max_row, 1, -1):
            names = [a.strip() for a in str(sp.cell(r, 1).value or "").split("|") if a.strip()]
            if not any(settings.norm(a) == settings.norm(name) for a in names):
                continue
            if settings.norm(names[0]) == settings.norm(name):
                sp.delete_rows(r)
            else:
                sp.cell(r, 1).value = "|".join(a for a in names if settings.norm(a) != settings.norm(name))

    def drop_company(ws, col, name):
        for r in range(ws.max_row, 1, -1):
            names = [a.strip() for a in str(ws.cell(r, col("Mailbox (as in Outlook)")).value or "").split("|") if a.strip()]
            keep = [a for a in names if settings.norm(a) != settings.norm(name)]
            if len(keep) != len(names):
                if not keep and str(ws.cell(r, col("Text to look for")).value or "").strip() == TEST_TEXT:
                    ws.delete_rows(r)                           # a TEST row without its mailbox has no use left
                else:
                    ws.cell(r, col("Mailbox (as in Outlook)")).value = "|".join(keep)

    ROLE_TEXT = {"fuel invoices": "Fuel invoices", "fleet card": "Fleet card", "security": "Security", "never": "Never",
                 "read all": settings.READ_ALL_TEXT}

    def set_special(wb, ws, col, name, role):
        sp = special_sheet(wb)
        drop_company(ws, col, name)
        if role in ("never", "read all"):                       # any number of mailboxes can have these: one row each
            text = ROLE_TEXT[role]
            already = any(settings.norm(name) in [settings.norm(a) for a in str(sp.cell(r, 1).value or "").split("|")]
                          and str(sp.cell(r, 2).value or "").strip().lower() == text.lower() for r in range(2, sp.max_row + 1))
            if not already:
                drop_special(sp, name)
                sp.append([name, text, "", "Yes", "Set in the panel (Setup step 2): this mailbox is not read." if role == "never" else
                           "Set in the panel (Setup step 2): every e-mail is read (Inbox + sub-folders); the company and site come "
                           "from each document, never from this mailbox."])
            return f"{name}  ->  " + ("never read" if role == "never" else "read every e-mail (company from the invoice)")
        row = next((r for r in range(2, sp.max_row + 1) if str(sp.cell(r, 2).value or "").strip().lower().startswith(role.split()[0])), None)
        if row:
            names = [a.strip() for a in str(sp.cell(row, 1).value or "").split("|") if a.strip()]
            if names and settings.norm(names[0]) == settings.norm(name):
                return f"{name}  ->  {ROLE_TEXT[role]} (already)"
        drop_special(sp, name)
        row = next((r for r in range(2, sp.max_row + 1) if str(sp.cell(r, 2).value or "").strip().lower().startswith(role.split()[0])), None)
        if row is None:
            default = next((d for d in settings.SPECIAL_MAILBOXES if d[1].lower() == ROLE_TEXT[role].lower()), None)
            sp.append([name, ROLE_TEXT[role], default[2] if default else "", "Yes", default[4] if default else ""])
        else:
            sp.cell(row, 1).value = name
            sp.cell(row, 4).value = "Yes"
        return f"{name}  ->  {ROLE_TEXT[role]}"

    def change(wb):
        ws = wb["Companies"]
        head = {str(ws.cell(1, c).value or "").strip(): c for c in range(1, ws.max_column + 1)}
        col = lambda h: head[h]
        for it in items:
            name, action = it["outlook"].strip(), it.get("action", "skip")
            if not name or action == "skip":
                continue
            if action == "special":
                role = str(it.get("role", "")).strip().lower()
                if role not in ROLE_TEXT:
                    messages.append(f"{name}: unknown choice '{role}'")
                    continue
                messages.append(set_special(wb, ws, col, name, role))
                continue
            if action == "company":
                if "SpecialMailboxes" in wb.sheetnames:
                    drop_special(wb["SpecialMailboxes"], name)
                for r in range(2, ws.max_row + 1):
                    if str(ws.cell(r, col("Company name")).value or "").strip() == it["company"]:
                        old = [a.strip() for a in str(ws.cell(r, col("Mailbox (as in Outlook)")).value or "").split("|") if a.strip()]
                        new = [name] + [a for a in old if settings.norm(a) != settings.norm(name)]
                        ws.cell(r, col("Mailbox (as in Outlook)")).value = "|".join(new)
                        on = it.get("active", True) is not False         # the Active tick box in step 2
                        ws.cell(r, col("Active")).value = "Yes" if on else "No"
                        messages.append(f"{name}  ->  {it['company']}" + ("" if on else "  (switched OFF: not read)"))
                        break
                else:
                    messages.append(f"{name}: company '{it['company']}' not found")
            elif action == "test":
                exists = any(settings.norm(name) in [settings.norm(a) for a in str(ws.cell(r, col("Mailbox (as in Outlook)")).value or "").split("|")]
                             for r in range(2, ws.max_row + 1))
                if exists:
                    messages.append(f"{name}: already in the Companies sheet")
                    continue
                row = [None] * ws.max_column
                row[col("Company name") - 1] = "TEST - " + name
                row[col("Text to look for") - 1] = TEST_TEXT
                row[col("Folder name") - 1] = test_folder_name(name)
                row[col("Mailbox (as in Outlook)") - 1] = name
                row[col("Active") - 1] = "Yes"
                ws.append(row)
                messages.append(f"{name}  ->  new TEST mailbox, files go to the folder '{test_folder_name(name)}'")

    settings.edit(change)
    return messages


def main():
    """Prints {'outlook': [names]} or {'error': text}. Runs in its own process so Outlook problems can never hang the panel."""
    try:
        from .outlook_io import ClassicOutlook
        print(json.dumps({"outlook": ClassicOutlook().list_mailboxes()}))
    except ImportError:
        print(json.dumps({"error": "The Python package 'pywin32' is missing (run Start_Bot.bat) - or this is not a Windows PC."}))
    except Exception as e:
        print(json.dumps({"error": f"Could not connect to Outlook: {e}. Classic Outlook must be installed, open and signed in."}))


if __name__ == "__main__":
    sys.exit(main())
