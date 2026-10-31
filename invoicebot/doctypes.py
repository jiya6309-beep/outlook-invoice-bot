"""Document types: what a subject / title / file name means, and what each type does."""
import re

from . import settings

HEADER_MIN_LEN = 7        # in a document TITLE only longer keywords count ('invoice', 'credit note', 'statement', ...)


def _active(value):
    return settings.yes(value, default=True)


DOCKET_PHRASES = re.compile(r"\bdelivery\s+docket\b|\bdel\.?\s+docket\b|\bdelivery\s+note\b|\bgoods\s+received\b|"
                            r"\bproof\s+of\s+delivery\b|\bdelivery\s+receipt\b|\bpacking\s+slip\b", re.IGNORECASE)
# 5.16.0: documents that are an ORDER, not an invoice (Fresho / Rarity, PFD sales orders, MobiGo order printouts)
NOT_INVOICE = re.compile(r"\bthis\s+is\s+not\s+an?\s+(?:tax\s+)?invoice\b|\bnot\s+a\s+tax\s+invoice\b|\bindicative\s+total\b", re.IGNORECASE)
ORDER_TITLE = re.compile(r"\bsales\s+order\b|^\s*order\s*(?:#|no\.?|number|:)\s*[A-Z]{0,4}\d", re.IGNORECASE | re.MULTILINE)
INVOICE_TITLE = re.compile(r"\btax\s+invoice\b|\binvoice\s+(?:no|number|date|#)\b|\binv\s+(?:no|date)\b|\bcredit\s+note\b|"
                           r"\badjustment\s+note\b|\bstatement\b", re.IGNORECASE)
STATEMENT_PAGE = re.compile(r"\boverdue\b[^\n]{0,40}\bcurrent\b[^\n]{0,40}\b(?:total\s+amount\s+due|balance\s+due)\b", re.IGNORECASE)
# 5.20.0: 'Overdue  Current  Total Amount Due' (an account statement, e.g. Neverfail - its title 'statement' is a picture)
OC = "Order Confirmation"
CAR = "Customer Activity Report"
TAX_INVOICE = re.compile(r"\btax\s+invoice\b", re.IGNORECASE)
INVOICE_MARKERS = re.compile(r"\btax\s+invoice\b|\binvoice\s+(?:no|number|date|#)\b|\bamount\s+due\b|\btotal\s+(?:inc|incl)\b", re.IGNORECASE)


def keyword_rows():
    out, have = [], set()
    for r in settings.rows("Keywords"):
        have.add(settings.words(r["Keyword"]))
        if _active(r.get("Active", "")) and r["Keyword"].strip():
            out.append((settings.words(r["Keyword"]), r["Document type"].strip()))
    for kw, typ in settings.ADDED_KEYWORDS:                  # docket words: added for sheets made before they existed (a row in the sheet, even Active = No, wins)
        if kw not in have:
            out.append((kw, typ))
    out.sort(key=lambda kv: -len(kv[0]))
    return out


def docket_in_text(text):
    """The page itself says it is a delivery docket / note / goods-received document and shows no sign of being an invoice."""
    return bool(DOCKET_PHRASES.search(text or "")) and not INVOICE_MARKERS.search(text or "")


def _find(text, keyword, plurals=False):
    """Position of the keyword in the text as a whole word (digits may follow: INV8024), or -1.
    With plurals, a keyword of 5+ letters also matches with a plural ending: invoice -> invoices, credit note -> credit notes."""
    ending = r"(?:s|es)?" if plurals and len(keyword) >= 5 else ""
    m = re.search(r"(?<![a-z0-9])" + re.escape(keyword).replace(r"\ ", " ") + ending + r"(?![a-z])", text)
    return m.start() if m else -1


HEADER_IGNORE = {"service"}      # in a page TITLE area 'service' is part of company names ('Snackfood Services'), not a document type


def best_match(text, min_len=0, ignore=()):
    """(document type, keyword) of the longest keyword found in the text, earliest first on a tie."""
    t = settings.words(text)
    best = None
    plurals = settings.run_bool("Keywords also match plurals", True)
    for kw, typ in keyword_rows():
        if len(kw) < min_len or kw in ignore:
            continue
        pos = _find(t, kw, plurals)
        if pos >= 0:
            cand = (len(kw), -pos, typ, kw)
            if best is None or cand[:2] > best[:2]:
                best = cand
    return (best[2], best[3]) if best else (None, "")


def subject_matches(subject):
    """True if the subject contains any active document keyword."""
    return best_match(subject)[0] is not None


def blocked_word(subject):
    """The first blocked word found in the subject (whole word / phrase), or None."""
    t = settings.words(subject)
    for r in settings.rows("BlockedWords"):
        if not _active(r.get("Active", "")):
            continue
        w = settings.words(r["Word"])
        if w and settings.has_phrase(t, w):
            return r["Word"].strip()
    return None


def sender_row(sender_name, sender_email):
    """The Senders row that matches the sender (spaces and punctuation ignored), or None."""
    hay = settings.norm(f"{sender_name} {sender_email}")
    best = None
    for r in settings.rows("Senders"):
        key = settings.norm(r["Sender contains"])
        if key and key in hay and (best is None or len(key) > best[0]):
            best = (len(key), r)
    return best[1] if best else None


def always_download(sender_row_):
    return bool(sender_row_) and settings.yes(sender_row_.get("Always download", ""))


def decide(subject="", filename="", header="", sender=None, page_text="", supplier=""):
    """Work out the document type. Order: document title, what the page says (a delivery docket), subject, file name, exception sender.
    Words inside the supplier's own name ('Snackfood SERVICES') never count as a title.
    Returns (type or None, how it was decided)."""
    if supplier:
        header = re.sub(re.escape(supplier), " ", header, flags=re.IGNORECASE)
    has_oc = rule(OC) is not None
    header = NOT_INVOICE.sub(" ", header)                         # 'this is not a tax invoice' never counts as the title 'tax invoice'
    typ, kw = best_match(header, min_len=HEADER_MIN_LEN, ignore=HEADER_IGNORE)
    if typ == OC and INVOICE_TITLE.search(f"{header}\n{page_text or ''}"):  # 'Sales Order No' on a TAX INVOICE: still an invoice
        others = [k for k, t in keyword_rows() if t == OC]
        typ, kw = best_match(header, min_len=HEADER_MIN_LEN, ignore=set(HEADER_IGNORE) | set(others))
    if typ == "POD" and TAX_INVOICE.search(header):               # 5.18.0: 'TAX INVOICE' + 'Delivery Note 433201' (Lactalis) or a
        others = [k for k, t in keyword_rows() if t == "POD"]     # 'Delivery Docket' copy on the page (PacPlus): still an invoice
        typ, kw = best_match(header, min_len=HEADER_MIN_LEN, ignore=set(HEADER_IGNORE) | set(others))
    if typ in (None, "Invoice") and STATEMENT_PAGE.search(page_text or "") and not TAX_INVOICE.search(header) and rule("Statement"):
        return "Statement", "the page is an account statement (Overdue / Current / Total Amount Due)"
    if typ in (None, "Statement") and best_match(subject)[0] == CAR and rule(CAR):
        return CAR, f"subject '{best_match(subject)[1]}'"          # 5.18.0: Metcash's activity report PDF is titled Statement
    if has_oc and typ in (None, "Invoice") and NOT_INVOICE.search(page_text or "") \
            and not re.search(r"\btax\s+invoice\b", NOT_INVOICE.sub(" ", header), re.IGNORECASE):   # a TAX INVOICE title always wins
        return OC, "the page says it is not an invoice"          # 'THIS IS NOT AN INVOICE', 'Indicative total' (Fresho); statements keep their type
    if typ:
        return typ, f"title '{kw}'"
    if docket_in_text(page_text) and rule("POD") and not TAX_INVOICE.search(header):
        return "POD", "the page says it is a delivery docket"
    if has_oc and ORDER_TITLE.search(page_text or "") and not INVOICE_TITLE.search(page_text or "") \
            and not re.search(r"\binvoice\b", page_text or "", re.IGNORECASE):      # the word 'invoice' anywhere: not decided here
        return OC, "the page is an order (Sales Order / Order #), not a tax invoice"
    typ, kw = best_match(subject)
    if typ:
        return typ, f"subject '{kw}'"
    stem = re.sub(r"\.[A-Za-z0-9]{2,4}$", "", filename or "")
    typ, kw = best_match(stem)
    if typ:
        return typ, f"file name '{kw}'"
    if sender and sender.get("Default document type"):
        return sender["Default document type"], "exception sender"
    if sender and settings.yes(sender.get("Always download", "")):
        return "Invoice", "exception sender (default)"
    return None, ""


DEFAULT_LAYOUTS = {"statement": "{MONTH}\\{COMPANY}", "order confirmation": "{MONTH}", "direct debit notice": "{COMPANY}",
                   "account activity listing": "{SITE}", "fleet card statement": "{SITE}"}


def layout_parts(layout):
    """'{MONTH}\\{COMPANY}' -> ['{MONTH}', '{COMPANY}']"""
    return [p.strip() for p in re.split(r"[\\/]+", layout or "") if p.strip()]


def rule(type_name):
    """The DocTypes row for a type as a dict with parsed fields, or None."""
    rows = list(settings.rows("DocTypes"))
    have = {r["Document type"].strip().lower() for r in rows}
    for d in settings.DOCTYPES:                              # a Bot_Settings.xlsx made before a type existed still knows it
        if d[0].lower() not in have and d[0] in ("Order Confirmation", "Direct Debit Notice", "Account Activity Listing", "Fleet Card Statement"):
            rows.append(dict(zip(settings.SHEETS["DocTypes"][0], d)))
    for r in rows:
        if r["Document type"].strip().lower() == str(type_name).strip().lower():
            kind = r["Route"].strip().lower()
            route = ("unrecognised" if kind.startswith("unrec") else "holding" if kind.startswith("hold")
                     else "mailbox" if kind.startswith("mailbox") else "company")
            layout = (r.get("Folder layout") or "").strip() or DEFAULT_LAYOUTS.get(r["Document type"].strip().lower(), "{MONTH}")
            return {
                "type": r["Document type"].strip(), "route": route, "subfolder": r["Subfolder"].strip(),
                "read_amounts": settings.yes(r["Read amounts"]), "sign": r["Sign"].strip(),
                "required": [x.strip() for x in r["Required fields"].split(",") if x.strip()],
                "tag": r["File tag"].strip(), "layout": layout,
                "pattern": ("{supplier}_{site}_{date}_STM" if r["Name pattern"].strip() == settings.OLD_STATEMENT_PATTERN
                            else r["Name pattern"].strip() or "{supplier}_{number}"),
            }
    return None
