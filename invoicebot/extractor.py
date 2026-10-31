"""Reads ONE document file.  analyse(path, ctx) answers:

    who issued it (supplier)   what it is (invoice / credit note / POD / statement / ...)
    its number, date, GST and total (only for the types that have them)
    which company and site it is for    and   what is wrong with it (problems / warnings)

Order of work: open the file -> supplier -> document type -> read the fields that type needs
-> cross-checks.  Anything that does not add up becomes a PROBLEM; the bot never records a guess.
"""
import datetime
import hashlib
import os
import re

from . import companies, doctypes, labelreader, legacy, naming, ocr, pdftext, settings, sites, suppliers, xlsxdoc
from .parsing import find, parse_date, to_number

FUTURE_TOLERANCE_DAYS = 3        # an invoice dated further ahead than this is probably a misread date
SCANNED_MIN_CHARS = 40           # fewer characters on page 1 = a scanned picture
SPECIAL_READERS = [(re.compile(r"^PFD\b", re.IGNORECASE), "pfd")]   # suppliers whose totals box needs its own reader


def _blank():
    return dict(file_hash="", page_count=0, ocr_used=False, fatal=False,
                supplier="", supplier_printed="", supplier_source="", short="", folder_override="",
                doc_type=None, type_source="", invoice_number=None, invoice_date=None, amount=None, gst=None, ex_amount=None, remarks="",
                doc_company="", site=settings.UNASSIGNED, site_source="", site_company="", address_text="",
                problems=[], warnings=[], method="", site_text="", site_blocks=[])


def file_hash(path):
    h = hashlib.sha1()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


# ----------------------------------------------------------------------------------------
# cross-checks
# ----------------------------------------------------------------------------------------

def check_date(d):
    if settings.run_value("Month folder follows", "window").lower() == "window":
        return None                                   # a future date is handled by the month rules (it waits in the current month)
    if d and d > datetime.date.today() + datetime.timedelta(days=FUTURE_TOLERANCE_DAYS):
        return f"invoice date {d} is in the future - check the date"
    return None


def subject_warnings(subject, invoice_number, text):
    """The e-mail subject often holds the invoice number: use it as an independent check."""
    if not subject or not invoice_number:
        return []
    norm = lambda s: re.sub(r"[^a-z0-9]", "", s.lower())
    tokens = [t for t in re.findall(r"[A-Za-z0-9][A-Za-z0-9\-/]{3,}", subject)
              if re.search(r"\d", t) and parse_date(t) is None]
    mine = norm(invoice_number)
    if any(mine in norm(t) or norm(t) in mine for t in tokens):
        return []
    others = [t for t in tokens if len(norm(t)) >= 5 and norm(t) in norm(text)]
    if others:
        return [f"email subject mentions {others[0]}, but the invoice number read is {invoice_number}"]
    return []


def hint_invoice_number(hints, text):
    """The subject / file name often holds the invoice number. Accept it only if it is also in the PDF."""
    if not hints:
        return None
    norm = lambda s: re.sub(r"[^a-z0-9]", "", s.lower())
    body = norm(text)
    hints = re.sub(r"[_]+", " ", hints)                      # 'Account_9014462' must still read 'Account 9014462'
    cleaned = re.sub(r"\b(?:acc|account|po|p\.o\.|ref|customer|cust|abn)\b[\s:#.\-]*\S+", " ", hints, flags=re.IGNORECASE)
    first = re.search(r"\b(?:invoice|inv)\b[\s:#.\-]*(?:no\.?|number)?[\s:#.\-]*([A-Za-z0-9][A-Za-z0-9\-/]{3,})",
                      cleaned, re.IGNORECASE)
    tokens = ([first.group(1)] if first else []) + re.findall(r"[A-Za-z0-9][A-Za-z0-9\-/]{4,}", cleaned)
    for t in tokens:
        t = t.strip("-/")
        if re.search(r"\d", t) and parse_date(t) is None and len(norm(t)) >= 6 and norm(t) in body:
            return t
    return None


def apply_hints(core, text, hints):
    if core["invoice_number"] is None:
        number = hint_invoice_number(hints, text)
        if number:
            core["invoice_number"] = number
            core["problems"] = [p for p in core["problems"] if p != "invoice number not found"]
            core["notes"].append(f"invoice number {number} was taken from the email subject / file name "
                                 "(the same number is printed in the PDF)")


# ----------------------------------------------------------------------------------------
# the field readers
# ----------------------------------------------------------------------------------------

def read_pfd(text):
    core = {
        "invoice_number": find(r"Invoice\s*No\s+([A-Z0-9]+)", text) or legacy.find_invoice_number(text),
        "invoice_date": (parse_date(find(r"Invoice\s*Date\s+(\d{1,2}\s+[A-Za-z]{3}\s+\d{2,4})", text))
                         or legacy.find_date(text)),
        "amount": to_number(find(r"(?:INVOICE|CREDIT)\s+TOTAL\s*\(GST\s*Incl\.?\)\s*\$?\s*(-?[\d,]+\.\d{2})", text)),
        "gst": to_number(find(r"TOTAL\s+GST\s*\$?\s*(-?[\d,]+\.\d{2})", text)),
        "notes": [], "blocks": [], "method": "pfd",
    }
    names = {"invoice_number": "invoice number not found", "invoice_date": "invoice date not found"}
    core["problems"] = [msg for key, msg in names.items() if core[key] is None]     # the amounts are optional
    core["ex_amount"] = None
    return core


def _drop_impossible_gst(core):
    """A GST that is larger than a 10% share of the total cannot be the GST of this document (it was a rate, a line or a column):
    leave it blank rather than record a wrong figure. Nothing is calculated."""
    amount, gst = core.get("amount"), core.get("gst")
    if amount is not None and gst is not None and abs(gst) > abs(amount) / 11 + 0.10:
        core["gst"] = None
        core["notes"].append("a GST figure that cannot belong to this total was ignored")


def read_amounts(doc, printed, supplier, hints):
    """Invoice / credit note: number, date, GST, total - with the arithmetic checks."""
    labels = settings.labels_for(printed, supplier)
    pfd_problems = []
    if re.match(r"^PFD\b", printed or "", re.IGNORECASE):
        core = read_pfd(doc.text)
        core["blocks"] = labelreader._read_blocks(doc, labels["site"])
        if not core["problems"]:
            return core
        pfd_problems = core["problems"]              # a layout the PFD reader does not know: try the general readers too
    primary = labelreader.read(doc, labels)
    apply_hints(primary, doc.text, hints)
    if not primary["problems"]:
        if primary["amount"] is None and primary["gst"] is None and primary["ex_amount"] is None:
            second = legacy.read(doc.text)                        # nothing found by label: let the second reader try (as printed)
            if second["amount"] is not None:                      # the fallback reader is trusted only when it found the TOTAL
                primary.update(amount=second["amount"], gst=second["gst"], ex_amount=second["ex_amount"])
                primary["notes"].append("amounts read by the fallback reader")
        _drop_impossible_gst(primary)
        return primary
    fallback = legacy.read(doc.text)
    apply_hints(fallback, doc.text, hints)
    if not fallback["problems"]:
        fallback["notes"].append("read by the fallback reader; label reader said: " + "; ".join(primary["problems"]))
        fallback["blocks"] = primary["blocks"]
        for key in ("amount", "gst", "ex_amount"):                # what only the first reader found is kept (as printed)
            if fallback.get(key) is None and primary.get(key) is not None:
                fallback[key] = primary[key]
        if fallback.get("amount") is None:
            fallback["gst"] = fallback["ex_amount"] = None
        _drop_impossible_gst(fallback)
        return fallback
    _drop_impossible_gst(primary)
    primary["problems"] = primary["problems"] + ["fallback reader also failed: " + "; ".join(fallback["problems"])]
    if pfd_problems:
        primary["problems"].append("PFD reader: " + "; ".join(pfd_problems))
    return primary


# ----------------------------------------------------------------------------------------
# the main entry point
# ----------------------------------------------------------------------------------------

GST_FREE = re.compile(r"\bGST[\s\-]*(?:free|exempt)\b|\bno\s+GST\b|\bfree\s+of\s+GST\b", re.IGNORECASE)


def _is_encrypted(path):
    """True if the PDF carries an /Encrypt entry (they are in the file's trailer, near the end)."""
    try:
        size = os.path.getsize(path)
        with open(path, "rb") as f:
            head = f.read(4096)
            f.seek(max(size - 8192, 0))
            tail = f.read()
        return b"/Encrypt" in head or b"/Encrypt" in tail
    except OSError:
        return False


def _fatal(res, reason):
    res["fatal"] = True
    res["problems"].append(reason)
    return res


def analyse(path, ctx=None):
    """Read one file. `ctx` (all optional): subject, sender_name, sender_email, mailbox, is_forward,
    orig_name, to_name, extra_known (known supplier names), keep_text."""
    ctx = ctx or {}
    res = _blank()
    ext = os.path.splitext(path)[1].lower()
    orig = ctx.get("orig_name") or os.path.basename(path)
    subject = ctx.get("subject", "")
    hints = f"{subject} {orig}"
    try:
        res["file_hash"] = file_hash(path)
    except OSError as e:
        return _fatal(res, f"could not read the file: {e}")

    # 1. open the file (page 1 first - enough to identify most documents)
    scanned = False
    try:
        if ext == ".pdf":
            doc1 = pdftext.load(path, max_pages=1)
            res["page_count"] = doc1.page_count
            if len(doc1.text.strip()) < SCANNED_MIN_CHARS:
                scanned = True
        elif ext in (".xlsx", ".xlsm"):
            doc1 = xlsxdoc.load(path)
            res["page_count"] = 1
        else:
            return _fatal(res, f"file type {ext or '(none)'} cannot be read")
    except Exception as e:
        if ext == ".pdf" and _is_encrypted(path):
            return _fatal(res, "the PDF is password-protected: open it with the password and save an unprotected copy")
        detail = str(e)[:120].strip()
        return _fatal(res, f"could not open the file ({type(e).__name__})" + (f": {detail}" if detail else ""))

    if scanned:
        try:
            doc1 = ocr.ocr_document(path)
            res["ocr_used"] = True
            res["warnings"].append("scanned PDF read by OCR - please check the values")
        except Exception as e:
            return _fatal(res, f"scanned PDF and OCR is not available: {e}")
        if len(doc1.text.strip()) < 15:
            return _fatal(res, "scanned PDF: no text could be read")

    header = doc1.header(15)
    try:                                                     # the same top of the page as positioned lines: finds titles the plain text scrambles
        header += "\n" + "\n".join(l.text for l in doc1.lines[:20] if l.page == 0)
    except Exception:
        pass
    sender = doctypes.sender_row(ctx.get("sender_name", ""), ctx.get("sender_email", ""))
    if ctx.get("keep_text"):
        res["_text"] = doc1.text

    # 2. supplier
    logo = (lambda: ocr.logo_text(path)) if ext == ".pdf" and not scanned else None
    printed, src = suppliers.identify(doc1.text, sender, logo, ctx.get("extra_known", ()), ctx.get("sender_name", ""),
                                      sender_email=ctx.get("sender_email", ""), is_forward=bool(ctx.get("is_forward")))
    printed = naming.tidy(printed)
    if printed and src in ("printed", "logo") and not suppliers.plausible(printed):
        res["warnings"].append(f"'{printed[:60]}' was found where the supplier should be, but it does not look like a name")
        printed, src = None, ""
    if printed:
        record, short, override = settings.canonical_supplier(printed)
        res.update(supplier=record, supplier_printed=printed, supplier_source=src, short=short, folder_override=override)

    # 3. document type
    dtype, how = doctypes.decide(subject, orig, header, sender, doc1.text, printed or "")
    rule = doctypes.rule(dtype) if dtype else None
    res["type_source"] = how
    if not rule:
        res["problems"].append("document type not recognised" if not dtype else f"document type '{dtype}' is not in the DocTypes sheet")
    else:
        res["doc_type"] = rule["type"]

    labels = settings.labels_for(printed or "", res["supplier"])
    core = None
    # 4. read what this type needs
    if rule and rule["read_amounts"]:
        max_pages = settings.run_int("Max pages (invoice types)", 15)
        if ext == ".pdf" and not res["ocr_used"] and res["page_count"] > max_pages:
            res["problems"].append(f"too many pages ({res['page_count']} > {max_pages}); not read")
        else:
            doc = doc1
            if ext == ".pdf" and not res["ocr_used"] and res["page_count"] > 1:
                try:
                    doc = pdftext.load(path)
                except Exception as e:
                    return _fatal(res, f"could not read all pages: {str(e)[:100]}")
            core = read_amounts(doc, printed or "", res["supplier"], hints)
            if not printed:
                core["problems"] = [p for p in core["problems"]]
    elif rule:
        pod_labels = dict(labels, invoice_no=list(labels["invoice_no"]) + ["del docket", "delivery docket", "docket", "delivery note no",
                                                                           "delivery note", "grn no", "grn"])
        basic = labelreader.read_basic(doc1, pod_labels)
        core = {"invoice_number": basic["invoice_number"], "invoice_date": basic["invoice_date"], "amount": None,
                "gst": None, "problems": [], "notes": [], "blocks": basic["blocks"], "method": "basic"}
        apply_hints(core, doc1.text, hints)
        core["problems"] = []                                  # number/date are nice to have for these types
        if "number" in rule["required"] and not core["invoice_number"]:
            res["warnings"].append("no document number found; the original file name is used")
        if core["invoice_date"] is None:
            res["warnings"].append("no document date found; the e-mail's arrival date is used")
    if core is not None:
        res.update(invoice_number=naming.tidy(core["invoice_number"]), invoice_date=core["invoice_date"],
                   amount=core["amount"], gst=core["gst"], ex_amount=core.get("ex_amount"), method=core["method"])
        res["problems"] += core["problems"]
        res["warnings"] += core["notes"]
        res["warnings"] += subject_warnings(subject, core["invoice_number"], doc1.text)
        dp = check_date(core["invoice_date"])
        if dp:
            res["problems"].append(dp)
        if rule["read_amounts"]:
            if res["gst"] is None and res["amount"] is not None and GST_FREE.search(doc1.text) \
                    and (res["ex_amount"] is None or abs(abs(res["ex_amount"]) - abs(res["amount"])) < 0.005):
                res["gst"] = 0.0                                  # the document itself says GST free: 0.00 (never for statements)
            sign = {"-": -1, "+": 1}.get(rule["sign"])
            if sign:                                              # credit notes negative, invoices positive: each figure on its own
                for key in ("amount", "gst", "ex_amount"):
                    if res[key] is not None:
                        res[key] = sign * abs(res[key])
            if res["amount"] is None and res["gst"] is None and res["ex_amount"] is None:
                res["remarks"] = "GST NOT FOUND"                  # only when ALL THREE are missing
                if (settings.run_bool("Hold documents without amounts when the type came only from the subject", True)
                        and not res["type_source"].startswith("title")):
                    res["problems"].append("no amounts found, and the page itself does not say it is an invoice "
                                           f"(the type came from the e-mail: {res['type_source']})")
        blocks = core.get("blocks", [])
    else:
        blocks = []

    # 5. company and site named on the document
    found = companies.find_in_text(doc1.text)
    if found:
        res["doc_company"] = found[0][0]["Company name"]
        if len(found) > 1:
            res["warnings"].append("several companies are named on the document; used " + found[0][0]["Company name"])
    m = sites.match(doc1.text, subject=subject, blocks=blocks, supplier=res["supplier"], to_name=ctx.get("to_name", ""))
    res.update(site=m["site"], site_source=m["source"], site_company=m["company"])
    res["address_text"] = (sites.split_blocks(blocks)[1][0] if blocks else "")[:160]
    res["site_text"], res["site_blocks"] = doc1.text[:8000], [b[:300] for b in blocks[:6]]     # for the second look: the site WITHIN the company
    return res


def extract_document(path, mailbox=None, is_forward=False, subject=""):
    """Compatibility helper for the scorecard / diagnose tools: returns (data, first-page text)."""
    data = analyse(path, {"subject": subject, "mailbox": mailbox, "is_forward": is_forward, "keep_text": True})
    data["method"] = data.get("method") or "-"
    return data, data.pop("_text", "")
