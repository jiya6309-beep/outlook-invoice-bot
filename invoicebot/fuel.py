"""Fuel invoices (mailbox P I Fuel Invoice / pifuelinvoice@): one fixed layout from Pearl Investment Family Trust.

What is taken from each invoice:
    Customer        "Highland Energy 1 Pty Ltd - PEARL ENERGY TARAGO"   (the Payment Advice 'Customer' field)
      company part  before " - "  -> the company folder (Companies sheet)
      site part     after  " - "  -> the file name, exactly as printed
    Invoice Number  INV-115316   -> file name + duplicate check
    Invoice Date    19 Oct 2026
    Reference       G2617314B-AZ679   (only copied into the Excel list)

Words are read WITH their positions, so a value that wraps onto a second line is joined correctly
("...-SPLIT-" + "MORISSET" -> "...-SPLIT-MORISSET", the hyphen is kept)."""
import datetime
import re

import pdfplumber

INV = re.compile(r"\bINV-\d{3,}\b", re.IGNORECASE)
DATE = re.compile(r"\b(\d{1,2})\s+([A-Za-z]{3})[a-z]*\s+(\d{4})\b")
MONTHS = {m: i for i, m in enumerate(["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], 1)}
SUPPLIER = "Pearl Investment Family Trust"


def _date(text):
    m = DATE.search(text or "")
    if not m or m.group(2).lower()[:3] not in MONTHS:
        return None
    try:
        return datetime.date(int(m.group(3)), MONTHS[m.group(2).lower()[:3]], int(m.group(1)))
    except ValueError:
        return None


def _lines(words, tol=3):
    """Group words into lines by their top position."""
    out = []
    for w in sorted(words, key=lambda w: (round(w["top"]), w["x0"])):
        if out and abs(out[-1][0] - w["top"]) <= tol:
            out[-1][1].append(w)
        else:
            out.append([w["top"], [w]])
    return [(top, sorted(ws, key=lambda w: w["x0"])) for top, ws in out]


def _value_below(words, label, max_lines=3):
    """The text under a label ('Reference', 'Invoice Date'): words that start in the label's column, line by line,
    joined; a line ending in '-' is joined without a space."""
    first = label.split()[0]
    hits = [w for w in words if w["text"] == first]
    for h in hits:
        if label != first:
            nxt = [w for w in words if abs(w["top"] - h["top"]) < 3 and w["x0"] > h["x1"] and w["x0"] - h["x1"] < 12]
            if not nxt or not label.endswith(nxt[0]["text"]):
                continue
        below = [w for w in words if h["top"] + 4 < w["top"] < h["top"] + 14 * (max_lines + 1)
                 and abs(w["x0"] - h["x0"]) < 6]
        text, last_top = "", None
        for w in sorted(below, key=lambda w: w["top"]):
            if last_top is not None and w["top"] - last_top > 16:
                break
            text = text + w["text"] if (not text or text.endswith("-")) else text + " " + w["text"]
            last_top = w["top"]
        if text:
            return text.strip()
    return ""


def _customer(words):
    """The Payment Advice 'Customer' value: the words to the right of the label, until the 'Invoice Number' row."""
    labels = [w for w in words if w["text"] == "Customer"]
    for lab in labels:
        same = sorted([w for w in words if abs(w["top"] - lab["top"]) < 3 and w["x0"] > lab["x1"]], key=lambda w: w["x0"])
        if not same:
            continue
        x_val = same[0]["x0"]
        stop = [w["top"] for w in words if w["text"] == "Invoice" and abs(w["x0"] - lab["x0"]) < 6 and w["top"] > lab["top"] + 2]
        bottom = min(stop) - 1 if stop else lab["top"] + 30
        part = [w for w in words if lab["top"] - 3 <= w["top"] < bottom and w["x0"] >= x_val - 2]
        text = " ".join(" ".join(w["text"] for w in ws) for _top, ws in _lines(part))
        if text:
            return re.sub(r"\s+", " ", text).strip()
    return ""


MARKERS = ("pearlinvestmentfamilytrust", "44948199172")      # its name / ABN: how a fuel invoice is told apart in ANY mailbox (5.19.0)


def parse_if_fuel(path):
    """parse(), but only for a page that names Pearl Investment Family Trust (or its ABN) - used outside the Fuel invoices
    mailbox, where an ordinary supplier invoice with 'Customer' and 'INV-1234' must never be taken for a fuel invoice."""
    try:
        with pdfplumber.open(path) as pdf:
            text = (pdf.pages[0].extract_text() or "") if pdf.pages else ""
    except Exception:
        return None
    from . import settings
    if not any(m in settings.norm(text) for m in MARKERS):
        return None
    return parse(path)


def parse(path):
    """Returns the fields of a fuel invoice, or None when the file is not one (any other PDF in that mailbox follows
    the normal rules)."""
    try:
        with pdfplumber.open(path) as pdf:
            if not pdf.pages:
                return None
            page = pdf.pages[0]
            text = page.extract_text() or ""
            words = page.extract_words(keep_blank_chars=False, use_text_flow=False)
    except Exception:
        return None
    if "customer" not in text.lower() or not INV.search(text):
        return None
    customer = _customer(words)
    if " - " not in customer:
        return None
    company_text, site_text = (p.strip() for p in customer.split(" - ", 1))
    number = INV.search(text).group(0).upper()
    date = _date(_value_below(words, "Invoice Date")) or _date(text)
    return {"customer": customer, "company_text": company_text, "site_text": site_text, "number": number,
            "date": date, "reference": _value_below(words, "Reference"), "supplier": SUPPLIER}
