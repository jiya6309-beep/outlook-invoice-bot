"""The earlier, simpler reader (labels followed directly by an amount). Kept as a safety net:
if the label reader cannot read an invoice, this one gets a second try."""
import re

from .parsing import DATE_TOKEN, parse_date, to_number, find

_MONEY = r"(?:A?\$)?\s*(-?[\d,]+\.\d{2})"
AMOUNT_LABELS = [
    r"invoice\s+total", r"total\s+amount\s+(?:due|payable)", r"total\s+due",
    r"amount\s+due", r"grand\s+total", r"(?<![a-z])total", r"balance\s+due",
]
GST_LABELS = [r"total\s+gst", r"gst\s+total", r"gst\s+amount", r"total\s+tax", r"gst", r"(?<![a-z])tax(?!\s*invoice)"]
SUBTOTAL_LABELS = [
    r"sub\s*-?\s*total", r"net\s+line\s+total", r"net\s+total",
    r"total\s+(?:ex|excl\.?|excluding)\s*gst", r"total\s+before\s+gst",
]
NOT_THESE_BEFORE = {"gst", "sub", "net", "line", "ex", "excl"}


def find_money(labels, text):
    for label in labels:
        pattern = label + r"\s*(?:\([^)]*\))?\s*(?:\d+(?:\.\d+)?\s*%)?\s*[:\-]?\s*" + _MONEY
        found = None
        for m in re.finditer(pattern, text, re.IGNORECASE):
            same_line = text[:m.start()].split("\n")[-1].lower()
            prev = re.search(r"([a-z]+)[ \t]*$", same_line)
            if prev and prev.group(1) in NOT_THESE_BEFORE:
                continue
            found = to_number(m.group(1))
        if found is not None:
            return found
    return None


def find_invoice_number(text):
    has_digit = r"((?=[A-Z0-9\-/]*\d)[A-Z0-9][A-Z0-9\-/]{2,})"
    for pattern in (r"invoice\s*(?:no\.?|number|num|#)\s*[:#.\-]?\s*" + has_digit,
                    r"(?:credit\s+(?:note|memo)|adjustment\s+note|document|doc)\s*(?:no\.?|number|#)\s*[:#.\-]?\s*" + has_digit,
                    r"\binvoice\s+" + has_digit,
                    # label on one line, its value at the START of the next line (columns side by side: "Invoice # | Invoice To | Supplier" /
                    # "17112 | HONEY SHAH | address"): only when the label line carries no number of its own
                    r"(?m)^[ \t]*invoice[ \t]*(?:no\.?|number|num|#)[^\n\d]*\n[ \t]*" + has_digit + r"\b"):
        value = find(pattern, text)
        if value and parse_date(value) is None:
            return value
    return title_number(text) or crm_number(text)


ABN_GROUPS = re.compile(r"\b\d{2}\s?\d{3}\s?\d{3}\s?\d{3}\b")


def title_number(text):
    """No 'Invoice No' label at all: the number printed right after the 'Tax Invoice' title (often below it, or stuck to the ABN line).
    4-10 digits, not part of the ABN, not a date, within the next 4 lines."""
    lines = text.splitlines()
    for i, line in enumerate(lines[:12]):
        if re.match(r"^\s*tax\s+invoice\s*$", line, re.IGNORECASE):
            for nxt in lines[i + 1:i + 5]:
                rest = ABN_GROUPS.sub(" ", nxt)
                if re.search(r"[A-Za-z]{3,}", re.sub(r"(?i)\babn\b", " ", rest)):
                    continue
                m = re.search(r"(?<![\d.,/-])(\d{4,10})(?![\d.,/-])", rest)
                if m and parse_date(m.group(1)) is None:
                    return m.group(1)
    return None


def crm_number(text):
    """Credit notes of some suppliers carry only a 'CRM ID'. Used last: a real invoice number always wins."""
    return find(r"\bcrm\s*id\s*[:#.\-]?\s*((?=[A-Z0-9\-/]*\d)[A-Z0-9][A-Z0-9\-/]{2,})", text)


def find_date(text):
    value = find(r"(?<!due )date\s*[:\-]?\s*(" + DATE_TOKEN + ")", text)
    if parse_date(value):
        return parse_date(value)
    for m in re.finditer(DATE_TOKEN, text):
        if parse_date(m.group()):
            return parse_date(m.group())
    return None


def check_numbers(amount, gst, subtotal):
    if amount is None or gst is None:
        return None
    a, g = abs(amount), abs(gst)
    if subtotal is not None and abs(abs(subtotal) + g - a) <= 0.05:
        return None
    if subtotal is None and 0 <= g <= a / 11 + 0.05:
        return None
    return "totals do not add up (subtotal + GST is not the total)"


def read(text):
    """Number and date are required; the three amounts are optional and are kept exactly as printed (nothing is calculated or checked)."""
    result = {"invoice_number": find_invoice_number(text), "invoice_date": find_date(text),
              "amount": find_money(AMOUNT_LABELS, text), "gst": find_money(GST_LABELS, text),
              "ex_amount": find_money(SUBTOTAL_LABELS, text), "notes": [], "blocks": [], "method": "legacy"}
    problems = []
    if result["invoice_number"] is None:
        problems.append("invoice number not found")
    if result["invoice_date"] is None:
        problems.append("invoice date not found")
    result["problems"] = problems
    return result
