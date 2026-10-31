"""Finds invoice fields by LABEL (e.g. "Document No", "Total inc GST"), looking
  1. to the right of the label on the same line, and
  2. in the cell directly below it, when the label is a table heading.
Then it checks the arithmetic, so a wrong reading is rejected instead of recorded."""
import datetime
import re

from .parsing import DATE_TOKEN, MONEY_TOKEN, parse_date, parse_money

TOLERANCE = 0.10          # dollars of rounding difference we accept between lines and totals

SEP = r"[\s.:()/\-#]*"
_INC = r"(?:inc|incl|including|inclusive)"
_EXC = r"(?:ex|exc|excl|excluding|exclusive)"
EQUIV = {"inc": _INC, "incl": _INC, "including": _INC, "inclusive": _INC,
         "ex": _EXC, "exc": _EXC, "excl": _EXC, "excluding": _EXC, "exclusive": _EXC}

# A one-word label such as "total" or "gst" is ambiguous. Skip it when the word right before it
# turns it into a different label ("GST total", "Total excl GST", "Due date").
SKIP_BEFORE = {
    "total": {"gst", "sub", "net", "line", "taxable", "tax", "ex", "exc", "excl", "excluding"},
    "gst": {"ex", "exc", "excl", "excluding", "exclusive", "inc", "incl", "including",
            "inclusive", "for", "sales", "free", "taxable", "net", "before", "less"},
    "date": {"due", "delivery", "order", "dispatch", "ship", "shipping", "payment",
             "print", "printed", "expiry", "statement"},
}
CLOSE_GAP = 15            # points: a previous word closer than this belongs to the same label
LABEL_WORD_GAP = 20       # points: the words of one label must sit this close together

_PAREN = r"\((?!\s*-?\s*(?:AUD|A?\$)?\s*\d[\d,]*\.\d{2})[^)]*\)\s*"
MONEY_AFTER = re.compile(r"^(?:[\s:._]|-\s)*(?:" + _PAREN + r")*(?:\d+(?:\.\d+)?\s*%\s*)?("
                         + MONEY_TOKEN + ")", re.IGNORECASE)
DATE_AFTER = re.compile(r"^[\s:._\-]*(" + DATE_TOKEN + ")")
INVOICE_AFTER = re.compile(r"^[\s:#.\-]*(?:(?:no|number|num|nbr)\b[\s:#.\-]*)?"
                           r"((?=[A-Za-z0-9\-/]*\d)[A-Za-z0-9][A-Za-z0-9\-/]{2,24})")

# A heading row such as "Subtotal | GST | Total" is a summary. A row such as "Qty | Description | Unit
# Price | Line Total" is an item table, whose first row of numbers is an ITEM, not the invoice total.
SUMMARY_WORDS = {
    "sub", "subtotal", "total", "gst", "tax", "taxes", "net", "amount", "amounts", "freight", "discount",
    "levy", "balance", "due", "inc", "incl", "including", "inclusive", "ex", "excl", "excluding",
    "exclusive", "payable", "price", "taxable", "value", "aud", "credit", "included", "invoice",
    "goods", "charges", "fee", "fees", "surcharge", "non", "before", "after", "cdl", "of", "and",
}

_regex_cache = {}


def _tokens(label):
    return re.findall(r"[a-z0-9]+", label.lower())


def label_regex(label):
    key = label.lower()
    if key not in _regex_cache:
        tokens = _tokens(key)
        if not tokens:
            _regex_cache[key] = None
        else:
            body = SEP.join(EQUIV.get(t, re.escape(t)) for t in tokens)
            _regex_cache[key] = re.compile(r"(?<![a-z0-9])" + body + r"(?![a-z])", re.IGNORECASE)
    return _regex_cache[key]


def _occurrences(doc, label):
    rx = label_regex(label)
    tokens = _tokens(label)
    if rx is None:
        return
    single = tokens[0] if len(tokens) == 1 else None
    for li, line in enumerate(doc.lines):
        for m in rx.finditer(line.text):
            if line.max_gap(m.start(), m.end()) > LABEL_WORD_GAP:
                continue                      # words from two different table columns, not one label
            if single in SKIP_BEFORE:
                prev, gap = line.prev_word_gap(m.start())
                if prev and gap <= CLOSE_GAP and re.sub(r"[^a-z]", "", prev.lower()) in SKIP_BEFORE[single]:
                    continue
            yield li, line, m


def _below_lines(doc, li, count=2, max_gap=45):
    """The next few lines on the same page, each close to the one before."""
    out, prev = [], doc.lines[li]
    for nxt in doc.lines[li + 1: li + 1 + count]:
        if nxt.page != prev.page or nxt.top - prev.top > max_gap:
            break
        out.append(nxt)
        prev = nxt
    return out


# ------------------------------- the four kinds of value -------------------------------

def _is_summary_heading(line):
    words = re.findall(r"[A-Za-z]+", line.text.lower())
    return bool(words) and all(w in SUMMARY_WORDS for w in words)


def _money_words(line):
    return [w for w in line.words if re.fullmatch(MONEY_TOKEN, w["text"], re.IGNORECASE)]


def _money_value(doc, li, line, m):
    rem = line.text[m.end():]
    hit = MONEY_AFTER.match(rem)
    if hit:
        return parse_money(hit.group(1))
    if re.search(MONEY_TOKEN, rem):               # other amounts on this line: not a table heading
        return None
    if not _is_summary_heading(line):             # an item table: the numbers below are items
        return None
    x0, x1 = line.x_range(m.start(), m.end())
    cx = (x0 + x1) / 2
    prev = line
    for k in range(li + 1, min(li + 3, len(doc.lines))):
        nxt = doc.lines[k]
        if nxt.page != prev.page or nxt.top - prev.top > 45:
            break
        best = None
        for w in _money_words(nxt):
            if w["x1"] < x0 - 25 or w["x0"] > x1 + 25:
                continue
            dist = min(abs((w["x0"] + w["x1"]) / 2 - cx), abs(w["x1"] - x1))
            if best is None or dist < best[0]:
                best = (dist, w)
        if best:
            word = best[1]
            after = doc.lines[k + 1] if k + 1 < len(doc.lines) else None
            if after is not None and after.page == nxt.page and after.top - nxt.top <= 30:
                if any(a["x1"] >= word["x0"] - 10 and a["x0"] <= word["x1"] + 10 for a in _money_words(after)):
                    return None                   # more numbers in the same column: a table, not a summary
            return parse_money(word["text"])
        prev = nxt
    return None


def _valid_invoice_number(token, min_len=3):
    token = token.strip("-/")
    return (len(token) >= min_len and re.search(r"\d", token) is not None
            and parse_date(token) is None
            and not re.fullmatch(r"\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4}", token))


def _invoice_value(doc, li, line, m, allow_below):
    rem = line.text[m.end():]
    hit = INVOICE_AFTER.match(rem)
    if hit and _valid_invoice_number(hit.group(1)):
        return hit.group(1).strip("-/")
    if allow_below and not re.search(r"\d", rem):
        x0, x1 = line.x_range(m.start(), m.end())
        cx = (x0 + x1) / 2
        for nxt in _below_lines(doc, li):
            cands = sorted((abs((w["x0"] + w["x1"]) / 2 - cx), w["text"]) for w in nxt.words
                           if w["x1"] >= x0 - 60 and w["x0"] <= x1 + 60 and _valid_invoice_number(w["text"], 4))
            if cands:
                return cands[0][1].strip("-/")
    return None


COMPACT_TAG = re.compile(r"\s*\(ddmmyy\)\s*$", re.IGNORECASE)   # a date label written 'inv date (ddmmyy)' in sheet Labels
COMPACT_DATE = re.compile(r"^[\s:._\-]*(\d{6})(?!\d)")


def parse_ddmmyy(token):
    """'220726' -> 22 July 2026 (only for a label marked (ddmmyy): a bare 6-digit number is otherwise never a date)."""
    try:
        d = datetime.date(2000 + int(token[4:6]), int(token[2:4]), int(token[0:2]))
    except ValueError:
        return None
    return d


def _compact_date_value(doc, li, line, m):
    """The date after or under the label, written as DDMMYY (Peters: INV DATE 220726)."""
    hit = COMPACT_DATE.match(line.text[m.end():])
    if hit:
        return parse_ddmmyy(hit.group(1))
    x0, x1 = line.x_range(m.start(), m.end())
    cx = (x0 + x1) / 2
    for nxt in _below_lines(doc, li):
        cands = sorted((abs((w["x0"] + w["x1"]) / 2 - cx), w["text"]) for w in nxt.words
                       if w["x1"] >= x0 - 30 and w["x0"] <= x1 + 30)
        for _dist, text in cands:
            if re.fullmatch(r"\d{6}", text) and parse_ddmmyy(text):
                return parse_ddmmyy(text)
    return None


def _read_date(doc, labels):
    """The invoice date: a label ending in '(ddmmyy)' reads a compact DDMMYY date, every other label a normal date."""
    for label in labels:
        if COMPACT_TAG.search(label):
            value, _ = _read_first(doc, [COMPACT_TAG.sub("", label)], _compact_date_value)
        else:
            value, _ = _read_first(doc, [label], _date_value)
        if value is not None:
            return value, label
    return None, None


TIME_DAY_DATE = re.compile(r"^[\s:._\-]*(?:\d{1,2}[:.]\d{2}\s*(?:am|pm)?\s*)?(?:(?:mon|tue|wed|thu|fri|sat|sun)[a-z]*\.?,?\s*)?("
                           + DATE_TOKEN + ")", re.IGNORECASE)            # 'Received: 9:46 AM Thu Jul 23, 2026' (Fresho)


def _date_value(doc, li, line, m):
    rem = line.text[m.end():]
    hit = DATE_AFTER.match(rem) or TIME_DAY_DATE.match(rem)
    if hit and parse_date(hit.group(1)):
        return parse_date(hit.group(1))
    if not re.search(DATE_TOKEN, rem):
        x0, x1 = line.x_range(m.start(), m.end())
        cx = (x0 + x1) / 2
        for nxt in _below_lines(doc, li):
            cands = sorted((abs((w["x0"] + w["x1"]) / 2 - cx), i) for i, w in enumerate(nxt.words)
                           if w["x1"] >= x0 - 60 and w["x0"] <= x1 + 60)
            for _dist, i in cands:
                hit = DATE_AFTER.match(nxt.phrase_from(i))
                if hit and parse_date(hit.group(1)):
                    return parse_date(hit.group(1))
    return None


def _unlabelled_date(doc):
    """Last resort: the first date near the top of page 1 that is not a due/delivery/order date."""
    skip = SKIP_BEFORE["date"] | {"from", "to", "valid", "until", "period", "terms", "by"}
    for line in doc.lines[:60]:
        if line.page != 0:
            break
        for m in re.finditer(DATE_TOKEN, line.text):
            d = parse_date(m.group())
            if not d:
                continue
            before = re.findall(r"[a-z]+", line.text[:m.start()].lower())
            if any(w in skip for w in before[-3:]):
                continue
            return d
    return None


# ------------------------- totals ROW under column headings (e.g. Asahi) -------------------------

def _heading_above(doc, li, x0, x1):
    """The column heading above an amount, joining headings that wrap over several lines."""
    page = doc.lines[li].page
    parts, started, last_top = [], False, doc.lines[li].top
    for k in range(li - 1, -1, -1):
        line = doc.lines[k]
        if line.page != page:
            break
        limit = 24 if started else 160
        if last_top - line.top > limit:
            break
        cells = [c for c in line.cells() if c[2] >= x0 - 3 and c[1] <= x1 + 3]
        if not cells:
            if started:
                break
            continue
        text = " ".join(c[0] for c in cells)
        if re.search(r"\d", text):                # numbers: item rows (skip them) - or the end of the heading
            if started:
                break
            continue
        parts.insert(0, text)
        started, last_top = True, line.top
        if len(parts) == 3:
            break
    return " ".join(parts)


def _column_totals(doc, labels):
    """A row such as 'Total AUD 827.86 62.16 890.02 85.47 975.49': each amount is named by the
    heading of its column. Only rows that contain the word Total are used, so item rows are never taken."""
    for li in range(len(doc.lines) - 1, -1, -1):
        line = doc.lines[li]
        money = _money_words(line)
        if len(money) < 3:
            continue
        before = " ".join(w["text"] for w in line.words if w["x1"] <= money[0]["x0"])
        if not re.search(r"\btotals?\b", before, re.IGNORECASE):
            continue
        headings = [(_heading_above(doc, li, w["x0"], w["x1"]), parse_money(w["text"])) for w in money]
        found = {}
        for field in ("total", "gst", "ex"):
            for label in labels[field]:
                if len(_tokens(label)) < 2:        # one-word labels are too vague for headings
                    continue
                rx = label_regex(label)
                value = next((v for h, v in headings if h and rx.search(h)), None)
                if value is not None:
                    found[field] = (value, label)
                    break
        if "total" in found and "gst" in found:
            return found
    return None


def _column_text(words, x_start, first_max=None, max_gap=25):
    """Words from x_start rightwards, stopping at a wide gap (the next column)."""
    out, prev_x1 = [], None
    for w in words:
        if w["x0"] < x_start:
            continue
        if prev_x1 is None and first_max is not None and w["x0"] > first_max:
            return ""
        if prev_x1 is not None and w["x0"] - prev_x1 > max_gap:
            break
        out.append(w["text"])
        prev_x1 = w["x1"]
    return " ".join(out)


def _block_text(doc, li, line, m, max_lines=3):
    x0, x1 = line.x_range(m.start(), m.end())
    parts = [_column_text(line.words, x1 - 1)]
    for nxt in _below_lines(doc, li, count=max_lines):
        parts.append(_column_text(nxt.words, x0 - 10, first_max=x0 + 80))
    return " | ".join(p for p in parts if p)


# ------------------------------- reading a whole document -------------------------------

def _specific_number_label(label):
    """Only labels that clearly name a number ('Document No', 'Reference No') may read the cell UNDER
    them. A bare 'Invoice' / 'Tax Invoice' title may not: whatever is below it is not its value."""
    return "#" in str(label) or bool(set(_tokens(label)) & {"no", "number", "num", "nbr", "ref", "reference", "document", "doc", "transaction"})


def _read_first(doc, labels, fn, specific_only_below=False):
    for label in labels:
        allow_below = _specific_number_label(label)
        for li, line, m in _occurrences(doc, label):
            value = fn(doc, li, line, m, allow_below) if specific_only_below else fn(doc, li, line, m)
            if value is not None:
                return value, label
    return None, None


def _row_total(amounts):
    """Several amounts right after one total label ('Total AUD 1,125.56 94.08 1,219.64 116.80 1,336.44'): the TOTAL is the amount
    that is the sum of two or more of the others (the last one that fits wins: totals rows end with the amount payable).
    (value, True) when the maths proves it; (largest amount, False) when nothing adds up."""
    import itertools
    for i in range(len(amounts) - 1, -1, -1):
        others = amounts[:i] + amounts[i + 1:]
        for size in range(2, min(len(others), 4) + 1):
            if any(abs(sum(combo) - amounts[i]) <= 0.02 for combo in itertools.combinations(others, size)):
                return amounts[i], True
    return max(amounts, key=abs), False


def _amounts_after(line, m):
    return [v for t in re.findall(MONEY_TOKEN, line.text[m.end():], re.IGNORECASE) if (v := parse_money(t)) is not None]


def _read_last_money(doc, labels, notes=None):
    """Totals sit at the end of an invoice, so for amounts the LAST match wins.
    notes (a list) = this is the TOTAL: when a total label is followed by several amounts on its line, the amount payable is
    chosen by the maths (see _row_total), never simply the first one."""
    for label in labels:
        values = []
        for li, line, m in _occurrences(doc, label):
            if notes is not None:
                amounts = _amounts_after(line, m)
                if len(amounts) >= 2:
                    value, proved = _row_total(amounts)
                    values.append((value, None if proved else
                                   f"several amounts after '{label}' and none is the sum of the others: the largest "
                                   f"({value:.2f}) was taken as the total - please check"))
                    continue
            v = _money_value(doc, li, line, m)
            if v is not None:
                values.append((v, None))
        if values:
            value, note = values[-1]
            if note and notes is not None:
                notes.append(note)
            return value, label
    return None, None


SHIP_LABEL = re.compile(r"\b(deliver\w*|ship\w*|consign\w*|dispatch\w*)\b", re.IGNORECASE)
SHIP_TAG = "[SHIP] "          # same text as sites.SHIP_TAG: marks a block read under a ship-to / deliver-to label
SOLD_LABEL = re.compile(r"\bsold\b", re.IGNORECASE)
SOLD_TAG = "[SOLD] "          # same text as sites.SOLD_TAG: a block read under 'Sold to' (used for the site only when there is no ship-to)


def _read_blocks(doc, labels, limit=6):
    blocks = []
    for label in labels:
        for li, line, m in _occurrences(doc, label):
            text = _block_text(doc, li, line, m)
            if text:
                blocks.append((SHIP_TAG if SHIP_LABEL.search(label) else SOLD_TAG if SOLD_LABEL.search(label) else "") + text)
            if len(blocks) >= limit:
                return blocks
    return blocks


def believable(total, gst, ex, ex_label):
    """Is a set of amounts from a 'totals row under column headings' believable as a TOTALS row (and not a line of an item table)?
    The same test as before, but it only decides whether to use the row: nothing is calculated and nothing is reported."""
    if total is None:
        return False
    if gst is None:
        return ex is not None and -TOLERANCE <= abs(total) - abs(ex) <= abs(total) / 11 + 0.06
    return consistent(total, gst, ex, ex_label)


def consistent(total, gst, ex, ex_label):
    """Do the three amounts fit together? Only used to CHOOSE between two readings of the same page; nothing is calculated
    from them and a mismatch is never reported."""
    if total is None or gst is None:
        return False
    t, g = abs(total), abs(gst)
    if ex is not None:
        e = abs(ex)
        if abs(e + g - t) <= TOLERANCE:
            return True
        strict = re.search(r"\b(ex|exc|excl|excluding|exclusive)\b|before gst", (ex_label or "").lower())
        extra = t - e - g
        return bool(not strict and abs(g - e * 0.10) <= max(0.10, e * 0.005) and -TOLERANCE <= extra <= 2 * (e + g))
    return g <= t / 11 + 0.06


def read(doc, labels):
    """labels: dict with lists for site, invoice_no, date, gst, ex, total."""
    invoice_number, _ = _read_first(doc, labels["invoice_no"], _invoice_value, specific_only_below=True)
    invoice_date, _ = _read_date(doc, labels["date"])
    notes = []
    if invoice_date is None:
        invoice_date = _unlabelled_date(doc)
        if invoice_date:
            notes.append(f"invoice date {invoice_date} was taken from an unlabelled date near the top - please check it")

    # The three amounts (total incl. GST, GST, ex-GST) are read as printed. Missing ones stay empty; nothing is calculated.
    total_notes = []
    total, _ = _read_last_money(doc, labels["total"], notes=total_notes)
    gst, _ = _read_last_money(doc, labels["gst"])
    ex, ex_label = _read_last_money(doc, labels["ex"])
    by_label = (total, gst, ex)
    cols = _column_totals(doc, labels)                       # a totals row under column headings
    if cols:
        c_ex, c_label = cols.get("ex", (None, None))
        by_column = (cols["total"][0], cols["gst"][0], c_ex)
        if believable(*by_column, c_label):
            total, gst, ex = by_column
            notes.append("amounts read from the totals row under the column headings")
            total_notes = []
    notes.extend(total_notes)

    problems = []
    if invoice_number is None:
        problems.append("invoice number not found")
    if invoice_date is None:
        problems.append("invoice date not found")
    return {"invoice_number": invoice_number, "invoice_date": invoice_date,
            "amount": total, "gst": gst, "ex_amount": ex, "problems": problems, "notes": notes,
            "blocks": _read_blocks(doc, labels["site"]), "method": "labels"}


def read_basic(doc, labels):
    """Documents without amounts (POD, statements, ...): only the number and the date."""
    invoice_number, _ = _read_first(doc, labels["invoice_no"], _invoice_value, specific_only_below=True)
    invoice_date, _ = _read_date(doc, labels["date"])
    return {"invoice_number": invoice_number, "invoice_date": invoice_date,
            "blocks": _read_blocks(doc, labels["site"])}
