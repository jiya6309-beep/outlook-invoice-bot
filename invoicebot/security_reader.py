"""Reads the weekly alarm reports. Two layouts are known, and one file can hold several sites (each starts with its own
'From ... to ...' header and Client No):

  Basic Event History   'Client No: 0989-25'      12/10/2026 04:01:38 Disarmed: Disarmed Staff 2 7
  Event List            '4894-25 Phone: ...'       28/09/2026 4:00:36 AM 401 2 Disarmed User 2

Only DISARM and ARM events are kept (keyswitch disarms and 'User 0' arms count). Everything else is ignored: Late to
Close / Open, Early to Close, Activation, Picked Up, Test Message, Status, Low Battery, User on Premises ...
Times are used as printed: they are the site's local time. Seconds are dropped."""
import datetime
import re

import pdfplumber

STAMP = re.compile(r"(\d{1,2})/(\d{1,2})/(\d{4})\s+(\d{1,2}):(\d{2}):(\d{2})(?:\s*([AaPp][Mm]))?")
PERIOD = re.compile(r"From\s+(\d{1,2}/\d{1,2}/\d{4})\s+\d{1,2}:\d{2}:\d{2}\s*[AP]M\s+to\s+(\d{1,2}/\d{1,2}/\d{4})", re.IGNORECASE)
CLIENT_LABEL = re.compile(r"Client\s*No:?\s*([0-9A-Z]{3,6})(?:\(\d{1,2}\))?(-[0-9A-Z]{2,3})\b", re.IGNORECASE)   # 9307(1)-SA = area 1 of 9307-SA
CLIENT_PHONE = re.compile(r"(?m)^\s*([0-9]{3,6}-[0-9A-Z]{2,3})\b(?=[^\n]*\bPhone\b)")
NOISE = re.compile(r"activation|picked up|auto completed|late to|early to|test message|status|battery|sleeping|premises|"
                   r"successful send|notification|work order|attend", re.IGNORECASE)


def _date(text):
    d, m, y = (int(x) for x in text.split("/"))
    return datetime.date(y, m, d)


def _stamp(m):
    d, mo, y, h, mi = (int(m.group(i)) for i in range(1, 6))
    ap = (m.group(7) or "").upper()
    if ap:
        h = h % 12 + (12 if ap == "PM" else 0)
    return datetime.datetime(y, mo, d, h, mi)                # seconds dropped on purpose


def classify(description):
    """'disarm', 'arm' or None for the text that follows a time stamp."""
    text = description.strip()
    if NOISE.search(text) and not re.match(r"(?:\d+\s+\d+\s+)?Disarmed\b|Armed\s*:", text):
        return None
    if re.search(r"\bDisarmed\b", text):
        return "disarm"
    if re.search(r"(?<![A-Za-z])Armed\b", text):
        return "arm"
    return None


def _events(text):
    """[(datetime, 'disarm'/'arm', original line)] - every time stamp starts a new entry (page footers glued to the next
    line, '...Report Date 19/10/202615/10/2026 22:09:35 ...', are handled)."""
    out = []
    marks = list(STAMP.finditer(text))
    for i, m in enumerate(marks):
        end = marks[i + 1].start() if i + 1 < len(marks) else len(text)
        desc = text[m.end():end].split("\n")[0]
        if re.match(r"\s*(?:to\s|AM\s+to|PM\s+to)", desc, re.IGNORECASE):   # the 'From ... to ...' header, not an event
            continue
        kind = classify(desc)
        if kind:
            line = (m.group(0) + desc).strip()
            line = re.sub(r"\s+", " ", line)
            out.append((_stamp(m), kind, line[:160]))
    return out


def _client(section):
    m = CLIENT_LABEL.search(section)
    if m:
        return (m.group(1) + m.group(2)).upper()
    m = CLIENT_PHONE.search(section)
    return m.group(1).upper() if m else ""


def _site_lines(section):
    """Site name and address: the first two lines after the 'From ... to ...' header."""
    lines = [l.strip() for l in section.splitlines() if l.strip()]
    for i, l in enumerate(lines):
        if PERIOD.search(l):
            rest = [x for x in lines[i + 1:i + 4] if not x.lower().startswith("phone") and not CLIENT_LABEL.search(x)]
            name = rest[0] if rest else ""
            addr = rest[1] if len(rest) > 1 else ""
            name = re.sub(r"\s*Phone:.*$", "", CLIENT_PHONE.sub("", name)).strip()
            addr = re.sub(r"\s*(Phone|Email):.*$", "", CLIENT_PHONE.sub("", addr)).strip()
            return name, addr
    return "", ""


def read_text(path):
    with pdfplumber.open(path) as pdf:
        return "\n".join((p.extract_text() or "") for p in pdf.pages)


def parse_text(text):
    """[{'client', 'name', 'address', 'start', 'end', 'events': [(dt, kind, line)]}] - one entry per site section."""
    starts = [m.start() for m in PERIOD.finditer(text)]
    if not starts:
        starts = [0]
    sections = []
    for i, s in enumerate(starts):
        chunk = text[s:starts[i + 1] if i + 1 < len(starts) else len(text)]
        period = PERIOD.search(chunk)
        client = _client(chunk)
        name, addr = _site_lines(chunk)
        sections.append({"client": client, "name": name, "address": addr,
                         "start": _date(period.group(1)) if period else None,
                         "end": _date(period.group(2)) if period else None,
                         "events": _events(chunk)})
    merged = []                                              # the same site continued on later pages = one section
    for s in sections:
        if merged and s["client"] and s["client"] == merged[-1]["client"] and s["start"] == merged[-1]["start"]:
            merged[-1]["events"] += s["events"]
        else:
            merged.append(s)
    return merged


def parse(path):
    return parse_text(read_text(path))
