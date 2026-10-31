"""Small, well-tested helpers for reading dates and money. No PDF or Excel code in here."""
import re
import datetime

DATE_TOKEN = (r"(?:\d{4}-\d{2}-\d{2}"
              r"|\d{1,2}[./\-]\d{1,2}[./\-]\d{2,4}"
              r"|\d{1,2}(?:st|nd|rd|th)?[ \-][A-Za-z]{3,9}[ ,\-]{1,3}\d{2,4}"
              r"|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]{0,6}\.? \d{1,2}(?:st|nd|rd|th)?,? \d{4})")   # Jul 23, 2026 (Fresho)

# 1,234.56  $1,234.56  A$3,364.44  (110.00)  -110.00  110.00-
MONEY_TOKEN = r"\(?-?(?:AUD|A?\$)?\s*\d[\d,]*\.\d{2}(?!\d)\)?-?"

_DATE_FORMATS = ("%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y", "%d.%m.%Y", "%d.%m.%y",
                 "%d-%m-%Y", "%d-%m-%y", "%d %b %Y", "%d %b %y", "%d %B %Y",
                 "%d %B %y", "%d-%b-%Y", "%d-%b-%y", "%b %d %Y", "%B %d %Y")


def parse_date(text):
    """Return a date, or None. Day comes before month (Australian style)."""
    if not text:
        return None
    t = re.sub(r"[\u200b-\u200d\ufeff]", "", str(text))
    t = re.sub(r"\bsept(?:ember)?\b\.?", "Sep", t, flags=re.IGNORECASE)             # 26 Sept 2026
    t = re.sub(r"(?<=\d)(st|nd|rd|th)\b", "", t, flags=re.IGNORECASE)
    t = re.sub(r"\s+", " ", t.replace(",", " ")).strip()
    for fmt in _DATE_FORMATS:
        try:
            d = datetime.datetime.strptime(t, fmt).date()
        except ValueError:
            continue
        if 2000 <= d.year <= 2100:
            return d
    return None


def parse_money(text):
    """Return a float (negative for -x, (x) or x-), or None."""
    if text is None:
        return None
    s = str(text).strip()
    m = re.search(r"\d[\d,]*\.\d{2}", s)
    if not m:
        return None
    value = float(m.group().replace(",", ""))
    negative = s.startswith("-") or s.endswith("-") or ("(" in s and ")" in s)
    return -value if negative else value


def find(pattern, text):
    """First capture group of a pattern (ignoring capitals), or None."""
    match = re.search(pattern, text, re.IGNORECASE)
    return match.group(1).strip() if match else None


def to_number(text):
    return float(text.replace(",", "")) if text else None
