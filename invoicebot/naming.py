"""File names:  Supplier_InvoiceNumber.pdf  (the pattern is editable per document type)."""
import os
import re
import unicodedata

MAX_STEM = 110


def tidy(text):
    """Odd characters that PDF text brings along: dash-like characters become '-', Symbol-font characters (private-use codes such as
    U+F02D) become their ASCII twins, other spaces become a plain space and invisible characters (soft hyphen, zero-width ...) disappear.
    Without this a number like INV-3121 can reach the file name with a character Windows draws as a box."""
    if text is None:
        return None
    out = []
    for ch in str(text):
        code = ord(ch)
        if 0xF020 <= code <= 0xF07E:
            ch = chr(code - 0xF000)
        cat = unicodedata.category(ch)
        if cat == "Pd" or ch == "\u2212":
            out.append("-")
        elif cat == "Zs":
            out.append(" ")
        elif cat in ("Cf", "Cc", "Co", "Cs", "Cn"):
            out.append(" " if ch in "\t\n\r" else "")
        else:
            out.append(ch)
    return "".join(out)


def clean(text, keep_spaces=True):
    """Remove characters Windows / SharePoint do not allow, tidy spaces."""
    t = re.sub(r'[<>:"/\\|?*\x00-\x1f#%{}~]', "-" if keep_spaces else "", tidy(text))
    t = re.sub(r"\s+", " ", t).strip(" .-")
    return t


DEFAULT_DATE_FORMAT = "dd-mm-yyyy"


def format_date(d, fmt=DEFAULT_DATE_FORMAT):
    """The date as it appears in a file name.  dd-mm-yyyy -> 04-10-2026   (tokens: dd mm yyyy yy; anything else is kept)."""
    text = str(fmt or DEFAULT_DATE_FORMAT).lower()
    for token, code in (("yyyy", "%Y"), ("yy", "%y"), ("mm", "%m"), ("dd", "%d")):
        text = text.replace(token, code)
    return d.strftime(text)


def clean_folder(name):
    """A folder name for a path: only the characters Windows forbids are replaced; '&', '.', digits ... are kept
    so the name matches the real SharePoint folder character for character (for example 'P & O INVESTMENT', 'AMPOL  KELSO')."""
    t = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "-", str(name))
    return t.strip(" .")                    # spaces INSIDE the name are kept: the real folder may be 'AMPOL  KELSO' (two spaces)


def build(pattern, supplier="", short="", number="", date="", tag="", doc_type="", original="", ext=".pdf", site=""):
    """Fill the pattern. Missing parts are replaced by sensible fallbacks, never left empty."""
    stem_original = clean(os.path.splitext(original or "")[0], True) or "document"
    values = {
        "supplier": clean(supplier) or "Unknown supplier",
        "short": clean(short) or clean(supplier) or "Unknown",
        "number": clean(number) or stem_original[:40],
        "date": str(date or ""),
        "tag": clean(tag), "type": clean(doc_type), "site": clean(site) or "NoSite",
    }
    name = pattern
    for k, v in values.items():
        name = name.replace("{" + k + "}", v)
    name = re.sub(r"[{}]", "", name)
    name = re.sub(r"_{2,}", "_", name).strip("_ ")
    name = clean(name).replace(" _", "_").replace("_ ", "_")
    if len(name) > MAX_STEM:
        name = name[:MAX_STEM].rstrip(" ._-")
    ext = (ext or "").lower()
    return (name or "document") + ext


def unique_name(name, exists):
    """name, name_2, name_3 ... until `exists(candidate)` is False."""
    base, ext = os.path.splitext(name)
    candidate, n = name, 2
    while exists(candidate):
        candidate = f"{base}_{n}{ext}"
        n += 1
        if n > 200:
            break
    return candidate
