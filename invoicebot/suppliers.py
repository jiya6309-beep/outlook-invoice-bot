"""Who issued the document?  Four steps, stopping at the first that works:

  1. an alias from the Suppliers sheet appears in the first lines of page 1   (most reliable)
  2. the printed issuer line (a line ending in Pty Ltd / Limited, or the first plain name line)
  3. a hint from the sender (Senders sheet)
  4. the logo, read by OCR and matched to names we already know
"""
import difflib
import re

from . import companies, settings

COMPANY_SUFFIX = re.compile(
    r"(pty\.?\s*ltd\.?|pty\.?\s*limited|limited|ltd\.?|p/l|incorporated|inc\.?(?!\s*(?:gst|tax)))(?![a-z])", re.IGNORECASE)
NOT_A_NAME = re.compile(
    r"^(tax\s+invoice|invoice|statement|credit\s+note|credit\s+memo|adjustment\s+note|"
    r"customer\s+copy|original|duplicate|reprint|remittance.*|delivery\s+docket|"
    r"proof\s+of\s+delivery|page\s*\d.*|account\s+summary|\W*continued.*)$", re.IGNORECASE)
CUSTOMER_LINE = re.compile(
    r"\b(bill(ed)?|invoice|sold|ship(ped)?|deliver(y)?|charge)\s+to\b|\bpearl\b|\battn\b", re.IGNORECASE)
ADDRESS_WORDS = re.compile(
    r"\b(street|road|highway|hwy|level|po\s*box|suite|unit|parade|avenue|drive)\b|@|www\.|\.com", re.IGNORECASE)
LOOKS_LIKE_A_DOCUMENT = re.compile(
    r"invoice|statement|credit|abn|a\.b\.n|remittance|delivery|docket|receipt|payment\s+advice|account|report",
    re.IGNORECASE)
JUNK = re.compile(r"\b(statement|report|summary|remittance|reminder|notice|contents|page|date|invoice|tax|"
                  r"credit note|follow|catalogue|customer|total|gst|amount|balance|qty|description|price|deliver|"
                  r"subtotal|payable)\b", re.IGNORECASE)
TITLE_PREFIX = re.compile(
    r"^(?:(?:adjustment\s+note|adjustment|tax\s+invoice|credit\s+note|note\s*-\s*credit|credit\s+memo|"
    r"invoice|statement|original|duplicate|copy)\b[\s\-:]*)+", re.IGNORECASE)
STOP_NAMES = {"australia", "nsw", "qld", "vic", "tas", "act", "sa", "wa", "nt", "new south wales", "queensland",
              "victoria", "tasmania", "customer", "supplier", "remit", "page", "total", "balance", "amount",
              "abn", "tax invoice", "invoice", "tax", "payment due", "due date", "terms", "discount", "date", "invoice date", "order date", "payment terms"}
GENERIC_WORDS = {"pty", "ltd", "limited", "inc", "the", "and", "of", "australia", "australian", "group", "co", "company"}


def _title_case(name):
    out = name.title()
    return re.sub(r"\bMc([a-z])", lambda m: "Mc" + m.group(1).upper(), out)           # Mcallister -> McAllister


def tidy_name(name):
    name = TITLE_PREFIX.sub("", re.sub(r"\s+", " ", name).strip())
    name = name.strip(" -,:;*|.")
    name = _title_case(name) if name.isupper() else name
    endings = {e for e in settings.legal_endings() if e.isalpha()}
    return " ".join(w.capitalize() if w.lower().strip(".") in endings and w.isupper() else w for w in name.split(" "))


ACCOUNT_NAME = re.compile(r"account\s*name\s*[:\-]\s*([^\n]{3,90})", re.IGNORECASE)


def account_name(text):
    """The payee in the bank-deposit box ('Account Name: Envisage Holdings PTY LTD') - the company that issued the invoice."""
    for m in ACCOUNT_NAME.finditer(text):
        raw = re.split(r"\s+[A-Za-z][A-Za-z ]{1,20}\s*:", m.group(1))[0]          # stop at the next label ('Terms:')
        end = ending_end(raw)
        name = tidy_name(raw[:end] if end else raw)
        if name and not re.search(r"pearl", name, re.IGNORECASE) and plausible(name) and 3 <= len(name) <= 70:
            return name
    return None


def ending_end(line):
    """Where the legal ending of a company name ends in the line (position after it), or None.
    The endings come from the LegalEndings sheet (Pty, Pvt, Private, Limited, Ltd ...); a run such as 'Pty Ltd' / 'Pvt. Limited' is taken
    together, and a line whose ending is cut off ('SONS PTY.' with 'LIMITED' on the next line) ends at the ending it has.
    A bare ending with no name before it is not a name (returns None)."""
    endings = {e.strip("./") for e in settings.legal_endings()} | set(settings.legal_endings())
    toks = list(re.finditer(r"\S+", line))
    for i, t in enumerate(toks):
        word = t.group().strip(".,;:()").lower()
        if i == 0 or word not in endings:
            continue
        nxt = toks[i + 1].group().strip(".,;:").lower() if i + 1 < len(toks) else ""
        if word == "inc" and nxt in ("gst", "tax", "of", "this"):
            continue                                                  # 'Inc GST' is an amount label, not a company ending
        j = i
        while j + 1 < len(toks) and toks[j + 1].group().strip(".,;:()").lower() in endings:
            j += 1
        return toks[j].end()
    return None


def is_junk(name):
    return bool(JUNK.search(name or ""))


def looks_like_a_document(text):
    return bool(LOOKS_LIKE_A_DOCUMENT.search(text[:3000]))


LABEL_PREFIX = re.compile(r"^\s*(?:(?:bill(?:ed)?|invoice|sold|ship(?:ped)?|deliver(?:y)?|charge)\s+to|customer|"
                          r"(?:invoice|order|inv|ref(?:erence)?)\s*(?:#|no\.?|number|num))(?![A-Za-z0-9])[:#\s]*", re.IGNORECASE)


def _after_labels(line):
    """'Bill to Ship to Dorries Anglers Choice Pty Ltd' -> 'Dorries Anglers Choice Pty Ltd'. Two columns side by side (Bill to | Ship to | the
    supplier's address block on the right) are read as ONE line; the labels at its start belong to the customer columns, the end to the supplier."""
    rest, changed = line, False
    while True:
        m = LABEL_PREFIX.match(rest)
        if not m:
            break
        rest, changed = rest[m.end():], True
    return rest.strip() if changed else None


def heuristic_names(text):
    """(strong, weak) issuer names printed at the top of page 1.
    strong = a line ending in Pty Ltd / Limited ...;  weak = the first plain name-like line."""
    if not looks_like_a_document(text):
        return None, None
    lines = [l.strip().lstrip("*").strip() for l in text.splitlines() if l.strip()][:14]
    weak, address_block = None, 0
    for line in lines:
        tail = _after_labels(line)
        if tail and ending_end(tail) and not CUSTOMER_LINE.search(tail) and not companies.find_in_text(tail, header_lines=1):
            name = tidy_name(tail[:ending_end(tail)])
            if 3 <= len(name) <= 70 and re.search(r"[A-Za-z]{3}", name) and not is_junk(name):
                return name, weak                      # the supplier's own name, printed beside the Bill to / Ship to labels
        customer = CUSTOMER_LINE.search(line) or companies.find_in_text(line, header_lines=1)
        if customer:
            if weak:
                break                                  # the "Invoice To" section started: stop reading
            address_block = 4                          # its address lines are not the supplier
            continue
        if NOT_A_NAME.match(line) or re.search(r"account\s*name|\bbsb\b|\bpayee\b|beneficiary", line, re.IGNORECASE):
            continue                                   # bank-box lines are read separately (account_name)
        suffix_end = ending_end(line)
        if suffix_end:
            name = tidy_name(line[:suffix_end])
            if 3 <= len(name) <= 70 and re.search(r"[A-Za-z]{3}", name) and not is_junk(name):
                return name, weak
        elif weak is None and address_block <= 0 and not re.search(r"\d", line) and not ADDRESS_WORDS.search(line):
            cut = re.split(r"\s+(?:vendor|customer|phone|ph|tel|fax|abn|a\.b\.n|acn|email)\b|:",
                           line, flags=re.IGNORECASE)[0]
            name = tidy_name(cut)
            if (3 <= len(name) <= 60 and re.search(r"[A-Za-z]{3}", name) and not is_junk(name)
                    and name.lower() not in STOP_NAMES):
                weak = name
        address_block -= 1
    deeper = footer_name(text, skip=len(lines))
    if deeper:
        return deeper, weak
    return None, weak


PAYEE_PREFIX = re.compile(r"^\s*(?:remit(?:tance)?(?:\s+(?:to|advice))?|pay(?:ment)?\s+to|make\s+(?:cheques?\s+|payments?\s+)?payable\s+to|issued\s+by|trading\s+as)\s*[:\-]?\s*", re.IGNORECASE)
ABN_WORD = re.compile(r"\bA\.?B\.?N\b", re.IGNORECASE)


def footer_name(text, skip=14):
    """The supplier's name when it is printed LOWER on the page than the top lines: a line ending in Pty Ltd / Limited ... that sits next to the
    supplier's ABN ("XYZ Foods Pty Ltd  ABN 12 345 678 901") or after "Remit to / Pay to / Issued by / Trading as". Lines of the customer (Bill to,
    Deliver to, your own companies) and bank-box lines are never used. A Pty Ltd name with neither clue is NOT taken."""
    lines = [l.strip().lstrip("*").strip() for l in text.splitlines() if l.strip()]
    for i in range(skip, len(lines)):
        line = lines[i]
        if CUSTOMER_LINE.search(line) or companies.find_in_text(line, header_lines=1):
            continue
        if re.search(r"account\s*name|\bbsb\b|\bpayee\b|beneficiary", line, re.IGNORECASE):
            continue
        m = PAYEE_PREFIX.match(line)
        candidate = line[m.end():] if m else line
        end = ending_end(candidate)
        if not end:
            continue
        near_abn = bool(ABN_WORD.search(line) or (i + 1 < len(lines) and ABN_WORD.search(lines[i + 1])) or (i > 0 and ABN_WORD.search(lines[i - 1])))
        if not (m or near_abn):
            continue
        name = tidy_name(candidate[:end])
        if 3 <= len(name) <= 70 and re.search(r"[A-Za-z]{3}", name) and not is_junk(name):
            return name
    return None


def heuristic_name(text):
    strong, weak = heuristic_names(text)
    return strong or weak


def known_names(extra=()):
    """Every supplier name we can recognise a logo against."""
    names = [r["Record name"] for r in settings.rows("Suppliers") if r["Record name"]]
    names += [r["Supplier name"] for r in settings.rows("Senders") if r.get("Supplier name")]
    names += list(extra)
    seen, out = set(), []
    for n in names:
        if n and n.lower() not in seen:
            seen.add(n.lower())
            out.append(n)
    return out


def _tokens(text):
    return [t for t in settings.words(text).split() if len(t) >= 3 and t not in GENERIC_WORDS]


def match_logo_text(ocr_text, names):
    """Match noisy OCR text to a known supplier name. Returns (name, score) or (None, best score).
    Every distinctive word of the known name must be found (allowing small OCR typos)."""
    ocr_tokens = _tokens(ocr_text)
    if not ocr_tokens:
        return None, 0.0
    scored = []
    for name in names:
        want = _tokens(name)
        if not want:
            continue
        ok = 0
        for w in want:
            best = max((difflib.SequenceMatcher(None, w, t).ratio() for t in ocr_tokens), default=0)
            if best >= (0.80 if len(w) >= 7 else 0.88):
                ok += 1
        scored.append((ok / len(want), len(want), name))
    scored.sort(reverse=True)
    if scored and scored[0][0] >= 0.999 and (len(scored) == 1 or scored[1][0] < 0.999 or scored[1][1] < scored[0][1]):
        return scored[0][2], 1.0
    return None, scored[0][0] if scored else 0.0


GENERIC_SENDER_WORDS = {"accounts", "account", "accounts payable", "accountspayable", "noreply", "no reply", "donotreply", "do not reply",
                        "admin", "info", "invoices", "invoice", "invoicing", "sales", "orders", "order", "office", "support", "enquiries",
                        "mail", "team", "finance", "billing", "ar", "ap", "customer service", "reception", "service", "services",
                        "payments", "remittance", "statements", "statement", "debtors", "creditors", "postmaster", "system"}


def sender_name_guess(sender_name):
    """A last-resort supplier name: the sender's display name, when it looks like a business name (not a mailbox role, not Pearl)."""
    n = re.sub(r"\s+", " ", re.sub(r"\s+via\s+.*$", "", str(sender_name or ""), flags=re.IGNORECASE)).strip(" \"'<>")
    if not n or "@" in n or re.search(r"pearl", n, re.IGNORECASE) or settings.words(n) in GENERIC_SENDER_WORDS:
        return None
    n = tidy_name(n)
    return n if plausible(n) and re.search(r"[A-Za-z]{3}", n) else None


FREE_MAIL = {"gmail", "outlook", "hotmail", "yahoo", "live", "icloud", "bigpond", "optusnet", "msn", "me", "aol", "proton", "protonmail"}


def name_from_email_domain(text):
    """The supplier's own e-mail address is printed on the document (hello@goldcoastcookies.com.au): the words of the name that
    squash to the domain ('Gold Coast Cookies' = goldcoastcookies) are the supplier. Pearl's own addresses and free mail are skipped."""
    seen = set()
    for dom in re.findall(r"[\w.+-]+@([\w-]+(?:\.[\w-]+)+)", text or ""):
        label = dom.lower().split(".")[0]
        if label in seen or label in FREE_MAIL or "pearl" in label or len(label) < 5:
            continue
        seen.add(label)
        for line in text.splitlines():
            words = [w for w in re.split(r"\s+", line) if w]
            for i in range(len(words)):
                squashed = ""
                for j in range(i, min(i + 6, len(words))):
                    squashed += re.sub(r"[^a-z0-9]", "", words[j].lower())
                    if squashed == label and j > i:
                        name = tidy_name(" ".join(words[i:j + 1]))
                        if plausible(name):
                            return name
                    if len(squashed) >= len(label):
                        break
    return None


# the issuer's own block ends where the customer / the item table starts: a brand in the item list must never decide the supplier
BLOCK_END = re.compile(r"\b(?:sold|deliver(?:ed|y)?|ship(?:ped)?|bill(?:ed)?|invoice|charge)\s+to\b|\bcustomer\s+(?:no|number|name)\b|"
                       r"\b(?:qty|quantity|description|stock\s+code|product|item\s+code|brand|unit\s+price)\b", re.IGNORECASE)
TRADING_AS = re.compile(r"\b(?:t/a|t/as|trading\s+as)\b[\s:]*(.*)$", re.IGNORECASE)


def issuer_block(doc_text, limit=25):
    """The lines at the top of page 1 that belong to the supplier (before 'Sold To', 'Deliver To' or the item table)."""
    out = []
    for line in [l for l in doc_text.splitlines() if l.strip()][:limit]:
        if out and BLOCK_END.search(line):
            break
        out.append(line)
    return "\n".join(out)


def trading_name(block):
    """'Setlex Pty Ltd T/A' + next line 'Rum City Foods'  (or 'Setlex Pty Ltd T/A Rum City Foods') -> 'Rum City Foods'."""
    lines = [l.strip() for l in block.splitlines() if l.strip()]
    for i, line in enumerate(lines):
        m = TRADING_AS.search(line)
        if not m:
            continue
        rest = m.group(1).strip(" -:,")
        if not rest and i + 1 < len(lines):
            rest = lines[i + 1].strip()
        rest = re.split(r"\s{2,}|\bABN\b|\bA\.B\.N\b", rest)[0].strip(" -:,")
        name = tidy_name(rest)
        if name and plausible(name) and not companies.find_in_text(name, header_lines=1) and "pearl" not in name.lower():
            return name
    return None


def from_sender_domain(sender_email, doc_text):
    """An e-mail sent straight from the supplier (not a forward): its domain is a strong clue. A domain in the supplier list gives
    that supplier; otherwise the words on the page that squash to the domain ('Rum City Foods' = rumcityfoods)."""
    dom = (sender_email or "").lower().rsplit("@", 1)[-1] if "@" in (sender_email or "") else ""
    label = dom.split(".")[0] if dom else ""
    if not dom or label in FREE_MAIL or "pearl" in label or "anantaa" in label:
        return None
    row = settings.supplier_id_row_for("x@" + dom)
    if row:
        return row["Supplier (record name)"]
    return name_from_email_domain(f"{doc_text}\nfrom: x@{dom}")


def identify(doc_text, sender=None, logo_texts=None, extra_known=(), sender_name="", sender_email="", is_forward=False):
    """Returns (printed name or None, source) where source is table | printed | sender | logo | ''.
    `logo_texts` is a function returning OCR text of the logos (called only if needed).
    Order: the supplier list (ABN / e-mail domain / account name) -> a printed 'T/A' trading name -> the sender's own domain
    (not for forwards) -> the Suppliers sheet's 'Name contains', ONLY in the supplier's header block -> the printed name ..."""
    header = issuer_block(doc_text)
    ident = settings.supplier_id_row_for(doc_text)               # ABN, e-mail domain or "Account Name": ANY ONE is enough (the supplier list wins)
    if ident:
        return ident["Supplier (record name)"], "table"
    traded = trading_name(header)
    if traded:
        row = settings.supplier_row_for(traded)
        return (row["Record name"] if row else traded), "printed (trading as)"
    if not is_forward:
        by_domain = from_sender_domain(sender_email, doc_text)
        if by_domain:
            return by_domain, "sender e-mail domain"
    row = settings.supplier_row_for(header)
    if row:
        return row["Record name"], "table"
    strong, weak = heuristic_names(doc_text)
    if strong:
        return strong, "printed"
    paid_to = account_name(doc_text)
    if paid_to:
        return paid_to, "printed"
    if sender and sender.get("Supplier name"):
        return sender["Supplier name"], "sender"
    named = name_from_email_domain(doc_text)
    if named:
        return named, "printed"
    if weak:
        return weak, "printed"
    if logo_texts is not None and looks_like_a_document(doc_text):
        text = logo_texts()
        if text:
            row = settings.supplier_row_for(text)
            if row:
                return row["Record name"], "logo"
            name, _score = match_logo_text(text, known_names(extra_known))
            if name:
                return name, "logo"
    guess = sender_name_guess(sender_name)
    if guess:
        return guess, "sender name"
    return None, ""


FOOTER_PHRASES = ("view online", "view and pay online", "pay online", "legal title", "remains in", "terms and conditions", "please ", "thank you",
              "remittance", "click here", "your invoice", "tax invoice", "page ")
BARE_LEGAL_WORDS = {"limited", "ltd", "pty", "pty ltd", "pty limited", "pvt", "pvt ltd", "pvt limited", "private limited", "inc", "holdings",
                    "group", "company"}


GREETING_OR_LABEL = re.compile(
    r"^(?:dear|hi|hello|greetings|good\s+(?:morning|afternoon|day)|to\s+whom|attn|attention|"
    r"billing\s+(?:to|address|details)|bill(?:ed)?\s+to|invoice\s+to|ship(?:ping|ped)?\s+to|deliver(?:y|ed)?\s+to|sold\s+to|"
    r"payment\s+(?:to|details)|remit(?:tance)?\s+to|account\s+(?:name|number|details)|customer\s+(?:name|number|no))\b",
    re.IGNORECASE)


def plausible(name):
    """Does a printed supplier name look like a name? Rejects footer sentences ('Legal title of these goods remains in ...'),
    web-page links ('View and pay online') and a bare legal ending ('Limited')."""
    n = (name or "").strip()
    low = n.lower()
    if not n or len(n) > 70 or len(n.split()) > 9 or low in BARE_LEGAL_WORDS:
        return False
    if any(p in low for p in FOOTER_PHRASES):
        return False
    if GREETING_OR_LABEL.match(low):                     # 'Dear Valued Customer', 'Billing To ...' are never a supplier
        return False
    letters = sum(c.isalpha() for c in n)
    return letters >= 3 and letters >= len(n) * 0.5
