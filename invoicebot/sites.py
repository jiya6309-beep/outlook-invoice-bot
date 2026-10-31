"""Sites: matching a document to a site (address first, then names/aliases), and importing the site lists.

The site is information for the record and the handler e-mails. Where a file is SAVED is decided by
the company folder (see router.py), so an unknown site never blocks a document."""
import collections
import difflib
import os
import re

from . import companies, settings

STREET_TYPES = {"street", "st", "road", "rd", "avenue", "ave", "av", "highway", "hwy", "drive", "dr", "lane", "ln",
                "court", "ct", "place", "pl", "circuit", "cct", "crescent", "cres", "boulevard", "blvd", "parade",
                "pde", "way", "terrace", "tce", "close", "cl", "mwy", "motorway", "freeway", "fwy", "esplanade",
                "esp", "square", "sq", "track", "trail", "row", "loop", "grove", "gr", "rise", "ridge"}
STATES = {"nsw", "qld", "vic", "sa", "wa", "tas", "nt", "act"}
TYPO_FIXES = {"barleigh": "burleigh", "anangroove": "annangrove", "hasting": "hastings",
              "beresfeild": "beresfield", "farliegh": "farleigh", "leongath": "leongatha", "wadonga": "wodonga"}


def make_fragments(address):
    """Short phrases that identify an address on an invoice:  '19 balook', 'balook drive beresfield'."""
    toks = settings.words(address).split()
    i = 0
    while i < len(toks) and toks[i] in ("lot", "unit", "shop", "suite", "level"):
        i += 1
    nums = []
    while i < len(toks) and re.fullmatch(r"\d+[a-z]?", toks[i]) and len(nums) < 3:
        nums.append(toks[i])
        i += 1
    while i < len(toks) and toks[i] in ("cnr", "corner", "of", "and"):
        i += 1
    street, stype = [], ""
    while i < len(toks) and len(street) < 4:
        if toks[i] in STREET_TYPES:
            stype = toks[i]
            i += 1
            break
        street.append(toks[i])
        i += 1
    sub = []
    while i < len(toks) and toks[i] not in STATES and not re.fullmatch(r"\d{4}", toks[i]) and len(sub) < 4:
        sub.append(toks[i])
        i += 1
    frags = []
    if street and nums:
        for n in nums:
            frags.append(f"{n} {street[0]}")
        if len(nums) > 1:
            frags.append(" ".join(nums) + f" {street[0]}")
    if street and sub:
        frags.append(" ".join(street + ([stype] if stype else []) + sub))
    elif street and not nums:
        frags.append(" ".join(street + ([stype] if stype else [])))
    seen, out = set(), []
    for f in frags:
        if f and f not in seen:
            seen.add(f)
            out.append(f)
    return out


def _alts(cell):
    return [a.strip() for a in str(cell).split("|") if a.strip()]


GROUP_NAMES = ("pearl energy", "pearl")          # invoices are often addressed to "PEARL ENERGY <TOWN>"


def _place(name_words, brands):
    """The town part of a site name: 'mobil colyton' -> 'colyton'; 'pearl energy greta' -> 'greta'."""
    for group in GROUP_NAMES:
        if name_words.startswith(group + " "):
            return name_words[len(group) + 1:]
    first, _, rest = name_words.partition(" ")
    return rest if first in brands else ""


def company_of(site):
    """The 'Company name' (Companies sheet) that owns a site, or ''."""
    return (companies.by_name(site.get("company", "")) or {}).get("Company name", "")


def active_sites(only_company=None):
    """The active sites with every name they can be called. Besides the site name and the aliases in the sheet, a site whose
    town is unique in the list also answers to 'Pearl Energy <town>' (how invoices and dockets usually address it)."""
    out = []
    for r in settings.rows("Sites"):
        if not settings.yes(r.get("Active", ""), default=True):
            continue
        names = [r["Site name"]] + _alts(r.get("Aliases", ""))
        out.append({"name": r["Site name"], "company": r.get("Company", ""),
                    "names": [settings.words(n) for n in names if n] + settings.site_alias_names(r["Site name"]),
                    "frags": [settings.words(f) for f in _alts(r.get("Address fragments", ""))],
                    "handler": r.get("Handler", ""), "email": r.get("Handler email", ""),
                    "folder": r.get("SharePoint folder", "").strip()})
    firsts = collections.Counter(s["names"][0].split(" ")[0] for s in out if s["names"])
    brands = {w for w, n in firsts.items() if n >= 3 and w not in ("pearl",)}          # brands are recognised in the WHOLE list ...
    if only_company:
        out = [s for s in out if company_of(s) == only_company]                            # ... but a town only has to be unique within the company
    places = [(_place(s["names"][0], brands) if s["names"] else "") for s in out]
    counts = collections.Counter(p for p in places if p)
    for s, place in zip(out, places):
        s["place"], s["generic"] = place, ("pearl energy " + place if place else "")
        if place and counts[place] == 1:
            alias = "pearl energy " + place
            if alias not in s["names"]:
                s["names"].append(alias)
    return out


def site_row(name):
    for s in active_sites():
        if s["name"] == name:
            return s
    return None


def derive_folder(site_name):
    """The folder name a site has when no real one is recorded: its name in capitals, extra spaces removed."""
    return re.sub(r"\s+", " ", str(site_name)).strip().upper()


def folder_for(site_name):
    """The SharePoint folder of a site: the 'SharePoint folder' cell if filled in, otherwise derived from the name."""
    row = site_row(site_name)
    return (row["folder"] if row and row["folder"] else derive_folder(site_name))


def find_site(text):
    """The site whose name, alias or SharePoint folder equals the text (capitals/spaces ignored), or None."""
    want = settings.words(text)
    if not want:
        return None
    for s in active_sites():
        if want in s["names"] or (s["folder"] and settings.words(s["folder"]) == want):
            return s
    return None


def folder_names():
    """Every site's folder name, sorted (for the drop-down lists)."""
    return sorted({folder_for(s["name"]) for s in active_sites()})


def _name_hits(sites, text_words):
    """Sites whose name/alias appears; a match inside a longer match of ANOTHER site is dropped
    ('Ampol Beresfield' inside 'Ampol Beresfield North')."""
    matches = []
    for s in sites:
        for n in s["names"]:
            if n and settings.has_phrase(text_words, n):
                matches.append((s["name"], n))
    keep = set()
    for name, n in matches:
        covered = any(other != name and len(on) > len(n) and n in on for other, on in matches)
        if not covered:
            keep.add(name)
    return keep


_RANGE = re.compile(r"(?<![a-z0-9])(\d{1,5}) (\d{1,5})(?= [a-z])")


def _range_variants(text_words):
    """'102 104 wollombi road' (printed 102-104) also reads as '102 wollombi road' and '104 wollombi road' (5.18.0)."""
    out = [text_words]
    if _RANGE.search(text_words):
        def first(m):
            return m.group(1) if int(m.group(1)) < int(m.group(2)) else m.group(0)

        def second(m):
            return m.group(2) if int(m.group(1)) < int(m.group(2)) else m.group(0)
        out += [_RANGE.sub(first, text_words), _RANGE.sub(second, text_words)]
    return out


def _frag_hits(sites, text_words):
    texts = _range_variants(text_words)
    return {s["name"] for s in sites if any(f and settings.has_phrase(t, f) for f in s["frags"] for t in texts)}


def _fuzzy(sites, blocks_words):
    """Last resort: a site name that is almost (not exactly) in the text - supplier or list typos."""
    best, second = (0.0, None), 0.0
    tokens = blocks_words.split()
    for s in sites:
        for n in s["names"]:
            size = len(n.split())
            if size < 2 or size > len(tokens):
                continue
            for k in range(0, len(tokens) - size + 1):
                window = " ".join(tokens[k:k + size])
                ratio = difflib.SequenceMatcher(None, n, window).ratio()
                if ratio > best[0] and s["name"] != best[1]:
                    second, best = best[0], (ratio, s["name"])
                elif ratio > best[0]:
                    best = (ratio, s["name"])
    if best[0] >= 0.93 and best[0] - second >= 0.03:
        return best[1]
    return None


SHIP_TAG = "[SHIP] "          # put in front of a block that was read under a ship-to / deliver-to label (see labelreader._read_blocks)
SHIP_SOURCE = "ship-to / deliver-to address"


SHIP_HOLD = "SHIP-TO NOT IN SITE LIST"      # the start of site_source when a ship-to address is on the page but is not a known site
SHIP_CONFLICT = "SHIP-TO UNCLEAR"           # the ship-to block itself names two sites, or a company that does not own its site
SOLD_HOLD = "SOLD-TO NOT IN SITE LIST"      # 5.15.0: no ship-to, and the sold-to address is not a known site: never the bill-to


def ship_hold(source):
    """True when a site_source says the document waits because of its ship-to address."""
    return str(source or "").startswith((SHIP_HOLD, SHIP_CONFLICT, SOLD_HOLD))


def looks_like_address(text):
    """A ship-to block is only trusted as an ADDRESS when it has a postcode, a state or a street word, so a stray
    'Deliver to' in a sentence never makes a document wait."""
    toks = settings.words(text).split()
    if any(re.fullmatch(r"\d{4}", t) or t in STATES for t in toks):
        return True
    return any(re.fullmatch(r"\d+[a-z]?", t) and any(w in STREET_TYPES for w in toks[i + 2:i + 6])
               for i, t in enumerate(toks))                          # '1B BELL STREET': a number, then a street word


def _trusted_ship_block(text, sites):
    """A ship-to block decides only when it is really about a delivery place: an address, a known site, a known company,
    or a site-style name ('PEARL ENERGY PRESTON', 'AMPOL ...', 'BP ...') - not a stray 'Deliver to' in a sentence."""
    if looks_like_address(text) or companies.find_in_text(text.replace(" | ", "\n")):
        return True
    tw = settings.words(text)
    starts = {n.split()[0] for s in sites for n in s["names"][:1] if n}
    return any(settings.has_phrase(tw, g) for g in GROUP_NAMES) or any(settings.has_phrase(tw, b) for b in starts) \
        or bool(_ship_sites(sites, text, ""))


def _order_free_hits(sites, text_words):
    """Sites whose name appears with its words in another order: 'BP NORTH ORANGE' is the site 'BP Orange North'."""
    toks = text_words.split()
    hits = set()
    for s in sites:
        for n in s["names"]:
            parts = n.split()
            if len(parts) < 2 or len(parts) > len(toks):
                continue
            want = sorted(parts)
            if any(sorted(toks[k:k + len(parts)]) == want for k in range(len(toks) - len(parts) + 1)):
                hits.add(s["name"])
                break
    return hits


def _edits(a, b):
    """The number of single-letter changes (insert, delete, replace, swap two neighbours) between two words."""
    d = [[i + j if i * j == 0 else 0 for j in range(len(b) + 1)] for i in range(len(a) + 1)]
    for i in range(1, len(a) + 1):
        for j in range(1, len(b) + 1):
            d[i][j] = min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + (a[i - 1] != b[j - 1]))
            if i > 1 and j > 1 and a[i - 1] == b[j - 2] and a[i - 2] == b[j - 1]:
                d[i][j] = min(d[i][j], d[i - 2][j - 2] + 1)
    return d[-1][-1]


def _close_word(a, b):
    """A small spelling difference in a street word: 1 change for words of 5-6 letters, 2 for longer ones; digits must be identical."""
    if a == b:
        return True
    if any(c.isdigit() for c in a + b) or min(len(a), len(b)) < 5:
        return False
    return _edits(a, b) <= (2 if min(len(a), len(b)) >= 7 else 1)


def _fuzzy_frag_hits(sites, text_words):
    """Address fragments with a small spelling difference ('5 HANRAHAN PLACE' on the page, '5 hanharan' in the list).
    The street NUMBER must be exactly the same; only the street word may differ slightly."""
    toks = text_words.split()
    hits = set()
    for s in sites:
        for f in s["frags"]:
            parts = f.split()
            if len(parts) < 2 or len(parts) > len(toks):
                continue
            for k in range(len(toks) - len(parts) + 1):
                window = toks[k:k + len(parts)]
                if parts[0].isdigit() and window[0] != parts[0]:
                    continue
                pairs = list(zip(parts, window))
                if all(_close_word(a, b) for a, b in pairs) and any(len(a) >= 5 and not a.isdigit() for a, _ in pairs):
                    hits.add(s["name"])
                    break
            if s["name"] in hits:
                break
    return hits


def _town_hit(sites, text):
    """5.17.0: the address block names the COMPANY (sites = only that company's sites) and a TOWN that exactly one of them is
    named after: 'BP Greta, ttf Pearl Energy No.3, 122 New England Hwy, GRETA' -> Pearl Energy Greta. Never used without the
    company: two companies can both have a site in one town (Pearl Energy Greta / Ampol Greta)."""
    sw = settings.words(text)
    hits = {s["name"] for s in sites if s.get("place") and settings.has_phrase(sw, s["place"])}
    return hits if len(hits) == 1 else set()


def _ship_sites(sites, ship_text, supplier):
    """The sites the SHIP-TO block points to, using the strongest clue that gives an answer:
    exact address / name -> account number printed inside the block -> name in another word order / small address typo -> similar name."""
    sw = settings.words(ship_text)
    hits = _frag_hits(sites, sw) | _name_hits(sites, sw)
    if hits:
        return hits
    sup = settings.words(supplier)
    names = {s["name"] for s in sites}
    for a in settings.rows("Accounts"):
        digits = a["Account number"].lstrip("0")
        if digits and sup and a["Site name"] in names and settings.has_phrase(sup, settings.words(a["Supplier name contains"])) \
                and re.search(r"(?<!\d)0*" + re.escape(digits) + r"(?!\d)", ship_text):
            return {a["Site name"]}
    hits = set()
    for line in re.split(r"\s*\|\s*|\n", ship_text):                       # word order and typos: only WITHIN one printed line
        lw = settings.words(line)
        if lw:
            hits |= _order_free_hits(sites, lw) | _fuzzy_frag_hits(sites, lw)
    if hits:
        return hits
    near = _fuzzy(sites, sw)
    return {near} if near else set()


SOLD_TAG = "[SOLD] "          # a block read under a 'Sold to' label (see labelreader._read_blocks)
SOLD_SOURCE = "sold-to address (no ship-to on the document)"


def _untag(block):
    b = str(block)
    for tag in (SHIP_TAG, SOLD_TAG):
        if b.startswith(tag):
            return b[len(tag):]
    return b


def sold_blocks(blocks):
    """The blocks read under a 'Sold to' label (tag removed)."""
    return [_untag(b) for b in blocks if str(b).startswith(SOLD_TAG)]


def split_blocks(blocks):
    """(ship-to blocks, all blocks), tags removed."""
    ship = [b[len(SHIP_TAG):] for b in blocks if str(b).startswith(SHIP_TAG)]
    return ship, [_untag(b) for b in blocks]


def _block_names_another_company(site_name, text):
    """True when the text names a known company that is not the company of this site (a contradictory address block)."""
    row = site_row(site_name)
    owner = companies.by_name(row["company"]) if row else None
    found = companies.find_in_text(text)
    if not found or not owner:
        return False
    return all(settings.norm(f[0]["Company name"]) != settings.norm(owner["Company name"]) for f in found)


def match(text, subject="", blocks=(), supplier="", to_name="", header_lines=40, only_company=None):
    """Returns {"site": name or UNASSIGNED, "source": how, "company": the site's company text}.
    only_company = a 'Company name': look among the sites of that company only (company first, then site)."""
    sold = sold_blocks(blocks)
    ship_blocks, blocks = split_blocks(blocks)
    sites = active_sites(only_company)
    result = {"site": settings.UNASSIGNED, "source": "not found", "company": ""}
    if not sites:
        return result

    def done(name, source):
        s = site_row(name)
        return {"site": name, "source": source, "company": s["company"] if s else ""}

    ship_on = settings.run_bool("Ship-to address wins over bill-to", True)
    ship_only = ship_on and settings.run_bool("Use only the ship-to address", True)
    ship_blocks = [b for b in ship_blocks if _trusted_ship_block(b, active_sites())] if ship_only else ship_blocks
    if ship_blocks and ship_only:                                          # 0. ONLY the ship-to / deliver-to address decides
        ship_text = " | ".join(ship_blocks)
        shown = re.sub(r"\s+", " ", ship_text.split(" | ")[0] if not re.fullmatch(r"\d+", ship_text.split(" | ")[0].strip())
                       else " | ".join(ship_text.split(" | ")[1:3])).strip()[:70]
        named = [] if only_company or not settings.run_bool("Choose the site within the company first", True) \
            else companies.find_in_text(ship_text.replace(" | ", "\n"))
        if len(named) == 1:                                                # the ship-to block names the COMPANY: look among its sites first
            row = named[0][0]
            mine = _ship_sites(active_sites(row["Company name"]), ship_text, supplier)
            if len(mine) == 1:
                return done(next(iter(mine)), f"{SHIP_SOURCE} (within {row['Folder name']})")
            if mine:
                result["source"] = f"{SHIP_CONFLICT}: the ship-to address fits more than one site of {row['Folder name']} ({' / '.join(sorted(mine))})"
                return result
            town = _town_hit(active_sites(row["Company name"]), ship_text)
            if town:
                return done(next(iter(town)), f"{SHIP_SOURCE} (company + town, within {row['Folder name']})")
            # the company is known FROM THE SHIP-TO, only its site is not: the company decides (company folder + SITE NOT FOUND, or a person picks)
            result.update(source=f"no site of {row['Folder name']} fits the ship-to address ('{shown}')",
                          company=row["Company name"], ship_company=True)
            return result
        hits = _ship_sites(sites, ship_text, supplier)
        if len(hits) == 1:
            only = next(iter(hits))
            if not _block_names_another_company(only, ship_text):
                return done(only, SHIP_SOURCE)
            result["source"] = f"{SHIP_CONFLICT}: the ship-to address points to {only}, but names another company ('{shown}')"
            return result
        if len(hits) > 1:
            result["source"] = f"{SHIP_CONFLICT}: the ship-to address fits more than one site ({' / '.join(sorted(hits))})"
            return result
        result["source"] = f"{SHIP_HOLD}: '{shown}'"
        return result

    if sold and not ship_blocks:                                          # 0b. NO ship-to / deliver-to: the SOLD-TO block comes before the bill-to
        sold_text = " | ".join(sold)
        sw = settings.words(sold_text)
        hits = _frag_hits(sites, sw) | _name_hits(sites, sw)
        if len(hits) > 1:                                                 # two sites inside the sold-to block: the one also named there
            named = _name_hits(sites, sw)
            hits = hits & named if len(hits & named) == 1 else hits
        if len(hits) == 1:
            only = next(iter(hits))
            if not _block_names_another_company(only, sold_text):
                return done(only, SOLD_SOURCE)
        if not hits and not only_company:
            named = companies.find_in_text(sold_text.replace(" | ", "\n"))
            if len(named) == 1:                                           # the sold-to names the company and a town of one of its sites
                town = _town_hit(active_sites(named[0][0]["Company name"]), sold_text)
                if town:
                    return done(next(iter(town)), f"{SOLD_SOURCE} - company + town")
        elif not hits and only_company:
            town = _town_hit(sites, sold_text)
            if town:
                return done(next(iter(town)), f"{SOLD_SOURCE} - company + town")

    sup = settings.words(supplier)
    for a in settings.rows("Accounts"):                                   # 1. supplier account number
        digits = a["Account number"].lstrip("0")
        if digits and sup and settings.has_phrase(sup, settings.words(a["Supplier name contains"])) \
                and (not only_company or a["Site name"] in {s["name"] for s in sites}):
            if re.search(r"(?<!\d)0*" + re.escape(digits) + r"(?!\d)", text):
                return done(a["Site name"], "account number")

    if sold and not ship_blocks and ship_on:                              # 1a. a sold-to that is not a known site: it waits, never the bill-to
        shown = re.sub(r"\s+", " ", " | ".join(x for x in sold[0].split(" | ") if not re.match(r"(?i)\s*(payer\s+)?account\b", x))[:70]).strip()
        result["source"] = f"{SOLD_HOLD}: '{shown}'"
        return result

    if ship_blocks and ship_on:                                           # 1b. the SHIP-TO / DELIVER-TO block decides, whatever the bill-to says
        sw = settings.words(" ".join(ship_blocks))
        ship_hits = _frag_hits(sites, sw) | _name_hits(sites, sw)
        if len(ship_hits) == 1:
            only = next(iter(ship_hits))
            if not _block_names_another_company(only, " ".join(ship_blocks)):
                return done(only, SHIP_SOURCE)
            # the ship-to block itself names a DIFFERENT company than that site belongs to: a contradiction, so the older, careful rules decide
        elif len(ship_hits) > 1:                                             # two sites inside the ship-to block itself: the bill-to must not add noise
            blocks = ship_blocks

    tw = settings.words("\n".join(str(text).splitlines()[:max(header_lines, 1) * 8]))
    bw = settings.words(" ".join(blocks))
    block_hits = _frag_hits(sites, bw) if bw else set()
    text_hits = _frag_hits(sites, tw)
    name_hits = _name_hits(sites, settings.words(subject + " " + " ".join(blocks)) + " " + tw)

    addr = block_hits or text_hits
    where = "address in deliver-to block" if block_hits else "address in text"
    if len(addr) == 1:
        only = next(iter(addr))
        if name_hits and only not in name_hits:
            result["source"] = f"conflict: address says {only}, name says {', '.join(sorted(name_hits))}"
            return result
        return done(only, where)
    if len(addr) > 1:
        # several addresses (billing + delivery): prefer the one that is also named
        both = addr & name_hits
        if len(both) == 1:
            return done(next(iter(both)), where + " + name")
        result["source"] = "ambiguous: " + " / ".join(sorted(addr))
        return result
    if len(name_hits) == 1:
        only = next(iter(name_hits))
        row = next(s for s in sites if s["name"] == only)
        corpus = settings.words(subject + " " + " ".join(blocks)) + " " + tw
        specific = [n for n in row["names"] if n != row["generic"] and settings.has_phrase(corpus, n)]
        sharing = [s["name"] for s in sites if row["place"] and s["place"] == row["place"]]
        if not specific and len(sharing) > 1:                       # only "PEARL ENERGY <TOWN>", and the town has several sites: never guess
            result["source"] = "ambiguous: the town has more than one site (" + " / ".join(sorted(sharing)) + ") and the document gives no address"
            return result
        return done(only, "site name")
    if len(name_hits) > 1:
        result["source"] = "ambiguous: " + " / ".join(sorted(name_hits))
        return result
    rec = _name_hits(sites, settings.words(to_name)) if to_name else set()
    if len(rec) == 1:
        return done(next(iter(rec)), "recipient name")
    if only_company:                                       # the company is known: the TOWN and the STREET named anywhere on the document are enough
        corpus = settings.words(subject + " " + " ".join(blocks)) + " " + tw
        both = []
        for s in sites:
            streets = {f.split()[1] for f in s["frags"] if len(f.split()) >= 2 and f.split()[0].isdigit()}
            if s.get("place") and settings.has_phrase(corpus, s["place"]) and any(settings.has_phrase(corpus, st) for st in streets):
                both.append(s["name"])
        if len(both) == 1:
            return done(both[0], "town + street named on the document")
    near = _fuzzy(sites, bw or tw[:1500])
    if near:
        return done(near, "site name (spelling differs)")
    if blocks:
        result["source"] = "not found (deliver-to says: '" + str(blocks[0])[:60] + "')"
    return result


# ----------------------------------------------------------------------------------------
# Importing the company's site Excel lists (new list = main, old list = extra names)
# ----------------------------------------------------------------------------------------

def _clean(text):
    return re.sub(r"\s+", " ", str(text or "")).strip()


def read_site_list(path, sheet=0):
    """Read one site Excel file (sheet = position or name). Returns [{'company','name','address','handler'}]."""
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True, read_only=True)
    try:
        return _read_site_sheet(wb[sheet] if isinstance(sheet, str) else wb[wb.sheetnames[sheet]])
    finally:
        wb.close()


def _read_site_sheet(ws):
    out, company, cols = [], "", None
    for raw in ws.iter_rows(values_only=True):
        cells = [_clean(c) for c in raw]
        joined = " ".join(c for c in cells if c)
        m = re.match(r"Company Name\s*:\s*(.+)", joined, re.IGNORECASE)
        if m:
            company = _clean(m.group(1))
            continue
        upper = [c.upper() for c in cells]
        if "SITE NAME" in upper and "ADDRESS" in upper:                  # a header row: find the columns
            site_col, addr_col = upper.index("SITE NAME"), upper.index("ADDRESS")
            cols = (site_col, addr_col, addr_col - 1)
            continue
        if cols and company:
            site_col, addr_col, handler_col = cols
            if len(cells) > addr_col and cells[site_col] and cells[addr_col]:
                out.append({"company": company, "name": cells[site_col], "address": cells[addr_col],
                            "handler": cells[handler_col] if handler_col < len(cells) and handler_col != site_col else ""})
    return out


def _key(address):
    f = make_fragments(address)
    return f[0] if f else settings.words(address)[:30]


def build_rows(new_sites, old_sites=None):
    """Build Sites-sheet rows and a problems report from the new (and old) lists."""
    rows, report = [], []
    by_key = {}
    for s in new_sites:
        by_key[_key(s["address"])] = s["name"]
    old_alias = {}
    for s in old_sites or []:
        k = _key(s["address"])
        if k in by_key and settings.words(by_key[k]) != settings.words(s["name"]):
            old_alias.setdefault(by_key[k], set()).add(s["name"])
    known = {settings.words(s["name"]) for s in new_sites}
    for s in new_sites:
        aliases = set(old_alias.get(s["name"], set()))
        fixed = " ".join(TYPO_FIXES.get(w, w) for w in settings.words(s["name"]).split())
        if fixed != settings.words(s["name"]):
            aliases.add(fixed.title())
        frags = make_fragments(s["address"])
        notes = ""
        if not frags:
            notes = "no address fragment could be built - add one by hand"
            report.append(f"{s['name']}: {notes}")
        rows.append([s["name"], s["company"], "|".join(sorted(aliases)), "|".join(frags), s["handler"], "", "Yes",
                     "new list", notes])
    for s in old_sites or []:                                            # sites only in the old list
        if _key(s["address"]) not in by_key and settings.words(s["name"]) not in known:
            frags = make_fragments(s["address"])
            rows.append([s["name"], s["company"], "", "|".join(frags), s["handler"], "", "Yes", "old list only",
                         "only in the old list - disable if the site is closed"])
            report.append(f"{s['name']}: only in the old list")
    # fragments shared by more than one site cannot tell them apart
    owners = {}
    for r in rows:
        for f in r[3].split("|"):
            if f:
                owners.setdefault(f, []).append(r[0])
    for f, names in owners.items():
        if len(names) > 1:
            report.append(f"fragment '{f}' is shared by: {', '.join(names)}")
    return rows, report


# ----------------------------------------------------------------------------------------
# Matching the site list to the REAL SharePoint folders
# ----------------------------------------------------------------------------------------

FUZZY_MIN, FUZZY_MARGIN = 0.90, 0.03


def _match(rows, folders):
    """Match sites to a list of folder names. Capitals, extra spaces and punctuation are ignored."""
    by_key = {}
    for f in folders:
        by_key.setdefault(settings.words(f), []).append(f)
    result = {"matched": {}, "fuzzy": {}, "ambiguous": {}, "missing": [], "extra": []}
    claimed, pending = set(), []
    for s in rows:
        hit = next((by_key[n] for n in s["names"] if by_key.get(n)), None)
        if hit is None and s["folder"]:
            hit = by_key.get(settings.words(s["folder"]))
        if hit is None:
            pending.append(s)
        elif len(hit) == 1:
            result["matched"][s["name"]] = hit[0]
            claimed.add(hit[0])
        else:
            result["ambiguous"][s["name"]] = hit
    free = [f for f in folders if f not in claimed]
    for s in pending:
        scored = sorted(((max(difflib.SequenceMatcher(None, n, settings.words(f)).ratio() for n in s["names"]), f) for f in free), reverse=True)
        if scored and scored[0][0] >= FUZZY_MIN and (len(scored) == 1 or scored[0][0] - scored[1][0] >= FUZZY_MARGIN):
            result["fuzzy"][s["name"]] = (scored[0][1], round(scored[0][0], 3))
            free.remove(scored[0][1])
        else:
            result["missing"].append(s["name"])
    result["extra"] = sorted(free)
    return result


def _subfolders(path):
    return sorted(d for d in os.listdir(path) if os.path.isdir(os.path.join(path, d)))


def _month_folder(root):
    """The folder of the current month under the root, else the most recently changed '... EMAIL INVOICES' folder."""
    import datetime
    from .router import month_folder_name
    current = os.path.join(root, month_folder_name(datetime.date.today()))
    if os.path.isdir(current):
        return current
    months = [os.path.join(root, d) for d in _subfolders(root) if d.upper().endswith("EMAIL INVOICES")]
    return max(months, key=os.path.getmtime) if months else None


def match_folders(root):
    """Compare the sites with the real folders. Capitals, extra spaces and punctuation are ignored.
    The layout decides where to look: for  month > company > site  the bot opens the current month's folder, then every company
    folder in it, and matches the sites of THAT company with the folders inside it; for a flat layout the sub-folders of `root`.
    Returns {'matched': {site: folder}, 'fuzzy': {site: (folder, ratio)}, 'ambiguous': {site: [folders]},
             'missing': [sites without a folder], 'extra': [folders without a site], 'all_folders': [...], 'looked_in': text}."""
    from .router import segment_kind, split_pattern
    segments = split_pattern(settings.run_value("Destination pattern", "{MONTH_FOLDER}\\{COMPANY}"))
    kinds = [segment_kind(x) for x in segments]
    rows = active_sites()
    if kinds[-2:] == ["company", "site"]:
        base = _month_folder(root) if kinds[0] == "month" and len(kinds) == 3 else root
        if base is None:
            return {"matched": {}, "fuzzy": {}, "ambiguous": {}, "missing": [s["name"] for s in rows], "extra": [], "all_folders": [], "looked_in": "no month folder found under the root"}
        out = {"matched": {}, "fuzzy": {}, "ambiguous": {}, "missing": [], "extra": [], "all_folders": [], "looked_in": base}
        seen_companies = set()
        for company_folder in _subfolders(base):
            crow = next((c for c in companies.all_active() if settings.words(c["Folder name"]) == settings.words(company_folder)), None)
            if crow is None:
                continue
            seen_companies.add(crow["Company name"])
            mine = [s for s in rows if company_of(s) == crow["Company name"]]
            part = _match(mine, _subfolders(os.path.join(base, company_folder)))
            out["all_folders"] += _subfolders(os.path.join(base, company_folder))
            for key in ("matched", "fuzzy", "ambiguous"):
                out[key].update(part[key])
            out["missing"] += part["missing"]
            out["extra"] += [f"{company_folder}\\{x}" for x in part["extra"]]
        out["missing"] += [s["name"] for s in rows if company_of(s) not in seen_companies]       # a company with no folder at all
        return out
    folders = _subfolders(root)
    out = _match(rows, folders)
    out.update(all_folders=folders, looked_in=root)
    return out


def apply_folder_matches(result, root):
    """Write the matched folder names into the 'SharePoint folder' column of the Sites sheet. A folder name typed by hand that
    still exists in `root` is never replaced. Returns the number of cells written."""
    existing = {f.lower() for f in (result.get("all_folders") or os.listdir(root))}
    wanted = {name: folder for name, folder in result["matched"].items()}
    wanted.update({name: folder for name, (folder, _r) in result["fuzzy"].items()})
    written = [0]

    def change(wb):
        ws = wb["Sites"]
        head = {str(ws.cell(1, c).value or "").strip(): c for c in range(1, ws.max_column + 1)}
        if "SharePoint folder" not in head:
            ws.cell(1, ws.max_column + 1).value = "SharePoint folder"
            head["SharePoint folder"] = ws.max_column
        name_col, folder_col = head["Site name"], head["SharePoint folder"]
        for r in range(2, ws.max_row + 1):
            name = str(ws.cell(r, name_col).value or "").strip()
            if name in wanted:
                current = str(ws.cell(r, folder_col).value or "").strip()
                if current and current.lower() in existing and current != wanted[name]:
                    continue
                if current != wanted[name]:
                    ws.cell(r, folder_col).value = wanted[name]
                    written[0] += 1

    settings.edit(change)
    return written[0]


# ----------------------------------------------------------------------------------------------- learned sites / folders
def ship_keys(ship_text):
    """The keys a ship-to block is remembered by: its deliver-to number ('num:64228817') and its address words ('addr:...')."""
    keys = []
    for block in [b for b in str(ship_text or "").split(" || ") if b.strip()]:
        if block.startswith("merchant:"):                    # a Fleet Card statement: its merchant number
            keys.append(block.strip())
            continue
        b = _untag(block)
        parts = [p.strip() for p in b.split("|") if p.strip()]
        if parts and re.fullmatch(r"\d{5,10}", parts[0]):
            keys.append("num:" + parts[0])
            parts = parts[1:]
        else:
            m = re.search(r"\b(\d{6,10})\b", b)
            if m and not looks_like_address(m.group(1)):
                keys.append("num:" + m.group(1))
        addr = settings.words(" ".join(parts[:3]))
        if len(addr.split()) >= 2:
            keys.append("addr:" + addr)
    return list(dict.fromkeys(keys))


MERCHANT = re.compile(r"merchant\s*(?:number|no\.?|id)\s*:?\s*(\d{4,12})", re.IGNORECASE)


def merchant_number(text):
    m = MERCHANT.search(str(text or ""))
    return m.group(1) if m else ""


def site_folders_in(month_dir):
    """[(company folder, site folder)] that exist inside a month folder (layout month > company > site)."""
    out = []
    try:
        for c in sorted(os.listdir(month_dir)):
            cp = os.path.join(month_dir, c)
            if os.path.isdir(cp):
                for s_ in sorted(os.listdir(cp)):
                    if os.path.isdir(os.path.join(cp, s_)):
                        out.append((c, s_))
    except OSError:
        pass
    return out


def folder_in_ship_text(ship_text, pairs):
    """The one (company folder, site folder) whose site-folder name (2+ words) is written in the ship-to text, else None.
    Two different folders that both fit = no answer (a person decides)."""
    tw = settings.words(str(ship_text or "").replace("|", " "))
    hits = set()
    for company, folder in pairs:
        fw = settings.words(folder)
        if len(fw.split()) >= 2 and (settings.has_phrase(tw, fw) or _order_free_hits([{"name": folder, "names": [fw]}], tw)):
            hits.add((company, folder))
    return hits.pop() if len(hits) == 1 else None
