"""Reads a PDF into lines of words WITH their positions, so the bot can find a value
to the right of a label, or in the cell directly below a table heading."""
import logging

import pdfplumber

for _name in ("pdfminer", "pdfplumber", "PIL"):      # a damaged PDF must not print one warning line per broken object
    logging.getLogger(_name).setLevel(logging.CRITICAL)

LINE_TOLERANCE = 4        # words whose baselines (bottoms) differ by less than this are on one line


class Line:
    def __init__(self, words, page):
        self.words = sorted(words, key=lambda w: w["x0"])
        self.page = page
        # vertical position of the line = its baseline (bottom), the same for big and small fonts
        self.top = sum(w.get("bottom", w["top"]) for w in self.words) / len(self.words)
        parts, self.spans, pos = [], [], 0
        for w in self.words:
            self.spans.append((pos, pos + len(w["text"])))
            parts.append(w["text"])
            pos += len(w["text"]) + 1
        self.text = " ".join(parts)

    def x_range(self, start, end):
        """Horizontal extent of the words covering text[start:end]."""
        first = last = None
        for w, (s, e) in zip(self.words, self.spans):
            if e > start and s < end:
                first = first or w
                last = w
        return (first["x0"], last["x1"]) if first else (0.0, 0.0)

    def max_gap(self, start, end):
        """Widest gap between neighbouring words inside text[start:end] (0 for one word)."""
        idx = [i for i, (s, e) in enumerate(self.spans) if e > start and s < end]
        gaps = [self.words[j]["x0"] - self.words[j - 1]["x1"] for j in idx[1:]]
        return max(gaps) if gaps else 0.0

    def prev_word_gap(self, start):
        """(previous word, gap in points) just before text position `start`."""
        for i, (s, e) in enumerate(self.spans):
            if e > start:
                if i == 0:
                    return None, 999.0
                prev = self.words[i - 1]
                return prev["text"], self.words[i]["x0"] - prev["x1"]
        return None, 999.0

    def cells(self, gap=7):
        """The line split into table cells: words closer than `gap` points belong together.
        Returns a list of (text, x0, x1)."""
        out = []
        for w in self.words:
            if out and w["x0"] - out[-1][2] <= gap:
                t, x0, _ = out[-1]
                out[-1] = (t + " " + w["text"], x0, w["x1"])
            else:
                out.append((w["text"], w["x0"], w["x1"]))
        return out

    def phrase_from(self, index, max_gap=20):
        """Text starting at word number `index`, following words kept while close together."""
        out, prev_x1 = [], None
        for w in self.words[index:]:
            if prev_x1 is not None and w["x0"] - prev_x1 > max_gap:
                break
            out.append(w["text"])
            prev_x1 = w["x1"]
        return " ".join(out)

    def phrase_at(self, x_lo, x_hi, max_gap=20):
        """Text starting at the first word that overlaps [x_lo, x_hi], following words kept
        while they are close together."""
        out, prev_x1 = [], None
        for w in self.words:
            if not out:
                if w["x1"] >= x_lo and w["x0"] <= x_hi:
                    out.append(w["text"])
                    prev_x1 = w["x1"]
            elif w["x0"] - prev_x1 <= max_gap:
                out.append(w["text"])
                prev_x1 = w["x1"]
            else:
                break
        return " ".join(out)


class Document:
    def __init__(self, lines, text, page_count=1):
        self.lines = lines
        self.text = text
        self.page_count = page_count          # pages in the PDF (may be more than were read)
        self.ocr = False                      # True when the text came from OCR

    @classmethod
    def synthetic(cls, rows):
        """Build a document from rows of (x, text) cells. Used by the self-test."""
        lines, plain = [], []
        for i, row in enumerate(rows):
            words = []
            for x, cell in row:
                cursor = x
                for token in cell.split():
                    width = 6 * len(token)
                    words.append({"text": token, "x0": cursor, "x1": cursor + width,
                                  "top": 20.0 * i, "bottom": 20.0 * i + 10})
                    cursor += width + 6
            lines.append(Line(words, 0))
            plain.append(" ".join(w["text"] for w in words))
        return cls(lines, "\n".join(plain))

    def header(self, n=15):
        """The first n non-empty lines of text (where titles and the supplier name are)."""
        return "\n".join([l for l in self.text.splitlines() if l.strip()][:n])


TOUCHING = 1.3            # points: letters of one word that a PDF delivered as separate "words" touch each other


def join_touching(group):
    """Some PDFs (e.g. Bonfect totals) are delivered letter by letter: 'N e t A m o u n t'. Letters that touch belong to one word;
    a real space is at least ~2 points wide, so words that are really separate stay separate."""
    group = sorted(group, key=lambda w: w["x0"])
    out = []
    for w in group:
        if out and -1.5 <= w["x0"] - out[-1]["x1"] <= TOUCHING and (len(out[-1]["text"]) <= 3 or len(w["text"]) <= 3):
            prev = out[-1]
            out[-1] = dict(prev, text=prev["text"] + w["text"], x1=w["x1"], bottom=max(prev["bottom"], w["bottom"]),
                           top=min(prev["top"], w["top"]))
        else:
            out.append(w)
    return out


def build_lines(words, page_no=0):
    """Group words into lines by baseline. Words whose bottoms are within LINE_TOLERANCE share a line,
    whatever their font size."""
    words = sorted(words, key=lambda w: (w["bottom"], w["x0"]))
    lines, group, anchor = [], [], None
    for w in words:
        if anchor is None or abs(w["bottom"] - anchor) <= LINE_TOLERANCE:
            group.append(w)
            anchor = w["bottom"] if anchor is None else anchor
        else:
            lines.append(Line(join_touching(group), page_no))
            group, anchor = [w], w["bottom"]
    if group:
        lines.append(Line(join_touching(group), page_no))
    return lines


ZERO_WIDTH = "\u200b\u200c\u200d\ufeff"


def _clean_words(words):
    """Invisible characters (zero-width spaces) that some invoice templates put before every value are removed."""
    out = []
    for w in words:
        t = w["text"].translate({ord(c): None for c in ZERO_WIDTH})
        if t:
            out.append(w if t == w["text"] else dict(w, text=t))
    return out


def load(path, max_pages=None):
    """Read a PDF (the first `max_pages` pages, or all). Returns a Document.
    Pages are read one at a time and their cache is cleared, which keeps memory low on long files."""
    lines, texts = [], []
    with pdfplumber.open(path) as pdf:
        total = len(pdf.pages)
        pages = pdf.pages if max_pages is None else pdf.pages[:max_pages]
        for page_no, page in enumerate(pages):
            texts.append((page.extract_text() or "").translate({ord(c): None for c in ZERO_WIDTH}))
            lines += build_lines(_clean_words(page.extract_words(x_tolerance=2, y_tolerance=2) or []), page_no)
            try:
                page.flush_cache()
            except Exception:
                pass
    return Document(lines, "\n".join(texts), page_count=total)


def document_from_ocr_pages(pages_words):
    """Build a Document from OCR output: a list (one per page) of word dicts with text/x0/x1/top/bottom."""
    lines, texts = [], []
    for page_no, words in enumerate(pages_words):
        page_lines = build_lines(words, page_no)
        lines += page_lines
        texts.append("\n".join(l.text for l in page_lines))
    doc = Document(lines, "\n".join(texts), page_count=len(pages_words))
    doc.ocr = True
    return doc
