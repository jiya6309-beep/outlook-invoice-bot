"""Reading text from pictures: supplier logos and scanned PDFs (Tesseract OCR).

OCR is optional. If Tesseract is not installed the bot still works; scanned files simply go to
_Unrecognised with the reason, and logo-only suppliers fall back to the sender hint."""
import os
import shutil

from . import pdftext

_cfg = {"enabled": True, "tesseract": "", "max_pages": 3}
_state = {"checked": False, "ok": False, "why": "not checked"}

WINDOWS_PATHS = [r"C:\Program Files\Tesseract-OCR\tesseract.exe", r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"]


def configure(enabled=True, tesseract_path="", max_pages=3):
    _cfg.update(enabled=bool(enabled), tesseract=tesseract_path or "", max_pages=int(max_pages or 3))
    _state.update(checked=False, ok=False, why="not checked")


def config():
    """The current options, in the form configure() accepts (handed to the reader processes)."""
    return {"enabled": _cfg["enabled"], "tesseract_path": _cfg["tesseract"], "max_pages": _cfg["max_pages"]}


def available():
    """True if OCR can run. The reason is in why_unavailable()."""
    if _state["checked"]:
        return _state["ok"]
    _state["checked"] = True
    if not _cfg["enabled"]:
        _state.update(ok=False, why="OCR is switched off in Bot_Settings.xlsx (Run -> OCR enabled)")
        return False
    try:
        import pytesseract
        from PIL import Image  # noqa: F401
    except Exception:
        _state.update(ok=False, why="Python packages pytesseract/Pillow are not installed")
        return False
    exe = _cfg["tesseract"] or shutil.which("tesseract") or next((p for p in WINDOWS_PATHS if os.path.exists(p)), "")
    if not exe:
        _state.update(ok=False, why="Tesseract program not found (see README, 'Optional: OCR')")
        return False
    pytesseract.pytesseract.tesseract_cmd = exe
    try:
        pytesseract.get_tesseract_version()
    except Exception as e:
        _state.update(ok=False, why=f"Tesseract could not start: {e}")
        return False
    _state.update(ok=True, why="")
    return True


def why_unavailable():
    available()
    return _state["why"]


def _prep(img):
    from PIL import ImageOps, ImageStat
    g = ImageOps.grayscale(img)
    scale = 3 if max(g.size) < 900 else 2
    g = g.resize((g.width * scale, g.height * scale))
    if ImageStat.Stat(g).mean[0] < 110:                    # light text on a dark background
        g = ImageOps.invert(g)
    return ImageOps.autocontrast(g)


def logo_text(pdf_path, max_images=3):
    """Text read from the images in the top third of page 1 (the logo). '' if none or OCR unavailable."""
    if not available():
        return ""
    import pdfplumber
    import pytesseract
    out = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[0]
            imgs = [i for i in page.images if i["top"] < page.height * 0.35
                    and 40 < (i["x1"] - i["x0"]) < page.width * 0.6
                    and 15 < (i["bottom"] - i["top"]) < page.height * 0.25]
            for im in imgs[:max_images]:
                box = (max(im["x0"], 0), max(im["top"], 0), min(im["x1"], page.width), min(im["bottom"], page.height))
                crop = page.crop(box).to_image(resolution=300).original
                for psm in (7, 6):
                    txt = " ".join(pytesseract.image_to_string(_prep(crop), config=f"--psm {psm}").split())
                    if txt:
                        out.append(txt)
    except Exception:
        return " ".join(out)
    return " ".join(out)


def ocr_document(pdf_path, max_pages=None, dpi=200):
    """OCR the first pages of a scanned PDF into a Document (words with positions), so the normal
    reader can use it. Raises RuntimeError if OCR is not available."""
    if not available():
        raise RuntimeError(why_unavailable())
    import pdfplumber
    import pytesseract
    pages_words = []
    scale = 72.0 / dpi
    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)
        for page in pdf.pages[: (max_pages or _cfg["max_pages"])]:
            img = page.to_image(resolution=dpi).original
            data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT)
            words = []
            for i, txt in enumerate(data["text"]):
                txt = (txt or "").strip()
                try:
                    conf = float(data["conf"][i])
                except (TypeError, ValueError):
                    conf = -1
                if txt and conf >= 0:
                    x0, top = data["left"][i] * scale, data["top"][i] * scale
                    words.append({"text": txt, "x0": x0, "x1": x0 + data["width"][i] * scale,
                                  "top": top, "bottom": top + data["height"][i] * scale})
            pages_words.append(words)
            try:
                page.flush_cache()
            except Exception:
                pass
    doc = pdftext.document_from_ocr_pages(pages_words)
    doc.page_count = total
    return doc
