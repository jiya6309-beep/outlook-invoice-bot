"""Excel invoices: turned into rows and columns of words so the same reader can use them."""
import datetime
import re

from .pdftext import Document

MONEY_ROW = re.compile(r"total|gst|amount|balance|subtotal|tax|price|due|credit|payable", re.IGNORECASE)


def _fmt(cell, money_row=False):
    v = cell.value
    if v is None:
        return ""
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%d/%m/%Y")
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, (int, float)):
        pattern = str(getattr(cell, "number_format", "") or "")
        if money_row or "0.00" in pattern or (isinstance(v, float) and not v.is_integer()):
            return f"{v:,.2f}"
        return str(int(v)) if float(v).is_integer() else str(v)
    return str(v).strip()


def load(path, max_rows=400, max_cols=30, max_sheets=3):
    """Read an .xlsx/.xlsm file. Raises on unreadable files."""
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    rows = []
    try:
        for ws in wb.worksheets[:max_sheets]:
            for n, row in enumerate(ws.iter_rows(max_col=max_cols), start=1):
                if n > max_rows:
                    break
                money_row = any(isinstance(c.value, str) and MONEY_ROW.search(c.value) for c in row)
                cells = [(i * 90, _fmt(c, money_row)) for i, c in enumerate(row) if _fmt(c, money_row)]
                if cells:
                    rows.append(cells)
    finally:
        wb.close()
    doc = Document.synthetic(rows) if rows else Document([], "")
    return doc
