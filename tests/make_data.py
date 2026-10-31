"""Generates the test documents (needs reportlab + Pillow; the finished files are shipped in tests/data,
so the tests run without reportlab).   python tests/make_data.py"""
import os
import random

from PIL import Image, ImageDraw, ImageFont
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data")
FONT_BOLD = next((f for f in ("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
                               r"C:\Windows\Fonts\arialbd.ttf") if os.path.exists(f)), None)


def _font(size):
    return ImageFont.truetype(FONT_BOLD, size) if FONT_BOLD else ImageFont.load_default()


def invoice(path, supplier="Acme Foods Pty Ltd", number="INV-1001", date="02/10/2026", net=100.00, title="TAX INVOICE",
            number_label="Invoice No", date_label="Invoice Date", customer=("PEARL ENERGY COLYTON",),
            address="88-90 GREAT WESTERN HWY COLYTON NSW 2760", gst=None, total=None, pages=1, show_supplier=True, totals=True,
            logo_text=None, filler=0):
    """A simple label-based invoice. Totals: Total excl GST / Total GST / Total inc GST."""
    gst = round(net * 0.1, 2) if gst is None else gst
    total = net + gst if total is None else total
    c = canvas.Canvas(path, pagesize=A4)
    for pg in range(pages):
        y = 800
        if logo_text and pg == 0:
            im = Image.new("RGB", (640, 110), (0, 40, 120))
            ImageDraw.Draw(im).text((20, 30), logo_text, fill=(255, 255, 255), font=_font(40))
            png = os.path.join(os.path.dirname(path), "_logo_tmp.png")
            im.save(png)
            c.drawImage(png, 40, 740, width=230, height=40)
            os.remove(png)
            y = 725
        if show_supplier:
            c.setFont("Helvetica-Bold", 14); c.drawString(40, y, supplier); y -= 16
            c.setFont("Helvetica", 9); c.drawString(40, y, "ABN 12 345 678 901"); y -= 24
        c.setFont("Helvetica-Bold", 13); c.drawString(40, y, title); y -= 22
        c.setFont("Helvetica", 9)
        c.drawString(40, y, f"{number_label}: {number}"); y -= 13
        c.drawString(40, y, f"{date_label}: {date}"); y -= 22
        c.drawString(40, y, "Deliver To"); y -= 12
        for line in customer:
            c.drawString(40, y, line); y -= 12
        c.drawString(40, y, address); y -= 24
        c.drawString(40, y, "Qty"); c.drawString(90, y, "Description"); c.drawString(420, y, "Unit Price"); c.drawString(500, y, "Line Total"); y -= 13
        for i in range(3 + filler):
            c.drawString(40, y, str(i + 1)); c.drawString(90, y, f"Goods item {i + 1}")
            c.drawRightString(470, y, f"{net / 3:.2f}"); c.drawRightString(550, y, f"{net / 3:.2f}"); y -= 12
        if totals and pg == pages - 1:
            c.setFont("Helvetica", 9)
            c.drawString(330, 150, "Total excl GST"); c.drawRightString(550, 150, f"{net:,.2f}")
            c.drawString(330, 136, "Total GST"); c.drawRightString(550, 136, f"{gst:,.2f}")
            c.drawString(330, 122, "Total inc GST"); c.drawRightString(550, 122, f"{total:,.2f}")
        c.showPage()
    c.save()


def simple(path, title, lines, supplier="Zed Supplies Pty Ltd"):
    """A text document (POD, statement, account summary ...)."""
    c = canvas.Canvas(path, pagesize=A4)
    y = 800
    c.setFont("Helvetica-Bold", 14); c.drawString(40, y, supplier); y -= 16
    c.setFont("Helvetica", 9); c.drawString(40, y, "ABN 98 765 432 109"); y -= 26
    c.setFont("Helvetica-Bold", 13); c.drawString(40, y, title); y -= 24
    c.setFont("Helvetica", 9)
    for ln in lines:
        c.drawString(40, y, ln); y -= 13
    c.showPage(); c.save()


def scanned(path, number="SC-7001", net=200.00):
    """A picture of an invoice: no text layer at all."""
    gst = round(net * 0.1, 2)
    im = Image.new("RGB", (1240, 1754), "white")
    d = ImageDraw.Draw(im)
    rows = [("Scan Wholesale Pty Ltd", 60), ("TAX INVOICE", 62), (f"Invoice No: {number}", 40), ("Invoice Date: 02/10/2026", 40),
            ("Deliver To", 36), ("PEARL ENERGY COLYTON", 36), ("", 20),
            (f"Total excl GST   {net:,.2f}", 40), (f"Total GST   {gst:,.2f}", 40), (f"Total inc GST   {net + gst:,.2f}", 40)]
    y = 80
    for text, size in rows:
        d.text((80, y), text, fill="black", font=_font(size))
        y += size + 22
    im.save(path, "PDF", resolution=150.0)


def corrupt(path):
    with open(path, "wb") as f:
        f.write(b"%PDF-1.4\nthis is not really a pdf " + bytes(random.Random(1).randrange(256) for _ in range(200)))


def excel_invoice(path, number="XL-5001", net=100.00):
    from openpyxl import Workbook
    wb = Workbook(); ws = wb.active
    gst = round(net * 0.1, 2)
    ws.append(["Beta Distributors Pty Ltd"]); ws.append(["TAX INVOICE"])
    ws.append(["Invoice No", number]); ws.append(["Invoice Date", "02/10/2026"])
    ws.append(["Deliver To", "PEARL ENERGY COLYTON"]); ws.append([])
    ws.append(["Qty", "Description", "Amount"]); ws.append([2, "Goods", net])
    ws.append(["Total excl GST", net]); ws.append(["Total GST", gst]); ws.append(["Total inc GST", net + gst])
    for r in range(9, 12):
        ws.cell(r, 2).number_format = "#,##0.00"
    wb.save(path)


def unilever(path, number, date, bill_no, bill, ship_no, ship, customer_no=None):
    """The Unilever layout: 'Invoice to <no>' and 'Deliver to <no>' side by side, each with an address box below (5.11.0)."""
    c = canvas.Canvas(path, pagesize=A4)
    c.setFont("Helvetica-Bold", 18); c.drawString(110, 800, "Unilever Australia Ltd.")
    c.setFont("Helvetica", 9); c.drawString(110, 782, "219 North Rocks Road North Rocks NSW 2151")
    c.drawString(110, 768, "A.B.N.: 66004050828")
    c.setFont("Helvetica-Bold", 14); c.drawString(400, 800, "UNILEVER TAX INVOICE")
    c.setFont("Helvetica", 9)
    c.drawString(390, 740, "Customer Number"); c.drawString(470, 740, customer_no or bill_no)
    c.drawString(390, 726, "Invoice Number"); c.drawString(470, 726, number)
    c.drawString(390, 712, "Invoice Date"); c.drawString(470, 712, date)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(60, 700, "Invoice to"); c.drawString(120, 700, bill_no)
    c.drawString(250, 700, "Deliver to"); c.drawString(310, 700, ship_no)
    c.setFont("Helvetica", 9)
    y = 686
    for i in range(max(len(bill), len(ship))):
        if i < len(bill):
            c.drawString(60, y, bill[i])
        if i < len(ship):
            c.drawString(250, y, ship[i])
        y -= 11
    y -= 20
    c.drawString(40, y, "Delivered Qty"); c.drawString(120, y, "Product Description"); c.drawString(420, y, "GST"); c.drawString(490, y, "Invoice Total")
    y -= 13
    c.drawString(40, y, "1"); c.drawString(120, y, "MAGNUM CLASSIC OOH 20X100ML"); c.drawRightString(450, y, "5.50"); c.drawRightString(550, y, "60.50")
    c.drawString(330, 150, "Total excl GST"); c.drawRightString(550, 150, "55.00")
    c.drawString(330, 136, "Total GST"); c.drawRightString(550, 136, "5.50")
    c.drawString(330, 122, "Total inc GST"); c.drawRightString(550, 122, "60.50")
    c.showPage()
    c.save()


def build_unilever_cases():
    """The two invoices the payments team reported in October 2026 (filed under the BILL-TO site by 5.10.0)."""
    p = lambda n: os.path.join(DATA, n)
    unilever(p("unilever_billto_denham_court_shipto_preston.pdf"), "3506013060", "14.10.2026",
             "15496413", ["AMPOL DENHAM COURT", "GOPINATH INVESTMENT", "194 CAMPBELLTOWN ROAD", "DENHAM COURT NSW AU 2565"],
             "64228817", ["PEARL ENERGY PRESTON", "1B BELL STREET", "PRESTON VIC AU 3072"])
    unilever(p("unilever_billto_harlaxton_shipto_bp_north_orange.pdf"), "3506011682", "13.10.2026",
             "15537500", ["PEARL ENERGY HARLAXTON", "PEARL ENERGY NO.3", "6 TO 10 RUTHVEN STREET", "HARLAXTON QLD AU 4350"],
             "15471177", ["BP NORTH ORANGE", "PEARL ENERGY NO.3", "5 HANRAHAN PLACE", "ORANGE NSW AU 2800"])


def build_all():
    os.makedirs(DATA, exist_ok=True)
    p = lambda n: os.path.join(DATA, n)
    invoice(p("invoice_acme.pdf"), "Acme Foods Pty Ltd", "INV-1001", net=100.00)
    invoice(p("invoice_bidfood.pdf"), "Bidfood Australia Ltd", "778899", net=250.00, number_label="Document No")
    invoice(p("invoice_other.pdf"), "Orbit Wholesale Pty Ltd", "OW-3030", net=80.00)
    invoice(p("invoice_company_named.pdf"), "Nova Snacks Pty Ltd", "NS-4040", net=60.00,
            customer=("RADHE KRISHNA INVESTMENT PTY LTD", "PEARL ENERGY ROBINA"), address="299 SCOTTSDALE DRIVE ROBINA QLD 4226")
    invoice(p("credit_note.pdf"), "Acme Foods Pty Ltd", "CN-9001", net=40.00, title="CREDIT NOTE", number_label="Credit Note No")
    invoice(p("invoice_logo_only.pdf"), show_supplier=False, number="RW-2020", net=90.00, logo_text="RARITY WHOLESALE")
    invoice(p("invoice_20_pages.pdf"), "Longform Pty Ltd", "LF-1", net=30.00, pages=20)
    invoice(p("invoice_bad_math.pdf"), "Shady Supplies Pty Ltd", "SH-55501", net=100.00, gst=10.00, total=150.00)  # 100 + 10 is not 150
    simple(p("pod.pdf"), "PROOF OF DELIVERY", ["Delivery No: D-77120", "Date: 02/10/2026", "Delivered to PEARL ENERGY COLYTON",
                                               "Received in good order"], "Haulage Express Pty Ltd")
    simple(p("statement.pdf"), "WEEKLY STATEMENT", ["Statement Date: 12/08/2026", "PEARL ENERGY COLYTON",
                                                    "1/07/2026 F55527024 Sale $60.52", "Total Amount Due $2,929.27"], "Fresh Fruit Co Pty Ltd")
    simple(p("account_summary.pdf"), "ACCOUNT SUMMARY", ["Account Number 1458043", "Date: 02/10/2026", "PEARL ENERGY COLYTON",
                                                         "Balance $2,651.91"], "Drinks Group Pty Ltd")
    simple(p("statement_rarity.pdf"), "WEEKLY STATEMENT", ["Statement Date: 12/08/2026", "PEARL ENERGY COLYTON",
                                                         "Total Amount Due $2,929.27"], "Rarity Wholesale")
    scanned(p("scanned_invoice.pdf"))
    corrupt(p("corrupt.pdf"))
    excel_invoice(p("invoice_excel.xlsx"))
    build_unilever_cases()
    print("generated", len(os.listdir(DATA)), "files in", DATA)


if __name__ == "__main__":
    build_all()
