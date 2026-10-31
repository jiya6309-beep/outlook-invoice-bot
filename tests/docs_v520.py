"""Rebuilt layouts for 5.20.0 (the _Unassigned review): Asahi credit note for 'MOBIL PORT MACQUARIE HASTINGS', PFD tax invoice for
'PEARL ENERGY MITTAGONG', Slick Chicks for 'Pearl Energy Beresfield South', a Neverfail statement and Fleet Card statements whose
address is the owner's postal address."""
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas


def _lines(c, x, y, lines, step=12):
    for i, t in enumerate(lines):
        c.drawString(x, y - step * i, t)


def asahi_credit(path):
    c = canvas.Canvas(path, pagesize=A4)
    c.setFont("Helvetica-Bold", 14)
    c.drawString(40, 800, "Asahi")
    c.drawString(220, 790, "CREDIT ADJUSTMENT NOTE")
    c.setFont("Helvetica", 9)
    c.drawString(420, 795, "9018622910   /   23.07.2026")
    c.drawString(250, 760, "Asahi Beverages Pty Ltd")
    c.drawString(40, 730, "Invoiced To:")
    _lines(c, 110, 730, ["PEARL ENERGY NO1 PTY LTD (PAYER", "TTF PEARL ENERGY NO. 1 FUT", "20B LEXINGTON DR, SUITE 202", "NSW 2153"])
    c.drawString(40, 660, "Deliver To:")
    _lines(c, 110, 660, ["MOBIL PORT MACQUARIE HASTINGS", "128 HASTINGS RIVER DRIVE", "PORT MACQUARIE", "NSW 2444"])
    _lines(c, 380, 730, ["A.B.N.:  51 004 243 994", "Credit No: 9018622910", "Date: 23/07/2026"])
    c.drawString(260, 560, "Total AUD")
    for x, t in zip((300, 345, 395, 445, 490), ("20.67", "1.68", "22.35", "2.24", "24.59")):
        c.drawString(x, 560, t)
    c.save()


def pfd_mittagong(path):
    c = canvas.Canvas(path, pagesize=A4)
    c.setFont("Helvetica-Bold", 14)
    c.drawString(200, 800, "PFD Food Services Pty Ltd")
    c.drawString(240, 784, "TAX INVOICE")
    c.setFont("Courier", 8)
    c.drawString(40, 750, "SOLD TO: 50068210  ATTN: PRATIK PATEL")
    c.drawString(300, 750, "SHIP TO: 50068211  ATTN: PRATIL PATEL")
    _lines(c, 40, 736, ["PEARL ENERGY MITTAGONG", "65 BOWRAL ROAD", "MITTAGONG    NSW 2575"])
    _lines(c, 300, 736, ["PEARL ENERGY MITTAGONG", "65 BOWRAL ROAD", "MITTAGONG NSW 2575"])
    c.setFont("Helvetica", 8)
    _lines(c, 450, 790, ["Invoice Date  07 Oct 26", "Invoice No  LW437327"])
    c.drawString(40, 600, "013556  180GX24 KING SIZE SAUSAGE ROLLS  2.00  62.10  124.20  12.42  136.62")
    _lines(c, 400, 200, ["Net Line Total  $311.46", "TOTAL GST  $31.41", "INVOICE TOTAL (GST Incl.)  $345.55"])
    c.save()


def slick_chicks(path):
    c = canvas.Canvas(path, pagesize=A4)
    c.setFont("Helvetica-Bold", 14)
    c.drawString(40, 800, "Slick Chicks")
    c.setFont("Helvetica", 9)
    c.drawString(380, 800, "Tax Invoice")
    c.drawString(380, 780, "Invoice #:  864259")
    c.drawString(40, 785, "ABN: 17076782537")
    c.drawString(40, 740, "Bill To:")
    c.drawString(320, 740, "Ship To:")
    _lines(c, 40, 720, ["Pearl Energy Beresfield South", "Pearl Energy No. 1 Pty Ltd atf Pearl Energy Fixed Unit Trust6", "Yangan Drive",
                        "BERESFIELD NSW 2322"])
    _lines(c, 320, 720, ["Pearl Energy Beresfield South", "6 Yangan Drive", "Beresfield  NSW  2322"])
    c.drawString(450, 660, "DATE  22/07/2026")
    _lines(c, 400, 200, ["SALE AMT.  $233.00", "GST  $0.35", "TOTAL AMT.  $233.35"])
    c.save()


def neverfail_statement(path):
    c = canvas.Canvas(path, pagesize=A4)
    c.setFont("Helvetica", 9)
    _lines(c, 420, 800, ["PO Box 1895", "North Sydney NSW 2059", "AUS", "ABN 43003559519"], 11)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(40, 740, "Bill To")
    c.setFont("Helvetica", 9)
    _lines(c, 40, 726, ["Pearl Energy Wyong", "300 Pacific Highway", "Wyong NSW 2259", "AUS"])
    c.drawString(400, 740, "Date")
    c.drawString(500, 740, "05/10/2026")
    c.drawString(400, 726, "Account #")
    c.drawString(500, 726, "100084274")
    c.drawString(40, 660, "Overdue    Current    Total Amount Due")
    c.drawString(40, 648, "$0.00    $475.47    $475.47")
    c.drawString(40, 610, "Account  Delivery Address  Invoice Date  Invoice Reference  Due Date  Balance Due")
    c.drawString(40, 596, "100084274  1-10 Amy Cl Wyong NSW AUS  11/09/2026  INV-002696872  11/10/2026  $293.13")
    c.drawString(40, 584, "100084274  1-10 Amy Cl Wyong NSW AUS  24/09/2026  INV-002711686  24/10/2026  $182.34")
    c.drawString(300, 570, "Total for Account # 100084274   $475.47")
    c.drawString(40, 100, "Neverfail Springwater Proprietary Limited is a wholly-owned subsidiary of Coca-Cola Europacific Partners")
    c.save()


def fleet_statement(path, name, merchant, address):
    c = canvas.Canvas(path, pagesize=A4)
    c.setFont("Helvetica-Bold", 16)
    c.drawString(40, 800, "BUSINESS FUEL CARDS")
    c.drawString(330, 800, "Merchant Payment Advice")
    c.setFont("Helvetica", 7)
    _lines(c, 420, 786, ["Business Fuel Cards Pty Ltd", "ABN No: 99 161 721 106", "GPO Box 3852, Sydney NSW 2001"], 9)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(260, 750, "Recipient Created Tax Invoice / Tax Invoice")
    c.drawString(40, 720, name)
    c.drawString(350, 720, "Merchant Number:")
    c.drawString(500, 720, merchant)
    c.setFont("Helvetica", 8)
    c.drawString(350, 708, "Report Issue Date:")
    c.drawString(500, 708, "20 Oct 26")
    c.setFont("Helvetica", 9)
    _lines(c, 40, 690, ["Attention: Ronak Patel"] + address)
    c.drawString(40, 600, "The Total Payment due will be direct credited into your bank account on the 21 Oct 26")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(40, 560, "TOTAL PAYMENT DUE")
    c.drawString(480, 560, "$2,404.81")
    c.save()
