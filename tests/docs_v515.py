"""Rebuilt layouts for 5.15.0 (from the payments team's screenshots): a Bega / BDD Australia invoice with 'Invoice to' (the payer)
and 'Sold to' (the real site) side by side and NO ship-to, and an Asahi invoice whose totals line has six columns
(Net Total ex.CDL ... Total Payable) after the label 'Total AUD'."""
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas


def bega_invoice(path, sold_lines=None, invoice_lines=None):
    sold_lines = sold_lines or ["Account: 217187", "Ampol Kelso", "The Trustee For Gopinath", "Investment", "Att: Nikesh Patel",
                                "17 Great Western Hwy", "KELSO NSW 2795"]
    invoice_lines = invoice_lines or ["Payer Account: 215194", "Mobil Jimboomba", "The Trustee For Gopinath Investment", "Att: Sara",
                                      "PO Box 7501", "BAULKHAM HILLS QLD 2153"]
    c = canvas.Canvas(path, pagesize=A4)
    c.setFont("Helvetica-Bold", 16)
    c.drawString(150, 800, "TAX INVOICE")
    c.setFont("Helvetica-Bold", 11)
    c.drawString(280, 800, "BDD Australia Pty Ltd")
    c.setFont("Helvetica", 8)
    c.drawString(280, 786, "ABN: 68 083 019 390")
    c.drawString(280, 770, "As Agent for:")
    for i, t in enumerate(("BDD Milk Pty Limited", "Berri Pty Ltd", "Bega Cheese Limited")):
        c.drawString(280, 760 - 10 * i, t)
    c.drawString(450, 800, "BDD Australia Pty Ltd")
    c.drawString(450, 790, "PO Box 23084")
    c.drawString(450, 780, "DOCKLANDS VIC 8012")
    c.setFont("Helvetica-Bold", 9)
    c.drawString(40, 720, "Invoice No.")
    c.drawString(110, 720, "577967880")
    c.drawString(40, 708, "Invoice Date")
    c.drawString(110, 708, "16/10/2026")
    c.drawString(40, 680, "Invoice to:")
    c.drawString(300, 680, "Sold to:")
    c.drawString(450, 680, "Additional Information:")
    c.setFont("Helvetica", 9)
    for i, t in enumerate(invoice_lines):
        c.drawString(40, 664 - 12 * i, t)
    for i, t in enumerate(sold_lines):
        c.drawString(300, 664 - 12 * i, t)
    for i, t in enumerate(("Delivery Date: 16/10/2026", "Delivery No.: 117988", "PO/Claim No. PO_217187_191026")):
        c.drawString(450, 664 - 12 * i, t)
    c.setFont("Helvetica-Bold", 8)
    y = 560
    for x, t in ((40, "Product Code"), (100, "Description"), (260, "Sales Qty"), (310, "Unit Price"), (360, "CDL"),
                 (420, "Net Value"), (480, "Total incl Taxes")):
        c.drawString(x, y, t)
    c.setFont("Helvetica", 8)
    rows = (("8953", "Dare Triple Espresso 500mL", "6", "2.35", "0.84", "14.94", "16.43"),
            ("8476", "DAILY Jce B/fast 500mL", "6", "2.18", "0.84", "13.92", "13.92"),
            ("7770", "Dare Mocha 500mL", "6", "2.35", "0.84", "14.94", "16.43"),
            ("7773", "Dare Espresso 750mL", "6", "2.82", "0.84", "17.76", "19.55"),
            ("6900", "BER LL Jce App 250mL", "12", "1.53", "1.68", "20.04", "20.04"),
            ("2168", "Dare Intense Espresso 500mL", "6", "2.35", "0.84", "14.94", "16.43"))
    for n, r in enumerate(rows):
        for x, t in zip((40, 100, 260, 310, 360, 420, 480), r):
            c.drawString(x, y - 14 * (n + 1), t)
    c.drawString(380, 400, "Total Net Value")
    c.drawString(500, 400, "96.54")
    c.drawString(380, 388, "Total GST included")
    c.drawString(500, 388, "6.26")
    c.drawString(380, 376, "Total Due")
    c.drawString(500, 376, "102.80")
    c.showPage()
    c.save()


def asahi_invoice(path, readable_headings=True):
    """readable_headings=False: the headings run into each other (as text extraction sometimes gives them), so only the maths
    of the totals row can find the amount payable."""
    c = canvas.Canvas(path, pagesize=A4)
    c.setFont("Helvetica-Bold", 16)
    c.drawString(40, 800, "Asahi")
    c.drawString(250, 800, "TAX INVOICE")
    c.setFont("Helvetica", 9)
    c.drawString(400, 800, "9019190928   /   23.10.2026")
    c.drawString(40, 784, "Asahi Lifestyle Beverages")
    c.drawString(250, 784, "Asahi Beverages Pty Ltd")
    c.drawString(40, 770, "ABN 31 004 243 994")
    c.drawString(40, 750, "Invoice No: 0814554380/16")
    c.drawString(300, 750, "Invoice Date: 23/10/2026")
    c.drawString(40, 730, "Deliver To:")
    c.drawString(40, 718, "Pearl Energy North Wyong")
    c.drawString(40, 706, "1 Pacific Highway")
    c.drawString(40, 694, "North Wyong NSW 2259")
    c.setFont("Helvetica-Bold", 7)
    cols = (40, 90, 120, 150, 260, 300, 345, 395, 445, 490, 535)
    heads = (("Material No.", ""), ("Qty", ""), ("UoM", ""), ("Description", ""), ("Net Unit", "Ex.CDL"), ("Net Total", "ex.CDL"),
             ("Cont.Dep.Le", "vy(CDL)"), ("Net Total ex.", "GST"), ("Total GST", ""), ("Total", "Payable"))
    for x, (t1, t2) in zip(cols, heads):         # wrapped over two lines, as printed
        if readable_headings:
            c.drawString(x, 650, t1)
            c.drawString(x, 641, t2)
        else:
            c.drawString(x, 645, (t1 + " " + t2).strip())
    c.setFont("Helvetica", 7)
    rows = (("10008623", "1", "CS", "600ML X 24 PET PEPSI COLA 7G", "49.71", "49.71", "3.36", "53.07", "5.31", "58.38"),
            ("10008624", "1", "CS", "1.25L X 12 PET PEPSI COLA 7G", "17.50", "17.50", "1.68", "19.18", "1.92", "21.10"),
            ("10008626", "2", "CS", "440ML X 24 CAN PEPSI COLA 7G", "33.38", "66.76", "6.72", "73.48", "7.35", "80.83"),
            ("10008632", "1", "CS", "1.25L X 12 PET MT DEW ENERGISED 7G", "17.50", "17.50", "1.68", "19.18", "1.92", "21.10"))
    for n, r in enumerate(rows):
        for x, t in zip(cols, r):
            c.drawString(x, 640 - 14 * (n + 1), t)
    c.setFont("Helvetica", 8)
    y = 560
    c.drawString(40, y, "No of Cartons:  35")
    c.drawString(150, y, "Total Weight:  470.443 KG")
    c.drawString(260, y, "Total AUD")
    for x, t in zip((300, 345, 395, 445, 490), ("1,125.56", "94.08", "1,219.64", "116.80", "1,336.44")):
        c.drawString(x, y, t)
    c.showPage()
    c.save()
