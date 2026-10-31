"""Rebuilt layouts for 5.16.0 (from the payments team's screenshots): a MobiGo Tax Invoice (only 'Order #'), a MobiGo order printout
(no 'Tax Invoice'), a Fresho / Rarity order ('THIS IS NOT AN INVOICE'), a PFD Sales Order and a Peters invoice (INV DATE 220726)."""
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas


def _c(path):
    c = canvas.Canvas(path, pagesize=A4)
    return c


def _rows(c, y, cols, heads, rows, size=8):
    c.setFont("Helvetica-Bold", size)
    for x, t in zip(cols, heads):
        c.drawString(x, y, t)
    c.setFont("Helvetica", size)
    for n, r in enumerate(rows):
        for x, t in zip(cols, r):
            c.drawString(x, y - 14 * (n + 1), t)


def mobigo_tax_invoice(path):
    c = _c(path)
    c.setFont("Helvetica-Bold", 13)
    c.drawString(40, 800, "MobiGo Distributors")
    c.drawString(250, 795, "Tax Invoice")
    c.setFont("Helvetica", 9)
    for i, t in enumerate(["ABN: 87169618660", "MobiGo Distributors", "4/20 Alta Road", "Caboolture, QLD 4510"]):
        c.drawString(40, 782 - 12 * i, t)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(390, 700, "Order #")
    c.drawString(480, 700, "Date")
    c.setFont("Helvetica", 9)
    c.drawString(390, 688, "S506325")
    c.drawString(470, 688, "23/07/2026")
    c.setFont("Helvetica-Bold", 9)
    c.drawString(40, 650, "Bill To:")
    c.drawString(300, 650, "Ship To:")
    c.setFont("Helvetica", 9)
    for i, t in enumerate(["PEARL ENERGY WOODRIDGE", "101 COMPTON ROAD", "WOODRIDGE, QLD 4114"]):
        c.drawString(40, 636 - 12 * i, t)
        c.drawString(300, 636 - 12 * i, t)
    _rows(c, 590, (40, 130, 240, 370, 450), ("Sales Rep", "Payment Terms", "Customer PO", "Carrier", "Date Scheduled"),
          [("PawanS", "NET 30", "PEWODRIGE180726", "Delivery", "18/07/2026")])
    _rows(c, 540, (40, 70, 240, 390, 430, 480, 530), ("Item #", "Number", "Description", "Tax Code", "Unit Price", "Qty", "Total Price"),
          [("1", "RLLTU5", "NULON COOLANT RED", "GST", "$89.99", "1 CT3", "$ 89.99"),
           ("2", "LLTU5", "NULON ANTI-FREEZE", "GST", "$63.31", "1 CT3", "$ 63.31")])
    c.drawString(400, 470, "Subtotal")
    c.drawString(500, 470, "$139.36")
    c.drawString(400, 458, "GST")
    c.drawString(500, 458, "$13.94")
    c.drawString(400, 446, "Total")
    c.drawString(500, 446, "$153.30")
    c.save()


def mobigo_order(path):
    c = _c(path)
    c.setFont("Helvetica-Bold", 14)
    c.drawString(60, 790, "MobiGo")
    c.setFont("Helvetica", 9)
    for i, t in enumerate(["MobiGo Distributors", "Unit 4,", "20 Alta Road", "Caboolture QLD QLD 4510", "Australia", "ABN: 87169618660"]):
        c.drawString(420, 800 - 12 * i, t)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(40, 700, "Invoicing Address")
    c.setFont("Helvetica", 9)
    for i, t in enumerate(["BP QUIRINDI, Pearl Energy No 3 Pty Ltd", "4", "Duke Street", "Quirindi NSW 2343", "Australia"]):
        c.drawString(40, 688 - 12 * i, t)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(40, 610, "Shipping Address")
    c.setFont("Helvetica", 9)
    for i, t in enumerate(["BP QUIRINDI", "4", "Duke Street", "Quirindi NSW 2343", "Australia"]):
        c.drawString(40, 598 - 12 * i, t)
    c.setFont("Helvetica-Bold", 16)
    c.drawString(40, 520, "Order # S00372")
    c.setFont("Helvetica-Bold", 9)
    c.drawString(40, 500, "Order Date")
    c.drawString(250, 500, "Salesperson")
    c.setFont("Helvetica", 9)
    c.drawString(40, 488, "22/10/2026")
    c.drawString(250, 488, "Administrator")
    _rows(c, 450, (40, 60, 250, 330, 390, 450, 510), ("#", "DESCRIPTION", "BARCODE", "QUANTITY", "UNIT PRICE", "TAXES", "AMOUNT"),
          [("1", "JB WELD COLD WELD", "43425612656", "3.00 Units", "12.59", "10% GST", "$ 37.77"),
           ("2", "ARMORALL GLASS CLEANER", "70612440020", "6.00 Units", "11.49", "10% GST", "$ 68.94")])
    c.save()


def fresho_order(path):
    c = _c(path)
    c.setFont("Helvetica-Bold", 14)
    c.drawString(90, 800, "Delivery: Fri 24 Jul")
    c.drawString(90, 784, "Order: F56490272")
    c.setFont("Helvetica", 9)
    c.drawString(90, 770, "Received: 9:46 AM Thu Jul 23, 2026")
    c.setFont("Helvetica-Bold", 14)
    c.drawString(330, 795, "RARITY WHOLESALE")
    c.setFont("Helvetica", 8)
    c.drawString(340, 782, "QUALITY FRESH & PROCESSED FRUITS & VEGETABLES")
    c.setFont("Helvetica-Bold", 13)
    c.drawString(50, 730, "Pearl Energy Anna Bay")
    c.setFont("Helvetica-Bold", 8)
    c.drawString(300, 738, "HIREN PATEL +61 (0) 2 7229 3952")
    c.drawString(300, 726, "4136 Nelson Bay Road, Anna Bay")
    _rows(c, 680, (50, 380, 420, 470, 540), ("Product", "Qty", "Unit", "Unit Price", "Price"),
          [("Carrots Grated", "1", "Kg", "$5.95", "$5.95"), ("Cucumber - Continental", "2", "Each", "$3.50", "$7.00")])
    c.drawString(450, 620, "Subtotal (ex GST)")
    c.drawString(540, 620, "$12.95")
    c.drawString(450, 608, "GST Charged")
    c.drawString(540, 608, "$0.00")
    c.setFont("Helvetica", 10)
    c.drawString(250, 580, "INDICATIVE TOTAL")
    c.drawString(520, 580, "$12.95")
    c.setFont("Helvetica-Bold", 11)
    c.drawString(220, 555, "THIS IS NOT AN INVOICE")
    c.setFont("Helvetica", 8)
    c.drawString(60, 542, "Final products, prices and quantities may vary due to seasons, weights, box sizes and market prices")
    c.save()


def pfd_sales_order(path):
    c = _c(path)
    c.setFont("Helvetica-Bold", 16)
    c.drawString(60, 800, "PFD")
    c.drawString(380, 800, "Sales Order")
    c.setFont("Helvetica", 9)
    c.drawString(500, 790, "Page 1")
    for i, t in enumerate(["A.B.N. 29 006 972 381", "PFD Food Servs P/L-Brisbane", "Brisbane DC 2", "38 Freight Street",
                           "LYTTON QLD 4178"]):
        c.drawString(40, 760 - 12 * i, t)
    c.drawString(320, 770, "Order Date  23 Jul 26")
    c.drawString(460, 770, "Due Date  24 Jul 26")
    c.drawString(320, 756, "Sales Order  VU959200")
    c.drawString(320, 742, "Ship Via  10276")
    c.drawString(40, 690, "SOLD TO: 07268554")
    c.drawString(320, 690, "SHIP TO: 07268554")
    for i, t in enumerate(["PEARL ENERGY DEAGON", "11 DEPORT ROAD", "DEAGON QLD 4017"]):
        c.drawString(40, 676 - 12 * i, t)
        c.drawString(320, 676 - 12 * i, t)
    _rows(c, 610, (40, 90, 300, 360, 420, 470, 520), ("Code", "Description", "Qty Ordered", "Price", "Net Value", "G.S.T.", "Value (Incl. G.S.T.)"),
          [("013486", "125GX24 SNACK SAUSAGE ROLLS", "1.00", "55.69", "55.69", "5.57", "61.26")])
    c.save()


def peters_invoice(path):
    c = _c(path)
    c.setFont("Helvetica-Bold", 16)
    c.drawString(60, 800, "Peters")
    c.drawString(260, 800, "TAX INVOICE")
    c.setFont("Helvetica-Bold", 7)
    cols = (140, 200, 262, 322, 380, 465, 520)
    for x, t in zip(cols, ("CUSTOMER NO", "ORDER DATE", "INVOICE NO", "ORDER NO", "YOUR ORDER NO", "INV DATE", "DUE DATE")):
        c.drawString(x, 770, t)
    c.setFont("Helvetica", 8)
    for x, t in zip(cols, ("5020443", "220726", "11620355", "2801909", "", "220726", "310826")):
        if t:
            c.drawString(x, 756, t)
    c.setFont("Helvetica", 8)
    for i, t in enumerate(["Australasian Food Group Pty Ltd", "Trading as Peters Ice Cream", "ABN 11154314913", "Private Bag 22",
                           "Mulgrave North VIC 3170"]):
        c.drawString(40, 740 - 11 * i, t)
    c.setFont("Helvetica-Bold", 8)
    c.drawString(270, 720, "INVOICE TO:")
    c.drawString(420, 720, "DELIVER TO:")
    c.setFont("Helvetica", 8)
    for i, (a, b) in enumerate((("TTF Pearl Energy No 2 Fut", "Mobil Ballina"), ("485 River Street", "485 River Street"),
                                ("BALLINA,2478,NSW", "BALLINA,2478,NSW"))):
        c.drawString(270, 708 - 11 * i, a)
        c.drawString(420, 708 - 11 * i, b)
    _rows(c, 640, (40, 100, 300, 360, 420, 480, 530), ("ITEM NO", "ITEM DESCRIPTION", "QTY SHIPPED", "UNIT PRICE", "EXTENDED", "GST", "NET AMOUNT"),
          [("109728", "DRUMSTICK Super Choc LP24", "1.00", "60.56", "79.68", "*", "60.56")])
    c.drawString(40, 560, "TOTAL")
    c.drawString(530, 560, "300.73")
    c.drawString(430, 540, "GST")
    c.drawString(530, 540, "30.07")
    c.drawString(430, 528, "TOTAL DUE")
    c.drawString(530, 528, "330.80")
    c.save()
