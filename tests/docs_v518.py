"""Rebuilt layouts for 5.18.0: a Metcash 'Cover Sheet - POD' (customer box '102-104 WOLLOMBI ROAD', boxes beside it), a Lactalis
TAX INVOICE with 'Delivery Note ...' and a Metcash Customer Activity Report whose PDF is titled Statement."""
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas


def metcash_pod(path):
    c = canvas.Canvas(path, pagesize=landscape(A4))
    c.setFont("Helvetica-Bold", 14)
    c.drawString(40, 560, "Metcash Food")
    c.drawString(220, 565, "METCASH FOOD & GROCERY PTY LTD")
    c.drawString(220, 545, "Cover Sheet - POD")
    c.setFont("Helvetica", 9)
    c.drawString(520, 565, "****Invoices E-Mailed****")
    c.drawString(680, 565, "POD NO")
    c.drawString(740, 565, "G2141383")
    c.setFont("Helvetica", 6)
    c.drawString(40, 520, "HUNTINGWOOD - GROCERY  A.B.N. 67 004 391 422")
    c.drawString(300, 520, "71 HUNTINGWOOD DRIVE EASTERN CREEK NSW 2766")
    c.setFont("Courier", 9)
    c.drawString(40, 495, "CUST#  086048   00-1960-1704")
    c.drawString(40, 480, "FGRC METRO CESSNOCK")
    c.drawString(40, 470, "102-104 WOLLOMBI ROAD")
    c.drawString(40, 460, "CESSNOCK, NSW 2325")
    c.drawString(40, 440, "TH 12-15  * C-STORE *")
    c.setFont("Helvetica-Bold", 7)
    for y, t in ((530, "CARRIER"), (505, "CHEP A/C NO"), (480, "LOSCAM A/C"), (465, "SEAL NUMBERS")):
        c.drawString(450, y, t)
    c.setFont("Courier", 9)
    c.drawString(560, 530, "HWK1")
    c.setFont("Helvetica", 8)
    c.drawString(40, 120, "INVOICE COUNT   2")
    c.drawString(160, 120, "INVOICE DATE  29/07/2026")
    c.drawString(40, 100, "INVOICES   68654   68655")
    c.drawString(40, 70, "GOODS RECEIVED IN GOOD ORDER AND CONDITION")
    c.save()


def lactalis_invoice(path):
    c = canvas.Canvas(path, pagesize=A4)
    c.setFont("Helvetica-Bold", 13)
    c.drawString(120, 800, "Lactalis Australia Pty Ltd")
    c.setFont("Helvetica-Bold", 18)
    c.drawString(430, 800, "TAX INVOICE")
    c.setFont("Helvetica", 8)
    for i, t in enumerate(["ABN 56 072 928 879", "P.O. Box 3012", "South Brisbane QLD 4101"]):
        c.drawString(120, 786 - 10 * i, t)
    for i, (a, b) in enumerate((("Date Ordered:", "06.10.2026"), ("Invoice Date:", "11.10.2026"))):
        c.drawString(300, 760 - 12 * i, a)
        c.drawString(370, 760 - 12 * i, b)
    for i, (a, b) in enumerate((("Document No:", "244596854"), ("Customer No:", "9014689"), ("Account No:", "9014463"))):
        c.drawString(450, 760 - 12 * i, a)
        c.drawString(510, 760 - 12 * i, b)
    c.setFont("Helvetica-Bold", 8)
    c.drawString(280, 700, "Ship To:")
    c.setFont("Helvetica", 8)
    for i, t in enumerate(["AMPOL FARLEIGH", "TTF PEARL ENERGY NO 2 FUT", "971 BRUCE HIGHWAY", "FARLEIGH 4741"]):
        c.drawString(280, 688 - 10 * i, t)
    c.drawString(450, 690, "TOTAL AMOUNT: $ 731.31")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(40, 620, "Delivery Note 433201 for 6.10.2026 Purchase Order 2359665")
    c.setFont("Helvetica", 8)
    c.drawString(40, 600, "PAULS FULL CREAM MILK BTL 2L        00001003     6 EA     22.50")
    c.drawString(400, 560, "Sub Total")
    c.drawString(500, 560, "664.83")
    c.drawString(400, 548, "GST")
    c.drawString(500, 548, "66.48")
    c.drawString(400, 536, "Total")
    c.drawString(500, 536, "731.31")
    c.save()


def metcash_activity_report(path):
    c = canvas.Canvas(path, pagesize=A4)
    c.setFont("Helvetica-Bold", 14)
    c.drawString(40, 800, "Metcash Food & Grocery Pty Ltd")
    c.drawString(400, 800, "Statement")
    c.setFont("Helvetica", 9)
    c.drawString(40, 784, "ABN 67 004 391 422")
    c.drawString(40, 760, "MOBIL TORRENSVILLE")
    c.drawString(40, 748, "239 HENLEY BEACH ROAD")
    c.drawString(40, 736, "TORRENSVILLE SA 5031")
    c.drawString(380, 760, "Account: 85852440")
    c.drawString(380, 748, "Statement Date: 19/10/2026")
    c.drawString(40, 700, "Date        Reference      Description           Amount")
    c.drawString(40, 686, "13/10/2026  68001          Invoice               512.40")
    c.drawString(40, 674, "15/10/2026  68044          Invoice               388.10")
    c.save()
