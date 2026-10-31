"""5.14.0: the supplier is never taken from a brand in the item table; a printed 'T/A' trading name wins; the sender's own e-mail
domain is a strong clue (not for forwards); companies can be switched off in Setup step 2 (Active) with a warning in every run."""
import os
import tempfile
import unittest

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

from tests.helpers import ONE, Isolated, connect_db, gateway, mail, run

from invoicebot import companies, mailboxes, settings, suppliers


def rum_city_credit_note(path, abn="38 509 899 256"):
    """The Rum City Foods credit note layout: 'Setlex Pty Ltd T/A' / 'Rum City Foods', and SUNTORY as the BRAND of the items."""
    c = canvas.Canvas(path, pagesize=A4)
    y = 800
    for text, size, bold in (("Setlex Pty Ltd T/A", 8, False), ("Rum City Foods", 10, True), (f"ABN {abn}", 7, False),
                             ("15 Commercial Street", 7, False), ("Bundaberg, QLD, 4670", 7, False)):
        c.setFont("Helvetica-Bold" if bold else "Helvetica", size)
        c.drawString(40, y, text)
        y -= 13
    c.setFont("Helvetica-Bold", 16)
    c.drawString(380, 800, "CREDIT NOTE")
    c.setFont("Helvetica", 8)
    for i, (k, v) in enumerate((("Date:", "14/10/2026"), ("Customer No.", "6119"), ("Credit Note No:", "6026709"), ("Invoice No:", "42507749"))):
        c.drawString(380, 780 - 12 * i, k)
        c.drawString(500, 780 - 12 * i, v)
    c.drawString(40, 690, "Sold To:")
    c.drawString(250, 690, "Deliver To:")
    c.drawString(40, 678, "PEARL ENERGY GIN GIN FOODSERVICE")
    c.drawString(250, 678, "PEARL ENERGY GIN GIN")
    c.drawString(40, 650, "PROD CLASS   BRAND     STOCK CODE   DESCRIPTION                 QTY   PRICE   GST   EXTENSION")
    c.drawString(40, 636, "DRY AMB      SUNTORY   FRU1094      12x500ml V BLUE CAN         5.00  11.40   -5.70  -62.74")
    c.drawString(40, 624, "DRY AMB      SUNTORY   FRU1099      24x500ml V GREEN CAN        3.00  22.79   -6.84  -75.21")
    c.drawString(380, 200, "Total excluding GST   -$125.41")
    c.drawString(380, 188, "GST   -$12.54")
    c.drawString(380, 176, "Total including GST   -$137.95")
    c.showPage()
    c.save()


class SupplierTests(Isolated, unittest.TestCase):
    def setUp(self):
        super().setUp()
        self.tmp = tempfile.mkdtemp(prefix="v514_")

    def pdf(self, name, **kw):
        p = os.path.join(self.tmp, name)
        rum_city_credit_note(p, **kw)
        return p

    def supplier_of(self, path, **mail_kw):
        g = gateway()
        g.add(ONE, mail("Please find Credit Note attached 6026709", [path], **mail_kw))
        run(g)
        return dict(connect_db(self.home).execute("SELECT supplier, doc_type, number FROM docs").fetchone())

    def test_rum_city_is_not_read_as_suntory(self):
        d = self.supplier_of(self.pdf("SACRERPT-55359370.pdf"))
        self.assertEqual(d["supplier"], "Rum City Foods")
        self.assertNotIn("Suntory", d["supplier"])

    def test_the_trading_name_wins_even_without_the_supplier_list(self):
        d = self.supplier_of(self.pdf("cn.pdf", abn="11 222 333 444"))       # ABN not in the supplier list
        self.assertEqual(d["supplier"], "Rum City Foods")

    def test_a_brand_in_the_item_table_never_decides_the_supplier(self):
        text = ("Unknown Wholesale\n15 Commercial Street\nCREDIT NOTE\nSold To: Deliver To:\nPEARL ENERGY GIN GIN\n"
                "PROD CLASS BRAND STOCK CODE DESCRIPTION\nDRY AMB SUNTORY FRU1094 12x500ml V BLUE CAN")
        self.assertEqual(suppliers.identify(text)[0], "Unknown Wholesale")
        self.assertEqual(suppliers.identify("Suntory Beverage & Food Australia Pty Ltd\nTAX INVOICE\nDeliver To: X")[0],
                         "Suntory Beverage & Food Australia Pty Ltd")    # a real Suntory invoice is still Suntory

    def test_the_senders_own_domain_is_a_clue_but_not_for_forwards(self):
        text = "Kierena Penney\nRum City Foods\nStatement\nDeliver To: X\nBRAND SUNTORY"
        self.assertEqual(suppliers.from_sender_domain("kierena@rumcityfoods.com.au", text), "Rum City Foods")
        self.assertIsNone(suppliers.from_sender_domain("Pearl.colyton@pearlenergy.com.au", text))
        self.assertIsNone(suppliers.from_sender_domain("someone@gmail.com", text))
        name, how = suppliers.identify(text.replace("Rum City Foods\n", ""), sender_email="x@suntory.com")
        self.assertEqual((name, how), ("Suntory Oceania", "sender e-mail domain"))          # a domain in the supplier list
        name, how = suppliers.identify(text.replace("Rum City Foods\n", ""), sender_email="x@suntory.com", is_forward=True)
        self.assertNotEqual(how, "sender e-mail domain")


class ActiveTests(Isolated, unittest.TestCase):
    def test_a_company_switched_off_in_step_2_is_not_read_and_the_run_says_so(self):
        two = companies.by_name("PEARL ENERGY NO: 2 PTY LTD")
        one = companies.by_name("PEARL ENERGY NO: 1 PTY LTD")
        mailboxes.apply([{"outlook": companies.mailbox_names(one)[0], "action": "company", "company": "PEARL ENERGY NO: 1 PTY LTD",
                          "active": False}])
        settings.reset()
        names = [r["Company name"] for r in companies.reading_list()]
        self.assertNotIn("PEARL ENERGY NO: 1 PTY LTD", names)
        self.assertIn("PEARL ENERGY NO: 2 PTY LTD", names)
        g = gateway()
        e = g.add(ONE, mail("Tax Invoice 778899", ["invoice_bidfood.pdf"]))
        s = run(g)
        self.assertEqual(s.c["documents"], 0)
        self.assertTrue(e.unread)
        self.assertTrue(any("switched OFF" in w and "PEARL ENERGY NO: 1" in w for w in s.warnings))
        from invoicebot.panel import Controller
        c = Controller()
        self.assertEqual(c.switched_off(), ["PEARL ENERGY NO: 1 PTY LTD"])
        self.assertTrue(c.activate_all()["ok"])
        settings.reset()
        self.assertEqual(c.switched_off(), [])
        self.assertTrue(two)


if __name__ == "__main__":
    unittest.main()
