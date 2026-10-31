"""Builds demo_data/ (sample documents + emails.json) used by  python run_bot.py --demo"""
import json
import os
import shutil
import sys

from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import make_data  # noqa: E402

ROOT = os.path.dirname(HERE)
DEMO = os.path.join(ROOT, "demo_data")


def main():
    make_data.build_all()
    os.makedirs(DEMO, exist_ok=True)
    for f in os.listdir(DEMO):
        os.remove(os.path.join(DEMO, f))
    for name in ("invoice_acme.pdf", "invoice_bidfood.pdf", "invoice_other.pdf", "invoice_company_named.pdf", "credit_note.pdf",
                 "pod.pdf", "statement.pdf", "account_summary.pdf", "scanned_invoice.pdf", "corrupt.pdf", "invoice_excel.xlsx",
                 "invoice_logo_only.pdf", "invoice_bad_math.pdf"):
        shutil.copy(os.path.join(make_data.DATA, name), os.path.join(DEMO, name))
    shutil.copy(os.path.join(make_data.DATA, "statement_rarity.pdf"), os.path.join(DEMO, "statement_rarity.pdf"))
    rnd = Image.effect_noise((600, 500), 90).convert("RGB")
    rnd.save(os.path.join(DEMO, "photo_of_invoice.jpg"), quality=95)

    one, colyton = "One@pearlenergy.com.au", "pearl.colyton@pearlenergy.com.au"
    att = lambda f, **k: dict(file=f, **k)
    emails = [
        dict(mailbox=one, hours_ago=3, subject="Tax Invoice 778899 Bidfood", sender="Bidfood <accounts@bidfood.example>",
             to="One Pearl Energy", attachments=[att("invoice_bidfood.pdf")], expect="downloaded -> ONE PEARL ENERGY folder, read + tick"),
        dict(mailbox=one, hours_ago=4, subject="Credit Note CN-9001", attachments=[att("credit_note.pdf")],
             expect="credit note: amounts stored NEGATIVE, file ends _CN"),
        dict(mailbox=one, hours_ago=5, subject="POD D-77120", attachments=[att("pod.pdf")], expect="proof of delivery: file ends _POD, no amounts"),
        dict(mailbox=one, hours_ago=6, subject="Statement for September", attachments=[att("statement.pdf")],
             expect="statement: kept in the _Statements holding folder"),
        dict(mailbox=one, hours_ago=7, subject="Account Summary for 0001458043 as of 02/10/2026", attachments=[att("account_summary.pdf")],
             expect="parked in _Unrecognised\\Account Summary"),
        dict(mailbox=one, hours_ago=8, subject="Overdue invoice reminder", attachments=[att("invoice_other.pdf")],
             expect="blocked word: NOT downloaded, red flag, stays unread"),
        dict(mailbox=one, hours_ago=9, subject="07254094 - PEARL ENERGY NO 1 PTY LTD **Stop Supply Warning**", attachments=[],
             expect="blocked word, no attachment: red flag, stays unread"),
        dict(mailbox=one, hours_ago=10, subject="FW: Invoice NS-4040", attachments=[att("invoice_company_named.pdf")],
             expect="forwarded: red category; the document names RADHE KRISHNA, so it is filed in RADHE KRISHNA INVESTMENT"),
        dict(mailbox=one, hours_ago=11, subject="Meeting tomorrow at 10", attachments=[], expect="not a document e-mail: left completely untouched"),
        dict(mailbox=one, hours_ago=12, subject="Invoice - photo attached", attachments=[att("photo_of_invoice.jpg")],
             expect="image only: orange 'Image attachment' category, unread, nothing downloaded"),
        dict(mailbox=one, hours_ago=13, subject="inv", attachments=[att("corrupt.pdf")], expect="unreadable file: goes to _Unrecognised"),
        dict(mailbox=one, hours_ago=14, subject="Invoice SC-7001 (scanned)", attachments=[att("scanned_invoice.pdf")],
             expect="scanned PDF: read by OCR if Tesseract is installed, else _Unrecognised"),
        dict(mailbox=one, hours_ago=15, subject="Invoice XL-5001", attachments=[att("invoice_excel.xlsx")], expect="Excel invoice is read like a PDF"),
        dict(mailbox=one, hours_ago=16, subject="From Rarity Wholesale", sender="Accounts <accounts@raritywholesale.com.au>",
             attachments=[att("statement_rarity.pdf")], expect="exception sender: downloaded although the subject has no keyword"),
        dict(mailbox=one, hours_ago=17, subject="Invoice RW-2020", attachments=[att("invoice_logo_only.pdf")],
             expect="supplier name only in the logo: found by OCR (needs Tesseract)"),
        dict(mailbox=one, hours_ago=18, subject="Invoice 100 - totals do not add up", attachments=[att("invoice_bad_math.pdf")],
             expect="100 + 10 is not 150: refused, goes to _Unrecognised with the reason"),
        dict(mailbox=colyton, hours_ago=2, subject="FW: Invoice INV-1001", attachments=[att("invoice_acme.pdf")],
             expect="filed in COLYTON FUELS (mailbox/company), red flag (FW)"),
        dict(mailbox=colyton, hours_ago=1, subject="Invoice INV-1001 again", attachments=[att("invoice_acme.pdf")],
             expect="the same invoice again: recognised as a duplicate, not saved twice"),
    ]
    with open(os.path.join(DEMO, "emails.json"), "w", encoding="utf-8") as f:
        json.dump(emails, f, indent=1)
    print("demo data:", len(os.listdir(DEMO)), "files,", len(emails), "e-mails")


if __name__ == "__main__":
    main()
