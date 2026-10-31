"""Explains how ONE file was read.    python diagnose.py FILE.pdf [--subject "Invoice 123"] [--sender "Name <a@b.com>"]

Shows the first lines of text, the supplier decision, the document type decision, the labels used, the values,
the problems, and the company / site decision. Use it when a document lands in _Unrecognised and you want to know why."""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def main(argv=None):
    ap = argparse.ArgumentParser(description="Explain how a file is read")
    ap.add_argument("file", help="a PDF / Excel file, or a saved .msg e-mail (its first PDF/Excel attachment is used)")
    ap.add_argument("--attachment", type=int, default=1, help="with a .msg: which document attachment (1 = the first)")
    ap.add_argument("--subject", default="")
    ap.add_argument("--sender", default="")
    ap.add_argument("--mailbox", default="")
    args = ap.parse_args(argv)
    from invoicebot import companies, doctypes, extractor, msgfile, ocr, pdftext, settings

    settings.reset()
    ocr.configure(settings.run_bool("OCR enabled", True), settings.run_value("Tesseract path"), settings.run_int("OCR max pages", 3))
    if args.file.lower().endswith(".msg"):                       # use the e-mail's own subject, sender and recipient
        import tempfile
        m = msgfile.load(args.file)
        docs = [a for a in m["attachments"] if a["method"] == 1 and not a["hidden"] and a["data"]
                and a["filename"].lower().endswith((".pdf", ".xlsx", ".xlsm"))]
        print(f"\nE-MAIL   : {m['subject']}\nFROM     : {m['sender_name']} <{m['sender_email']}>   TO: {m['to']}   ARRIVED: {m['received']}")
        word = doctypes.blocked_word(m["subject"])
        print("RULE     :", f"BLOCKED by '{word}': the bot would NOT download it (red flag)" if word else
              ("keyword match: the bot would download it" if doctypes.subject_matches(m["subject"]) else "no keyword: the bot would ignore it"))
        if not docs:
            print("No PDF / Excel attachment in this e-mail.")
            return 0
        a = docs[min(max(args.attachment, 1), len(docs)) - 1]
        args.file = os.path.join(tempfile.mkdtemp(prefix="diag_"), a["filename"])
        with open(args.file, "wb") as f:
            f.write(a["data"])
        args.subject = args.subject or m["subject"]
        args.sender = args.sender or f"{m['sender_name']} <{m['sender_email']}>"
        args.mailbox = args.mailbox or m["to"]
    name, _, addr = args.sender.partition("<")
    ctx = {"subject": args.subject, "sender_name": name.strip(), "sender_email": addr.rstrip(">").strip(), "mailbox": args.mailbox,
           "orig_name": os.path.basename(args.file), "keep_text": True}
    r = extractor.analyse(args.file, ctx)
    print(f"\nFILE: {args.file}   pages: {r['page_count']}   OCR used: {r['ocr_used']}")
    print("\nFIRST LINES OF PAGE 1")
    try:
        if args.file.lower().endswith(".pdf"):
            text = pdftext.load(args.file, max_pages=1).text
        else:
            text = r.get("_text", "")
    except Exception as e:
        text = f"(could not read: {e})"
    for line in [l for l in text.splitlines() if l.strip()][:25]:
        print("   |", line[:110])
    print("\nSUPPLIER :", r["supplier"] or "(not found)", f"  [found by: {r['supplier_source'] or '-'}; printed as: {r['supplier_printed'] or '-'}]")
    sender = doctypes.sender_row(ctx["sender_name"], ctx["sender_email"])
    print("SENDER   :", (sender["Sender contains"] + (" (always download)" if doctypes.always_download(sender) else " (hint only)")) if sender else "(no Senders row matches)")
    print("TYPE     :", r["doc_type"] or "(not recognised)", f"  [{r['type_source'] or '-'}]")
    print("LABELS   :", {k: v[:4] for k, v in settings.labels_for(r["supplier_printed"], r["supplier"]).items()})
    print("NUMBER   :", r["invoice_number"], "\nDATE     :", r["invoice_date"], "\nTOTAL    :", r["amount"], "\nGST      :", r["gst"], "\nMETHOD   :", r["method"])
    print("COMPANY  :", r["doc_company"] or "(none named on the document)", "| mailbox company:", (companies.by_name(args.mailbox) or companies.for_mailbox(args.mailbox) or {}).get("Folder name", "-"))
    print("SITE     :", r["site"], f"[{r['site_source']}]")
    print("\nPROBLEMS :", r["problems"] or "none  (this document would be filed)")
    print("WARNINGS :", r["warnings"] or "none")
    return 0


if __name__ == "__main__":
    sys.exit(main())
