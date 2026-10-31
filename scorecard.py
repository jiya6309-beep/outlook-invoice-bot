"""Grades the reader on a folder of real invoices.    python scorecard.py FOLDER [--out Scorecard.xlsx]

Every PDF / Excel file in the folder is read (the file name is used like an e-mail subject). Saved .msg e-mails work too:
their PDF/Excel attachments are extracted and the real subject, sender and recipient are used. Nothing is filed or
changed. The table shows what the bot concluded; 'problems' are documents the bot would send to _Unrecognised.
Use it on 50-100 real invoices to see how well the bot reads YOUR suppliers, then fix the Labels / Suppliers sheets."""
import argparse
import collections
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


DOC_EXT = (".pdf", ".xlsx", ".xlsm")


def collect(folder, tmp):
    """[(label, path, ctx, email rule)] from PDFs / Excel files and from .msg e-mails (attachments are extracted)."""
    from invoicebot import doctypes, msgfile
    items = []
    for name in sorted(os.listdir(folder)):
        full = os.path.join(folder, name)
        low = name.lower()
        if low.endswith(DOC_EXT):
            items.append((name, full, {"subject": os.path.splitext(name)[0].replace("_", " "), "orig_name": name}, ""))
        elif low.endswith(".msg"):
            try:
                m = msgfile.load(full)
            except (msgfile.MsgError, OSError) as e:
                print(f"SKIP {name}: {e}")
                continue
            word = doctypes.blocked_word(m["subject"])
            rule = f"BLOCKED ('{word}'): the bot would not download this" if word else (
                "keyword match" if doctypes.subject_matches(m["subject"]) else "no keyword: the bot would ignore this e-mail")
            for i, a in enumerate(m["attachments"], start=1):
                if a["method"] != 1 or a["hidden"] or not a["data"] or not a["filename"].lower().endswith(DOC_EXT):
                    continue
                path = os.path.join(tmp, f"{len(items):04d}_{i}_{a['filename']}")
                with open(path, "wb") as f:
                    f.write(a["data"])
                items.append((f"{name} > {a['filename']}", path, {
                    "subject": m["subject"], "sender_name": m["sender_name"], "sender_email": m["sender_email"],
                    "to_name": m["to"], "orig_name": a["filename"]}, rule))
            if not any(it[0].startswith(name + " >") for it in items):
                print(f"NOTE {name}: no PDF/Excel attachment ({rule})")
    return items


def main(argv=None):
    ap = argparse.ArgumentParser(description="Reader scorecard")
    ap.add_argument("folder", help="a folder of PDF / Excel files and/or saved .msg e-mails")
    ap.add_argument("--out", default="Scorecard.xlsx")
    args = ap.parse_args(argv)
    args.folder = args.folder.strip('"').strip()
    import tempfile
    import time
    from openpyxl import Workbook
    from invoicebot import extractor, ocr, settings

    settings.reset()
    ocr.configure(settings.run_bool("OCR enabled", True), settings.run_value("Tesseract path"), settings.run_int("OCR max pages", 3))
    tmp = tempfile.mkdtemp(prefix="scorecard_")
    items = collect(args.folder, tmp)
    if not items:
        print("No PDF / Excel / .msg files with documents in", args.folder)
        return 1
    wb = Workbook()
    ws = wb.active
    ws.title = "Scorecard"
    ws.append(["File", "E-mail rule", "Type", "Supplier", "How supplier found", "Number", "Date", "Amount", "GST",
               "Company on document", "Site", "Pages", "Seconds", "Problems", "Warnings"])
    stats, by_type, by_src = collections.Counter(), collections.Counter(), collections.Counter()
    for label, path, ctx, rule in items:
        t = time.time()
        r = extractor.analyse(path, ctx)
        secs = round(time.time() - t, 1)
        ok = not r["problems"]
        stats["read without problems" if ok else "problems"] += 1
        by_type[r["doc_type"] or "(unknown)"] += 1
        by_src[r["supplier_source"] or "(none)"] += 1
        print(f"{'OK  ' if ok else 'PROB'} {label[:50]:50} {str(r['doc_type'])[:12]:12} {r['supplier'][:24]:24} #{str(r['invoice_number'])[:12]:12} "
              f"{r['invoice_date']} {r['amount']} {('| ' + r['problems'][0][:50]) if r['problems'] else ''}")
        if rule.startswith("BLOCKED"):
            print(f"       ^ {rule}")
        ws.append([label, rule, r["doc_type"], r["supplier"], r["supplier_source"], r["invoice_number"], r["invoice_date"], r["amount"],
                   r["gst"], r["doc_company"], r["site"], r["page_count"], secs, "; ".join(r["problems"]), "; ".join(r["warnings"])])
    wb.save(args.out)
    n = len(items)
    print(f"\n{n} documents: {stats['read without problems']} read without problems ({100 * stats['read without problems'] // n}%), {stats['problems']} with problems")
    print("By type:", dict(by_type), "\nSupplier found by:", dict(by_src), "\nSaved:", args.out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
