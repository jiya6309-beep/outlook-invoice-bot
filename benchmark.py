"""Stress test:   python benchmark.py [N] [--workers W]      (default N = 500; the real load is about 2500 a day)

Builds N distinct invoices (needs reportlab: pip install reportlab), spreads them over the 13 mailboxes of a
simulated Outlook, runs the real pipeline, and reports speed and memory. Uses a temporary folder."""
import argparse
import os
import random
import shutil
import sys
import tempfile
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def main(argv=None):
    ap = argparse.ArgumentParser(description="Benchmark")
    ap.add_argument("n", nargs="?", type=int, default=500)
    ap.add_argument("--workers", type=int, default=None)
    args = ap.parse_args(argv)
    try:
        sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "tests"))
        import make_data
    except ImportError:
        print("reportlab / Pillow are needed for the benchmark:  pip install reportlab pillow")
        return 2
    home = tempfile.mkdtemp(prefix="botBench_")
    os.environ["BOT_HOME"] = home
    import datetime
    from invoicebot import companies, settings
    from invoicebot.fake_outlook import FakeMail, FakeOutlook
    from invoicebot.pipeline import Pipeline

    settings.reset()
    settings.get()
    docs = os.path.join(home, "_docs")
    os.makedirs(docs)
    rnd = random.Random(7)
    suppliers = ["Acme Foods Pty Ltd", "Bidfood Australia Ltd", "Orbit Wholesale Pty Ltd", "Nova Snacks Pty Ltd", "Delta Drinks Pty Ltd", "Metro Fuels Pty Ltd"]
    gw = FakeOutlook()
    boxes = [companies.mailbox_names(r)[0] for r in companies.all_active()]
    for b in boxes:
        gw.add_mailbox(b)
    now = datetime.datetime.now().replace(microsecond=0)
    t = time.time()
    print(f"building {args.n} documents ...")
    for i in range(args.n):
        kind = rnd.random()
        path = os.path.join(docs, f"doc_{i}.pdf")
        number = f"B{100000 + i}"
        if kind < 0.82:
            make_data.invoice(path, rnd.choice(suppliers), number, net=round(rnd.uniform(20, 900), 2))
            subject = f"Tax Invoice {number}"
        elif kind < 0.90:
            make_data.invoice(path, rnd.choice(suppliers), number, net=round(rnd.uniform(5, 100), 2), title="CREDIT NOTE", number_label="Credit Note No")
            subject = f"Credit Note {number}"
        elif kind < 0.95:
            make_data.simple(path, "WEEKLY STATEMENT", [f"Statement Date: 12/09/2026", f"Ref {number}"], rnd.choice(suppliers))
            subject = f"Statement {number}"
        else:
            make_data.simple(path, "PROOF OF DELIVERY", [f"Delivery No: {number}", "Date: 02/10/2026"], rnd.choice(suppliers))
            subject = f"POD {number}"
        gw.add(boxes[i % len(boxes)], FakeMail(subject, now - datetime.timedelta(minutes=1 + i % 600),
                                               attachments=[{"name": f"{number}.pdf", "path": path}]))
    print(f"built in {time.time() - t:.0f} s. Running the bot ...\n")
    run = Pipeline(gw, workers=args.workers, log=lambda *a, **k: None, now=now + datetime.timedelta(minutes=5))
    t = time.time()
    s = run.run()
    secs = time.time() - t
    print(s.text())
    per_min = s.c["documents"] / (secs / 60) if secs else 0
    print(f"\nSPEED: {s.c['documents']} documents in {secs:.0f} s = {per_min:.0f} documents per minute "
          f"({secs / max(s.c['documents'], 1):.2f} s per document) with {run.n_workers} reader process(es)")
    print(f"A day of 2500 documents would take about {2500 / max(per_min, 1):.0f} minutes on this PC.")
    shutil.rmtree(home, ignore_errors=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
