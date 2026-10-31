"""Compressed backups and restore, and the quarterly read-only records archive."""
import datetime
import gzip
import os
import stat
import subprocess
import sys
import unittest

from openpyxl import load_workbook

from tests.helpers import ONE, ROOT, Isolated, connect_db, gateway, mail, run, write

from invoicebot import db, panel, quarterly

D = datetime.date


def doc(conn, number, received, status="Filed", doc_type="Invoice", **kw):
    row = dict(status=status, supplier="Acme Foods Pty Ltd", doc_type=doc_type, number=number, doc_date=received[:10], amount=110.0, gst=10.0,
               company_folder="ONE PEARL ENERGY", site="Mobil Colyton", mailbox="M", received=received + " 10:00:00", dest_dir="C:\\x",
               dest_name=f"{number}.pdf", current_path=f"C:\\x\\{number}.pdf", file_hash=f"h{number}", reason="")
    row.update(kw)
    return db.insert_doc(conn, row)


class BackupTests(Isolated, unittest.TestCase):
    def setUp(self):
        super().setUp()
        self.folder = self.path("bk")
        self.conn = db.connect(self.path("data", "bot.db"))
        for i in range(300):
            doc(self.conn, f"INV-{i}", "2026-10-01")
        self.conn.commit()

    def test_a_backup_is_compressed_checked_and_once_a_day(self):
        self.conn.execute("PRAGMA wal_checkpoint(TRUNCATE)")                                           # SQLite keeps recent data in a side file until now
        b = db.backup(self.conn, self.folder)
        self.assertTrue(b.endswith(".db.gz"))
        self.assertLess(os.path.getsize(b), os.path.getsize(self.path("data", "bot.db")) / 2)       # clearly smaller
        self.assertIsNone(db.backup(self.conn, self.folder))
        self.assertEqual(sorted(os.listdir(self.folder)), [os.path.basename(b)])                      # no temporary files left behind
        with gzip.open(b, "rb") as g:
            self.assertTrue(g.read(16).startswith(b"SQLite format 3"))

    def test_old_backups_of_both_kinds_are_removed(self):
        os.makedirs(self.folder)
        write(os.path.join(self.folder, "bot_20200101.db"), "old plain")
        write(os.path.join(self.folder, "bot_20200102.db.gz"), "old compressed")
        write(os.path.join(self.folder, "notes.txt"), "mine")
        db.backup(self.conn, self.folder, keep_days=14)
        names = os.listdir(self.folder)
        self.assertNotIn("bot_20200101.db", names)
        self.assertNotIn("bot_20200102.db.gz", names)
        self.assertIn("notes.txt", names)                                                              # only backups are ever deleted

    def test_restore_brings_everything_back_and_keeps_the_old_database(self):
        b = db.backup(self.conn, self.folder)
        self.conn.execute("DELETE FROM docs")
        self.conn.commit()
        self.conn.close()
        kept = db.restore(b, self.path("data", "bot.db"))
        self.assertTrue(os.path.exists(kept))                                                         # the undo copy
        conn = connect_db(self.home)
        self.assertEqual(conn.execute("SELECT COUNT(*) FROM docs").fetchone()[0], 300)
        old = db.connect(kept)
        self.assertEqual(old.execute("SELECT COUNT(*) FROM docs").fetchone()[0], 0)
        old.close()

    def test_a_damaged_backup_is_refused_and_nothing_changes(self):
        os.makedirs(self.folder)
        bad = os.path.join(self.folder, "bot_20260101.db.gz")
        with gzip.open(bad, "wb") as g:
            g.write(b"this is not a database" * 100)
        self.conn.close()
        with self.assertRaises(Exception):
            db.restore(bad, self.path("data", "bot.db"))
        self.assertEqual(connect_db(self.home).execute("SELECT COUNT(*) FROM docs").fetchone()[0], 300)
        self.assertEqual([f for f in os.listdir(self.path("data")) if "restoring" in f], [])

    def test_the_restore_command(self):
        db.backup(self.conn, self.path("data", "backups"))
        self.conn.execute("DELETE FROM docs")
        self.conn.commit()
        self.conn.close()
        script = os.path.join(ROOT, "restore_backup.py")
        listing = subprocess.run([sys.executable, script], capture_output=True, text=True, cwd=ROOT)
        self.assertIn("1.", listing.stdout)
        self.assertIn(".db.gz", listing.stdout)
        write(self.path("data", "run.lock"), "pid=1")
        busy = subprocess.run([sys.executable, script, "1", "--yes"], capture_output=True, text=True, cwd=ROOT)
        self.assertEqual(busy.returncode, 3)                                                           # a run is in progress: refused
        os.remove(self.path("data", "run.lock"))
        done = subprocess.run([sys.executable, script, "1", "--yes"], capture_output=True, text=True, cwd=ROOT)
        self.assertEqual(done.returncode, 0, done.stdout + done.stderr)
        self.assertEqual(connect_db(self.home).execute("SELECT COUNT(*) FROM docs").fetchone()[0], 300)

    def test_a_live_run_makes_a_compressed_backup(self):
        gw = gateway()
        gw.add(ONE, mail("Tax Invoice 778899", ["invoice_bidfood.pdf"]))
        run(gw)
        self.assertEqual([f for f in os.listdir(self.path("data", "backups"))], [f"bot_{datetime.date.today():%Y%m%d}.db.gz"])


class QuarterTests(unittest.TestCase):
    def test_calendar_quarters(self):
        self.assertEqual(quarterly.bounds(2026, 3), (D(2026, 7, 1), D(2026, 9, 30)))
        self.assertEqual(quarterly.bounds(2026, 4), (D(2026, 10, 1), D(2026, 12, 31)))
        self.assertEqual(quarterly.bounds(2026, 1), (D(2026, 1, 1), D(2026, 3, 31)))
        self.assertEqual(quarterly.bounds(2024, 1), (D(2024, 1, 1), D(2024, 3, 31)))
        self.assertEqual(quarterly.last_finished(D(2026, 10, 5)), (2026, 3))
        self.assertEqual(quarterly.last_finished(D(2026, 1, 1)), (2025, 4))
        self.assertEqual(quarterly.last_finished(D(2026, 9, 30)), (2026, 2))                          # the quarter is not over yet
        self.assertEqual(quarterly.parse("2026-Q3"), (2026, 3))
        self.assertEqual(quarterly.parse(" 2026 q1 "), (2026, 1))
        with self.assertRaises(ValueError):
            quarterly.parse("2026-Q5")


class ArchiveTests(Isolated, unittest.TestCase):
    def setUp(self):
        super().setUp()
        self.conn = db.connect(self.path("data", "bot.db"))
        for i in range(5):
            doc(self.conn, f"Q3-{i}", f"2026-08-{i + 1:02d}")
        doc(self.conn, "Q3-OPEN", "2026-09-30", status="Unassigned", reason="site could not be determined")
        doc(self.conn, "Q3-STM", "2026-07-15", status="Parked", doc_type="Statement")
        doc(self.conn, "Q2-OLD", "2026-06-30")                                                          # before the quarter
        doc(self.conn, "Q4-NEW", "2026-10-02")                                                          # after it
        db.add_correction(self.conn, "R000001", "site", "UNASSIGNED", "Mobil Colyton", "jiya")
        self.conn.execute("UPDATE corrections SET at='2026-08-20 09:00:00'")
        self.conn.commit()

    def test_the_archive_holds_the_quarter_and_nothing_is_deleted(self):
        path, n = quarterly.make(self.conn, 2026, 3)
        self.assertEqual(n, 7)
        self.assertTrue(path.endswith(os.path.join("Reports", "Archive", "Archive_2026-07_to_2026-09.xlsx")))
        wb = load_workbook(path)
        self.assertEqual(wb.sheetnames, ["Record", "Needs attention", "Corrections", "About"])
        numbers = sorted(r[3] for r in wb["Record"].iter_rows(min_row=2, values_only=True))
        self.assertEqual(numbers, ["Q3-0", "Q3-1", "Q3-2", "Q3-3", "Q3-4", "Q3-OPEN", "Q3-STM"])      # open and parked items are in it too
        self.assertEqual([r[0] for r in wb["Needs attention"].iter_rows(min_row=2, values_only=True) if r[0]], ["R000006"])
        self.assertEqual([tuple(r[:5]) for r in wb["Corrections"].iter_rows(min_row=2, values_only=True)], [("R000001", "site", "UNASSIGNED", "Mobil Colyton", "jiya")])
        self.assertIn("nothing was deleted", " ".join(str(r[0]) for r in wb["About"].iter_rows(values_only=True)))
        wb.close()
        self.assertEqual(self.conn.execute("SELECT COUNT(*) FROM docs").fetchone()[0], 9)               # the database still holds everything

    def test_the_file_is_read_only_and_not_written_twice(self):
        path, _n = quarterly.make(self.conn, 2026, 3)
        self.assertEqual(os.stat(path).st_mode & stat.S_IWUSR, 0)
        self.assertEqual(quarterly.make(self.conn, 2026, 3), (None, 0))                                # already there
        again, n = quarterly.make(self.conn, 2026, 3, force=True)                                      # asked for by name: written again
        self.assertEqual((again, n), (path, 7))

    def test_due_only_after_the_quarter_ends_and_only_with_documents(self):
        q2, n2 = quarterly.make_due(self.conn, D(2026, 9, 30))                                         # Q3 is not over: the last finished quarter is Q2
        self.assertEqual((os.path.basename(q2), n2), ("Archive_2026-04_to_2026-06.xlsx", 1))
        made, n = quarterly.make_due(self.conn, D(2026, 10, 1))                                        # the day after Q3: Q3 is due
        self.assertEqual((os.path.basename(made), n), ("Archive_2026-07_to_2026-09.xlsx", 7))
        self.assertIsNone(quarterly.make_due(self.conn, D(2027, 4, 2))[0])                             # Q1 2027 has no documents: no empty file
        self.assertEqual(sorted(os.listdir(os.path.dirname(made))), ["Archive_2026-04_to_2026-06.xlsx", "Archive_2026-07_to_2026-09.xlsx"])

    def test_the_archive_folder_can_be_moved(self):
        self.set_run(Records_archive_folder=self.path("NAS", "archives"))
        path, _n = quarterly.make(self.conn, 2026, 3)
        self.assertEqual(os.path.dirname(path), self.path("NAS", "archives"))

    def test_a_live_run_writes_it_once_when_due(self):
        self.conn.close()
        gw = gateway()
        s1 = run(gw, now=datetime.datetime(2026, 10, 5, 14, 0))
        self.assertTrue(any("Quarterly records archive written" in n and "7 documents" in n for n in s1.notes), s1.notes)
        archive = self.path("Reports", "Archive", "Archive_2026-07_to_2026-09.xlsx")
        self.assertTrue(os.path.exists(archive))
        stamp = os.path.getmtime(archive)
        s2 = run(gw, now=datetime.datetime(2026, 10, 5, 14, 30))
        self.assertFalse(any("Quarterly" in n for n in s2.notes))
        self.assertEqual(os.path.getmtime(archive), stamp)                                             # not touched again

    def test_the_switch_turns_it_off(self):
        self.conn.close()
        self.set_run(Quarterly_records_archive="No")
        run(gateway(), now=datetime.datetime(2026, 10, 5, 14, 0))
        self.assertFalse(os.path.exists(self.path("Reports", "Archive")))

    def test_a_failing_archive_never_stops_the_run(self):
        self.conn.close()
        self.set_run(Records_archive_folder=self.path("blocked"))
        write(self.path("blocked"), "a FILE where the folder should be")
        gw = gateway()
        now = datetime.datetime(2026, 10, 5, 14, 0)
        gw.add(ONE, mail("Tax Invoice 778899", ["invoice_bidfood.pdf"], hours_ago=2, now=now))
        s = run(gw, now=now)
        self.assertEqual(s.c["Filed"], 1)
        self.assertTrue(any("quarterly records archive failed" in w for w in s.warnings))

    def test_the_command(self):
        self.conn.close()
        env = dict(os.environ)
        out = subprocess.run([sys.executable, os.path.join(ROOT, "run_bot.py"), "--archive-quarter", "2026-Q3"], capture_output=True, text=True, cwd=ROOT, env=env)
        self.assertEqual(out.returncode, 0, out.stdout + out.stderr)
        self.assertIn("7 documents written", out.stdout)
        bad = subprocess.run([sys.executable, os.path.join(ROOT, "run_bot.py"), "--archive-quarter", "2026-Q9"], capture_output=True, text=True, cwd=ROOT, env=env)
        self.assertNotEqual(bad.returncode, 0)

    def test_the_panel_can_open_the_archive_folder(self):
        opened = []
        c = panel.Controller(opener=opened.append)
        self.assertTrue(c.open_path("records_archive")["ok"])
        self.assertEqual(opened, [self.path("Reports", "Archive")])


if __name__ == "__main__":
    unittest.main()
