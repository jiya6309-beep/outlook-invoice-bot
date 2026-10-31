"""Puts a database backup back.      python restore_backup.py            (lists the backups)
                                   python restore_backup.py 1          (restores number 1, after asking)
                                   python restore_backup.py 1 --yes    (no question)

Close the bot first (the control panel may stay open but no run may be in progress). The current database is NOT deleted:
it is kept next to it as  data\\bot.db.before_restore_<time>  so the restore can be undone."""
import argparse
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def main(argv=None):
    ap = argparse.ArgumentParser(description="Restore a database backup")
    ap.add_argument("number", nargs="?", type=int, help="the number in the list (1 = newest)")
    ap.add_argument("--yes", action="store_true", help="do not ask for confirmation")
    args = ap.parse_args(argv)
    from invoicebot import db, paths, settings

    folder = settings.run_value("Backup folder") or paths.p("data", "backups")
    backups = db.list_backups(folder)
    if not backups:
        print(f"There are no backups in {folder}.")
        return 1
    print(f"Backups in {folder}:")
    for i, (path, d, size) in enumerate(backups, start=1):
        print(f"  {i:2}. {d}   {size / 1048576:7.1f} MB   {os.path.basename(path)}")
    if args.number is None:
        print("\nTo restore one:   python restore_backup.py <number>")
        return 0
    if not 1 <= args.number <= len(backups):
        print("That number is not in the list.")
        return 2
    lock = paths.p("data", "run.lock")
    if os.path.exists(lock) and time.time() - os.path.getmtime(lock) < 3600:
        print("A run seems to be in progress (data\\run.lock). Wait for it to finish, or press Stop in the panel, then try again.")
        return 3
    path, d, _size = backups[args.number - 1]
    if not args.yes and input(f"Restore the backup of {d}? The current database is kept as a .before_restore file. Type yes: ").strip().lower() != "yes":
        print("Nothing was changed.")
        return 0
    try:
        kept = db.restore(path, paths.p("data", "bot.db"))
    except Exception as e:
        print(f"Not restored: {e}")
        return 4
    print(f"Restored the backup of {d}." + (f"\nThe previous database was kept as: {kept}" if kept else ""))
    return 0


if __name__ == "__main__":
    sys.exit(main())
