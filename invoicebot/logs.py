"""Run log: printed to the screen and written to logs/run_<time>.txt (plain ASCII, safe on any Windows console)."""
import datetime
import os
import sys
import time

from . import paths


class Logger:
    def __init__(self, to_file=True, quiet=False):
        self.quiet = quiet
        self.path = None
        self._fh = None
        if to_file:
            folder = paths.ensure("logs")
            self._clean_old(folder)
            self.path = os.path.join(folder, f"run_{datetime.datetime.now():%Y%m%d_%H%M%S}.txt")
            self._fh = open(self.path, "w", encoding="utf-8", errors="replace")

    @staticmethod
    def _clean_old(folder, days=30):
        cutoff = time.time() - days * 86400
        for name in os.listdir(folder):
            path = os.path.join(folder, name)
            try:
                if name.startswith("run_") and name.endswith(".txt") and os.path.getmtime(path) < cutoff:
                    os.remove(path)
            except OSError:
                pass

    def __call__(self, text="", stamp=True):
        line = f"{datetime.datetime.now():%H:%M:%S}  {text}" if (stamp and text) else str(text)
        if not self.quiet:
            try:
                print(line)
            except UnicodeEncodeError:
                print(line.encode(sys.stdout.encoding or "ascii", "replace").decode(sys.stdout.encoding or "ascii"))
        if self._fh:
            self._fh.write(line + "\n")
            self._fh.flush()

    def close(self):
        if self._fh:
            self._fh.close()
            self._fh = None
