"""Runs every automated test (reader, rules, end-to-end workflow with a simulated Outlook).

    python selftest.py          quick summary
    python selftest.py -v       every test by name

It uses a temporary folder and a simulated Outlook: it needs no Outlook and changes nothing on this PC."""
import os
import sys
import unittest

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)


def main():
    verbosity = 2 if "-v" in sys.argv else 1
    suite = unittest.defaultTestLoader.discover(os.path.join(ROOT, "tests"), top_level_dir=ROOT)
    result = unittest.TextTestRunner(verbosity=verbosity).run(suite)
    print("\n" + ("ALL TESTS PASSED (%d tests)" % result.testsRun if result.wasSuccessful() else "SOME TESTS FAILED"))
    return 0 if result.wasSuccessful() else 1


if __name__ == "__main__":
    sys.exit(main())
