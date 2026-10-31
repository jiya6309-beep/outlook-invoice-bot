"""Outlook Invoice Bot - control panel launcher.   Double-click  Start_Bot.bat  (it runs this file).

Opens the control panel in your web browser: big buttons, no commands to type. Keep this window open while you use the
panel; close it (or press Ctrl+C) to quit."""
import argparse
import os
import sys
import webbrowser

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def main(argv=None):
    ap = argparse.ArgumentParser(description="Outlook Invoice Bot control panel")
    ap.add_argument("--port", type=int, default=8765, help="port on this PC (default 8765; another is chosen if it is busy)")
    ap.add_argument("--no-browser", action="store_true", help="do not open the browser automatically")
    args = ap.parse_args(argv)

    from invoicebot import __version__, panel
    controller = panel.Controller()
    try:
        server, token = panel.make_server(controller, port=args.port)
    except OSError:
        server, token = panel.make_server(controller, port=0)
    url = f"http://127.0.0.1:{server.server_address[1]}/?t={token}"
    print(f"Outlook Invoice Bot {__version__} - control panel")
    print("=" * 60)
    print("The control panel is open in your browser.")
    print("If it did not open, copy this address into your browser:\n")
    print("   " + url + "\n")
    print("Keep this window open while you work. Close it to quit.")
    if not args.no_browser:
        webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
