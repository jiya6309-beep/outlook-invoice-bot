"""Where everything lives.

By default the bot works inside the folder it was unzipped to, so it can sit anywhere
(C:\\OutlookInvoiceBot, D:\\Bots, ...). Set the environment variable BOT_HOME to move it."""
import os


def project_dir():
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def home():
    return os.path.abspath(os.environ.get("BOT_HOME") or project_dir())


def p(*parts):
    return os.path.join(home(), *parts)


def ensure(*parts):
    """Create (if needed) and return a folder under home."""
    path = p(*parts)
    os.makedirs(path, exist_ok=True)
    return path
