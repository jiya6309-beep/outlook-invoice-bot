"""Optional: tell each site handler where new files were saved.  Run option  'Email path mode':
off (default) | draft (creates a draft in Outlook, nothing is sent) | send."""
from collections import defaultdict

from . import sites


def path_emails(gateway, filed, mode, log=print):
    """filed = list of dicts (site, supplier, number, dest_dir, dest_name). Returns the number of e-mails created."""
    if mode not in ("draft", "send") or not filed:
        return 0
    groups = defaultdict(list)
    for d in filed:
        row = sites.site_row(d.get("site", ""))
        if row and row["email"]:
            groups[row["email"]].append((row["name"], d))
    created = 0
    for email, items in groups.items():
        lines = [f"- {site}: {d['dest_name']}   ({d['supplier']} {d.get('number') or ''})\n    {d['dest_dir']}" for site, d in items]
        subject = f"Invoice bot: {len(items)} new file(s) saved"
        body = "Hello,\n\nThe following files were saved by the invoice bot:\n\n" + "\n".join(lines) + "\n\nRegards,\nInvoice Bot"
        try:
            gateway.create_mail(email, subject, body, send=(mode == "send"))
            created += 1
        except Exception as e:
            log(f"could not create the e-mail for {email}: {e}")
    return created
