# Go-live checklist

Work through this in order. Everything is a button in the control panel (`Start_Bot.bat`). Each step is reversible until step E.

## A. Prepare the bot PC
- [ ] The Windows user that runs Outlook is the user that will run the bot. Classic Outlook installed, **open**, all 13 mailboxes visible.
- [ ] Python 3.10+ installed ("Add to PATH" ticked). `Start_Bot.bat` opened the control panel (it installed the packages the first time).
- [ ] *(Recommended)* Tesseract OCR installed. **Check my PC** says "Tesseract OCR ready".
- [ ] OneDrive is signed in and the SharePoint library is synced to this PC. The bot's account can **edit** the folders.
- [ ] The PC does not sleep while the bot should run (Windows power plan).
- [ ] The synced SharePoint path plus `\MONTH YEAR EMAIL INVOICES\COMPANY\file` stays under 260 characters (or Windows long paths are enabled).

## B. Connect and check the rules
- [ ] **Setup 1 → Find my Outlook mailboxes**: every company has its mailbox (or the account is saved as a TEST mailbox). **Check my PC** reports "all company mailboxes found".
- [ ] `Bot_Settings.xlsx` → **Companies**: *Folder name* is spelled exactly like the SharePoint company folder (for example `TWEED HEADS ENERGY NO.1`).
- [ ] **Run → Month folder pattern** produces your real month folder name (default → `OCTOBER 2026 EMAIL INVOICES`).
- [ ] **Keywords / BlockedWords / Senders** reviewed against the written requirements.
- [ ] **Suppliers → Folder override**: `MILK INVOICES` set for the milk suppliers (e.g. Lactalis), if that is how they are filed.
- [ ] **Sites**: handler e-mails filled in if *Email path mode* will be used.

## C. Rehearse (nothing is changed)
- [ ] **Try the demo** works and shows the sample results.
- [ ] **Rehearsal** on the real mailboxes for **several days**. Open `Reports\Shadow_Report_*.xlsx`, sheet *Would do*: are the folders, file names, flags and categories what a person would have done?
- [ ] Sheet *New suppliers*: names look right (add aliases in **Suppliers** if not). Sheet *Skipped with attachments*: nothing important is being ignored.
- [ ] `python scorecard.py <folder of 50–100 real invoices or saved .msg e-mails>`: the large majority "read without problems"; tune **Labels** for the rest
      (`python diagnose.py file.pdf --subject "…"` explains each one).
- [ ] The run summary's "Time" and "Peak memory" are comfortable for the PC.

## D. Test mode with real e-mails
- [ ] **Run now** in *TEST* mode: files go to `TestOutput\`; e-mails are ticked / flagged for real. Prefer one mailbox first (the "Only one mailbox?" box).
- [ ] Spot-check 20 documents against the original e-mails (number, date, amount, GST, company folder).
- [ ] Practise the To-Fix loop once: open the To-Fix list, fix one document, save and close, **Run now**, watch it move.
- [ ] Open this month's report (Open → This month's report): the Record sheet lists the documents you expect, and Needs attention lists only what really needs a person.

## E. Switch to the real destination
- [ ] OneDrive has synced the library on this PC. **Setup 2**: leave the folder box empty (the bot finds `…\OFFICE INDIA\EMAIL INVOICES` itself), keep the **Folder layout** on *Month, company, site*, press **Switch to LIVE**. The message says where it found the folder — check it.
- [ ] **Setup 2 → Match my sites to the SharePoint folders**. Read the report: sites without a folder, folders without a site (add an alias or a site row), similar spellings (check them). Run it again until you are satisfied.
- [ ] **Test Outlook marking on ONE e-mail** (panel, Setup 1): read, tick, red flag and category all say OK. If the tick says NO, read README section 5a.
- [ ] **Check my PC**: the root exists and "site folders found" is close to the number of sites.
- [ ] Decide: staff keep creating site folders (default) **or** set *Create site folder in live mode* = Yes (then staff stop creating them).
- [ ] First live run for one day only (**Run for dates**, today to today, untick *Rehearsal only*). Verify the files appear in SharePoint (sync icon green).

## F. Run unattended
- [ ] **Setup 3 → Run automatically** every 30 minutes (runs while the user is signed in; keep Outlook open).
- [ ] Someone looks at the **Needs your attention** box and the 🚩 red-flagged e-mails each day.
- [ ] The last 3 days of each month: the run summary warns if next month's folder does not exist yet.
- [ ] Point *Run → Backup folder* at the NAS (backups are compressed, about 1 GB after a year) so the database backup is off the PC. Practise **once** with `Advanced\Restore_backup.bat` that a backup can be listed.
- [ ] Decide where the quarterly archives go (`Records archive folder`; default `Reports\Archive`). Copy the first one somewhere safe when it appears.

## If something looks wrong
1. Press **Stop** (and turn **Run automatically** off) — nothing is lost, the e-mails are still in Outlook.
2. **Open → Last run log**.
3. Switch back to TEST, or use **Rehearsal**, while you investigate.
