# Outlook Invoice Bot

Collects supplier invoices, credit notes and delivery dockets from the group's Outlook mailboxes, reads them, names them consistently,
files them into the right **month → company** folder, keeps a searchable record, and shows what needs a person.
Built for the Pearl Energy group: 13 mailboxes, about 2,500 documents a day.

> **Version 5.20.0** · Windows + classic desktop Outlook · Python 3.10 or newer
> Inside the bot and every technology used: [`docs/HOW_IT_WORKS.md`](docs/HOW_IT_WORKS.md) · Before going live: [`docs/GO_LIVE_CHECKLIST.md`](docs/GO_LIVE_CHECKLIST.md)

---

## 1. Start in two steps

1. **Install Python once** from python.org. On the first screen tick **"Add python.exe to PATH"**. (Skip if it is already installed.)
2. **Double-click `Start_Bot.bat`.**
   The first time it installs the parts it needs by itself (about a minute, needs internet once). Then your browser opens the **control panel**.

That is all. Everything else is a button.

```
 ┌──────────────────────────────────────────────────────────────────────────────┐
 │  📨 Outlook Invoice Bot v5.1                                   [ TEST MODE ] │
 │                                                                              │
 │  What would you like to do?                                                  │
 │  [ ▶ Run now ] [ 🧪 Rehearsal ] [ 📅 Run for dates ] [ ✔ Check my PC ]       │
 │                                                                              │
 │  Activity .......... live log and result cards (Filed, Parked, Waiting ...)  │
 │  Needs your attention .......... [ Open the To-Fix list ] [ This month's report ] │
 │  Open .......... filed documents · holding folders · reports · logs · settings│
 │  Setup  1 Connect the Outlook mailboxes  2 Where should documents go?  3 Run automatically │
 └──────────────────────────────────────────────────────────────────────────────┘
```

### The first day, in order
| # | In the panel | What happens |
|---|---|---|
| 1 | **Check my PC** | Tells you whether everything is ready. Fix anything marked `FAIL`. |
| 2 | **Try the demo** (small link under the buttons) | Runs the whole bot on sample e-mails in a *simulated* Outlook. Your real mailboxes are not touched. |
| 3 | **Setup 1 → Find my Outlook mailboxes** | The bot lists the mailboxes open in Outlook and suggests which company each belongs to. Choose, press **Save these**. A personal or training mailbox can be saved as **📥 Read every e-mail** (files like the company mailboxes, company from each invoice) or as a **TEST mailbox** (own TEST folder). **🚫 Never read** = the bot ignores that mailbox completely. |
| 4 | **Rehearsal** | Reads the real mailboxes but **changes nothing**. Open the report it writes (`Reports\Shadow_Report_*.xlsx`) and check it. Repeat for a few days. |
| 5 | **Run now** | Files documents to the *test* folder `TestOutput\` and marks the e-mails (read ✓, red flags, categories). |
| 6 | **Setup 2** | The LIVE folder is already filled in (`OFFICE INDIA\EMAIL INVOICES`; it can be changed in `Bot_Settings.xlsx`, sheet Run, `Live destination root`). Press **Switch to LIVE**, then **Match my sites to the SharePoint folders** (section 3). |
| 7 | **Setup 3** | Turn on **Run automatically** every 30 minutes. |

*"New Outlook" and Outlook on the web cannot be automated: classic desktop Outlook must be installed, open and signed in.*

---

## 1a. Where everything goes (5.15.0)

`OI` = `C:\SHAREPOINT\pearlenergy.com.au\Patel Group - Documents\OFFICE INDIA`

| What | Where | Name |
|---|---|---|
| Fuel invoice (P I Fuel Invoice: Inbox + its sub-folders) | `OI\FUEL INVOICES\<COMPANY>\` | `<INVOICE NUMBER>_<SITE>.pdf`, listed in `Reports\PI_Invoice_list_<MON>_<YEAR>.xlsx` |
| Invoice / credit note / docket | `OI\EMAIL INVOICES\<MONTH> <YEAR> EMAIL INVOICES\<COMPANY>\<SITE>\` | `Supplier_Number(.pdf / _CN / _POD)` |
| Fleet Card statement (mstatement@) | `OI\TEMPORARY DOWNLOADS\FLEETCARD STATEMENTS\<SITE>\` | `<report issue date>_STM.pdf` |
| Security reports (security run only) | `OI\TEMPORARY DOWNLOADS\SECURITY REPORTS\<MONTH YEAR>\` | `Downloaded Reports\` + `SITE_DATA_<MON>_<YEAR>_<dd-mm>_TO_<dd-mm>.xlsx` |
| Supplier statement | `OI\EMAIL INVOICES\HOLDINGS\_Statements\<Month Year>\<COMPANY>\` | `Supplier_Site_Date_STM.pdf` |
| Order confirmation | `OI\EMAIL INVOICES\HOLDINGS\_Order Confirmations\<Month Year>\` | `Supplier_Site_Date_OC.pdf` |
| Direct Debit Notice | `OI\EMAIL INVOICES\HOLDINGS\_Direct Debit Notices\<COMPANY>\` | `Supplier_Date_DDN.pdf` |
| Account Activity Listing | `OI\EMAIL INVOICES\HOLDINGS\_Account Activity Listings\<SITE>\` | `Supplier_Date_AAL.pdf` |
| Site / site folder not found | `OI\EMAIL INVOICES\HOLDINGS\_Unassigned\` | waits for a fix |
| Could not be read | `OI\EMAIL INVOICES\HOLDINGS\_Unrecognised\<Month Year>\` | |
| To_Fix, Record, A/C Excel, quarterly archive | `OI\EMAIL INVOICES\Reports\` | |
| Database, logs, Scratch, 7-day originals | next to the program (local disk, never synced) | |

Folders: LIVE creates every missing folder exactly like TEST (month, company, site, `HOLDINGS`, `Reports` - an existing folder is
always used, never made twice). Only `EMAIL INVOICES` itself is never created: if it is missing, SharePoint is not synced on that PC
and the run stops before anything is written. The paths are in `Bot_Settings.xlsx`, sheet Run: `Live destination root` (blank
`Holding root` / `Reports root` = inside it). Fleet Card / M STATEMENTS site folders are created when missing.
Mailboxes: the 13 companies + P I Fuel Invoice (fuel invoices) + mstatement@ (Fleet Card). Security is read ONLY by the security
run (its own buttons, schedule and Security_Settings.xlsx); office@ and serviceinvoices@ are NEVER read. In TEST mode nothing goes
to SharePoint
(sheet SpecialMailboxes). Only the ship-to / deliver-to address decides the site when it is printed - never the bill-to. Without a ship-to, the **Sold to**
address decides (Bega), still never the bill-to / payer; an unknown sold-to waits in To_Fix. The amount is always the **amount payable**
(Total Payable, inc GST).

## 2. The control panel

| Button | What it does |
|---|---|
| **▶ Run now** | Reads the new e-mails of all mailboxes, files the documents, marks the e-mails. Asks for confirmation first. |
| **🧪 Rehearsal** | The same, but **changes nothing** (no Outlook change, no file written). Produces a report of what *would* happen. |
| **📅 Run for dates** | Reads e-mails that *arrived* between two dates (both days included). Can be a rehearsal. |
| **✔ Check my PC** | Python packages, OCR, settings, folders, database, Outlook connection, missing mailboxes. |
| **Only one mailbox?** box | Type part of a mailbox name to limit any run to it (handy for testing with one account). |
| **Only touch e-mails whose subject contains …** box | For testing inside a mailbox that also holds REAL e-mails: type a word such as `BOTTEST`; every other e-mail is left completely alone. |
| **Try the demo** | Sample e-mails in a simulated Outlook. |
| **Test with saved e-mails** | Runs the bot on a folder of `.msg` files (drag e-mails out of Outlook into a folder). Real Outlook untouched. |
| **■ Stop** | Stops a running run safely. Nothing is lost: the next run continues from where it reached. |
| **Needs your attention** | How many documents wait or could not be read, with a button to the To-Fix list. |
| **Open …** | Shortcuts to filed documents, holding folders, reports, logs, settings and this guide. |
| **Setup** | Mailboxes · test/live destination · automatic running. |

The panel only listens on your own PC (`127.0.0.1`), needs a secret in its address, and runs the bot as a separate process,
so a problem in the page can never damage a run. Close the black window (or press **Close panel**) to quit it; an automatic schedule keeps running.

---

## 3. ⭐ The destination folder — found automatically, nothing to type

In LIVE mode the bot **finds the synced `EMAIL INVOICES` folder on its own**: it looks under this PC's user folder / OneDrive for a folder called `EMAIL INVOICES` inside `OFFICE INDIA`
(for example `C:\Users\<user>\pearlenergy.com.au\Patel Group - Documents\OFFICE INDIA\EMAIL INVOICES`). It only lists folder names (never opens or downloads a file), remembers what it found
(`data\destination.txt`) and checks it again at every run. Both names are settings (`Auto-find folder name`, `Auto-find parent folder`), so a renamed folder is one cell.

**So: Setup 2 → leave the box empty → Switch to LIVE.** (Until then the bot runs in **TEST mode** and writes under `TestOutput\`.) Prerequisite: OneDrive is signed in and has synced that library on the bot PC.

A path is only needed to *force* a different folder; it can then be given in any of these ways (a typed path always wins over the automatic search):

| Way | How |
|---|---|
| **A. Control panel** | **Setup 2** → type the folder → **Switch to LIVE**. |
| **B. Excel** | `Bot_Settings.xlsx` → sheet **Run** → `Live destination root`, `Destination mode` = `live`. |
| **C. Code** | `config.py` → `PUT THE REAL SHAREPOINT PATH HERE (LATER)`: remove the `#`, edit the line. |

If the folder cannot be found the run **stops with a clear message** and changes nothing (switch the search off with `Find the destination folder automatically` = No).

### The folder layout (which sub-folders the bot uses)
Sheet **Run** → `Destination pattern` (or the **Folder layout** list in panel Setup 2). The version you received is set to **month folder, then company folder, then site folder** (the shape of your SharePoint):

| Layout | Result |
|---|---|
| `{SITE}` | `<root>\MOBIL BALLINA\Supplier_Number.pdf` |
| `{COMPANY}\{SITE}` | `<root>\COLYTON FUELS\MOBIL BALLINA\…` |
| `{SITE}\{MONTH} {YEAR}` | `<root>\MOBIL BALLINA\OCTOBER 2026\…` |
| `{MONTH_FOLDER}\{COMPANY}` | `<root>\OCTOBER 2026 EMAIL INVOICES\COLYTON FUELS\Supplier_Number.pdf` |
| **`{MONTH_FOLDER}\{COMPANY}\{SITE}`** *(shipped)* | `<root>\OCTOBER 2026 EMAIL INVOICES\ONE PEARL ENERGY\MOBIL BALLINA\Supplier_Number.pdf` |

Invoices, credit notes and PODs use the layout. Statements go to `Holding\_Statements` whatever the layout.

### How the bot decides: first the company, then the site, then the folder
1. **Company** — the company named on the document (*Invoice To: RADHE KRISHNA INVESTMENT PTY LTD*), else the company mailbox the e-mail was **sent or forwarded from**, else the mailbox it arrived in.
2. **Site, among that company's sites only** — the supplier **account number** → the **delivery address** → a **site name** printed on the document → the e-mail's recipient name → a similar spelling.
   A document addressed to **"PEARL ENERGY <TOWN>"** finds the site of that town if only one site of the company has that town. **If two sites share a town** (for example Wagga Wagga, where one site is literally called *Pearl Energy Wagga wagga*), a document that gives only "PEARL ENERGY WAGGA WAGGA" and no street address **waits** instead of being guessed; naming the brand site (*MOBIL WAGGA WAGGA*) or printing the address settles it. This is why two sites with the same town but different companies
   (for example the two Gretas) are never confused: once the company is known, only its sites are considered.
3. **Folder** — the folder of your layout: month, company, site (shipped). **If the site cannot be decided, the document is filed in the COMPANY folder (one level up) and the Remarks column says `SITE NOT FOUND`**; set `File in the company folder when the site is not found` = No to make it wait in `_Unassigned` instead.

Two safeguards: a company **named on the document** is trusted — if none of its sites fits, the file waits instead of being searched in other companies. A company that only comes from the sender or the mailbox
is a weaker clue — if none of its sites fits, the other companies' sites are tried, and the difference is written in the record. When the site is still uncertain the file waits in `Holding\_Unassigned`
(reason in `To_Fix.xlsx`) and a person picks the **Site folder** from the drop-down. It is never put in a guessed folder. Switch the two-step order off with `Choose the site within the company first` = No.

### Match the site list to your real folders (do this once, after switching to LIVE)
Folder names must be exactly right. Press **Setup 2 → 🔗 Match my sites to the SharePoint folders** (or `python run_bot.py --match-folders`). The bot reads the real folders, ignores capitals and extra spaces
(so `AMPOL  KELSO` with two spaces is kept exactly), accepts a similar spelling, writes each site's real name into the Sites sheet column **SharePoint folder**, and tells you:
* **sites without a folder** — documents for them wait in `_Unassigned`; create the folder or correct the name,
* **folders without a site** — add the site to the Sites sheet, or add the folder's old name as an *alias* of the right site,
* **similar spellings** it used — check them.
Re-run it any time folders change; names you typed yourself are never overwritten.

| Situation in LIVE mode | What the bot does |
|---|---|
| The site folder exists | files the document there |
| The site folder (or a month / company folder of your layout) does not exist | keeps the file in `Holding\_Unassigned`, **retries every run**, files it as soon as the folder exists |
| You want the bot to create folders | sheet **Run** → `Create site folder in live mode` (also `…month…` / `…company…`) = `Yes` — only if staff stop creating them, to avoid sync duplicates |
| A layout starting with the month folder | warns in the last 3 days of the month if next month's folder is missing |

Company folder names are in sheet **Companies** (column *Folder name*); the **MILK INVOICES** override per supplier (sheet **Suppliers**) applies to layouts that use `{COMPANY}`.

---

## 4. The Excel files it produces (`Reports\`)

| File | For | What is in it |
|---|---|---|
| **`Record_YYYY-MM.xlsx`** | management, accounts | **Record**: every document of the month (filters, colour-coded status, clickable *Open folder*, amounts and GST as numbers) · **Needs attention**: the documents waiting or unrecognised, with the reason |
| **`To_Fix.xlsx`** | the person who fixes things | **How to use** guide · **To fix** (yellow cells with drop-down lists, colour-coded status, *Open folder* links) · image-only e-mails · overdue/urgent e-mails · new supplier names |
| `Shadow_Report_<time>.xlsx` | after a Rehearsal | exactly what a real run would do, per e-mail and attachment |
| `Report_<from>_<to>.xlsx` | on request | the same management report for any period: **Run for dates → 📊 Make an Excel report for these dates** (reads no e-mail) |

The monthly report is rebuilt only when a run adds to that month (a full month of 55,000 documents takes about 10 seconds); a run with nothing new writes no Excel file at all. Do not type in it; it is a report.
If a file is open in Excel the new version is saved beside it as `…_new.xlsx` — nothing is lost.

**Fixing things — `To_Fix.xlsx`:** each row is a document that needs a person. Pick the **Company folder** from the drop-down (this is all that is needed for files waiting in `_Unassigned`),
or type the *Supplier / Number / Document type* for a document the bot could not read. Save, **close** the file. At the next run the bot moves the file,
logs who changed what and when, and the row disappears.

---

## 5. The rules the bot follows

### E-mails (subject-based, checked in this order)
| # | If the e-mail … | the bot … |
|---|---|---|
| 1 | was already handled | skips it (never twice) |
| 2 | has a **blocked word** in the subject: `overdue`, `urgent`, `reminder`, `remainder`, `stop supply`, `warning` | **does not download**; 🚩 red flag (only if it has no flag yet); stays unread; FW:/RE: is the same (red flag, no category) |
| 3 | has a **keyword** (`invoice`, `inv`, `tax invoice`, `credit note`, `cn`, `adj`, `statement`, `pod`, `docket`, `account summary`, …) or comes from an **exception sender** (Rarity Wholesale, Moorees, Tip Top, …) **and** a PDF/Excel attachment | downloads, reads, files; marks **read + ✓ complete** |
| 3a | is a **FW: / RE:** (any subject) with a PDF/Excel attachment (setting `Download attachments of FW/RE e-mails`) | downloads, reads, files (or `_Unrecognised`); the e-mail is left **UNREAD**, **no tick**, 🚩 red flag (no category) — a person looks at it |
| 4 | has a keyword and **only pictures** (≥ 40 KB, signature logos ignored) | orange category *Image attachment*; unread; nothing downloaded |
| 5 | has a keyword and **no attachment** | untouched (red flag if FW:/RE:) |
| 6 | anything else | untouched |

The bot **never** moves, deletes or replies to an e-mail, and never overwrites a flag you set yourself. It reads the **Inbox** only.
All word lists are editable in `Bot_Settings.xlsx` (sheets **Keywords**, **BlockedWords**, **Senders**).

### Documents
| Type | Decided by (in order) | Saved to | File name |
|---|---|---|---|
| Invoice | title in the PDF → subject → file name | `<root>\<SITE FOLDER>` (the folder layout, section 3) | `Supplier_Number.pdf` |
| Credit note | same | same (amounts stored **negative**) | `Supplier_Number_CN.pdf` |
| POD / delivery docket / delivery note | same, **and** what the page says (a page that calls itself a delivery docket and is not an invoice is a POD, even when the subject says invoice) | same (no amounts read, no `GST NOT FOUND`) | `Supplier_Number_POD.pdf` |
| Statement | same | **`Holding\_Statements\<Month Year>`** (its own folder, month folders inside) | `Supplier_Site_30-09-2026_STM.pdf` (`NoSite` when no site was found) |
| Order Confirmation | same | **`Holding\_Order Confirmations\<Month Year>`** | `Supplier_Site_30-09-2026_OC.pdf` (`NoSite` when no site was found) |
| Account Summary / Customer Activity Report / Hire Statement | same | `Holding\_Unrecognised\<same name>` (parked on purpose; change *Route* in the DocTypes sheet to file them like invoices) | `Supplier_Date_ACCSUM.pdf` … |
| Other (direct-debit notice, order receipt, follow-up report …) | same | `Holding\_Unrecognised` | `Supplier_Date_OTH.pdf` |
| Unreadable / unknown type / no supplier / no invoice number or date / over 15 pages / over 25 MB | — | `Holding\_Unrecognised` (reason in the record) | `Supplier_Number.pdf` if both were read, else `03-10-2026_originalname.pdf` |
| Known, but no destination yet (folder missing, first-time supplier on hold) | — | `Holding\_Unassigned` (retried every run) | final name |

**Month folders (5.8.0):** inside `_Statements`, `_Order Confirmations` and `_Unrecognised` the files sit in a folder named like `October 2026`. The document date decides, when it is within `Month window (months)` of the e-mail arrival month; otherwise the arrival month. A statement or order confirmation with no site found is named `...NoSite...`, filed anyway, and listed in To_Fix.xlsx (orange row, column *Statement missing data*; nothing to type, rename by hand if you want).

**One-time tidy of old files:** `python run_bot.py --tidy-holding` only LISTS what would move (and writes `Reports\Tidy_Holding_Preview.xlsx`). `python run_bot.py --tidy-holding --apply` moves and renames the old files into month folders (copy, size check, then remove the original). Only files the bot filed itself are touched.

**The holding folders** sit side by side in `Holding\` and are created at the first run: **`_Unrecognised`** (could not be read; plus the sub-folders for account summaries and reports), **`_Unassigned`** (known document, destination not available yet) and **`_Statements`**. The statements folder name is the `Subfolder` cell of the Statement row in sheet **DocTypes**; set `Route` to `Company folder` there if statements should be filed like invoices instead.

**Dates inside file names are written dd-mm-yyyy** (for example `Supplier_30-09-2026_STM.pdf`); the format is the `File name date format` cell in the Run sheet.

A name is never overwritten: identical file → skipped; different file, same name → `_2`. **Duplicates** (same content, or same supplier + type + number — for example a *Reprint*) are saved once: the second copy is not filed and the original row's Remarks says `duplicate arrived dd-mm-yyyy, not filed again`.
**A spreadsheet that arrives with an invoice** (for example "as per attached spreadsheet") is filed beside the invoice as `<invoice name>_breakdown.xlsx`.

### The supplier: how it is found (any one clue is enough)
1. the **Suppliers** sheet (a name in the top lines) → 2. sheet **SupplierIDs**: the **ABN**, the **e-mail domain** or the **account name** (for suppliers that print no name, e.g. JP Bakery; Bonfect) → 3. the printed issuer line: a line that ends in a **legal ending** (sheet **LegalEndings**: Pty, Pvt, Private, Limited, Ltd, …; `MCALLISTER & SONS PTY.` with `LIMITED` on the next line gives *McAllister & Sons Pty*) → 4. **Account Name:** in the bank box → 5. the sender (Senders sheet) → 6. the logo → 7. the sender's display name. Pty / Pvt / Private / Limited / Ltd count as the same ending when names are compared.

### Which month folder? (setting `Month folder follows`)
* Invoice dated **in the month the e-mail arrived** → that month's folder.
* Dated **1–3 months earlier** (`Month window (months)`) → its **own, older month folder**; the row in the arrival month's Record is **yellow**, Remarks says `INVOICE DATED AUGUST 2026: filed in the AUGUST 2026 folder`.
* Dated **more than 3 months back or ahead** → the current month's folder; **lilac**.
* **Future-dated within 3 months** → waits in the current month's folder (**orange**, `FUTURE-DATED NOVEMBER 2026 …`). In the last two days of the month the bot builds the next month's folders and **moves** it there; the old row stays orange with the note changed to `MOVED TO NOVEMBER 2026 on …`, and the November Record lists it again in **light blue** (`ARRIVED EARLY`) — it is counted only once (in the month it arrived), also in the quarterly archive. If the file was moved or renamed by hand the bot leaves it, adds `COULD NOT MOVE …` to Remarks and lists it in the run summary.
* A document **without an invoice date** still waits for a person.

### Month folders are built by the bot (setting `Create next month's folders`)
In the **last two days** of a month the bot copies the folder structure (company and site **folders only, never files**) of the latest month to the next month's folder; on the first run of a new month it builds a missing folder. Staff do not create month folders any more (two people creating the same folder makes OneDrive duplicates such as `Folder (1)`).

### Which company folder?
1. a company named on the document (e.g. *Invoice To: RADHE KRISHNA INVESTMENT PTY LTD*), 2. the company that owns the matched site, 3. **the company mailbox the e-mail was sent or forwarded from** (e.g. *Pearl Energy Colyton*, *P & O Investment*, *Shri Hari Investment*), 4. the mailbox the e-mail arrived in.
The sender clue is what makes a shared mailbox work: when one mailbox receives the invoices of every company, each file still goes to its own company's folder. Editable: sheet **Run** → `Company source order` and `Use sender as company clue`. If the sources disagree the bot uses the first and writes a warning in the record. Folder names are used exactly as written in the **Companies** sheet (so `P & O INVESTMENT` stays `P & O INVESTMENT`).

### Accuracy safeguards
Number, date and the three amounts are read next to their labels (also in table layouts, and in PDFs that deliver their letters one by one). The **invoice number, date, supplier and document type are required**; a future date is no longer rejected (see *Which month folder?*); the e-mail subject is an independent check on the invoice number. A GST figure that cannot belong to the total (larger than a tenth of it) is left empty instead of recorded.
**The three amounts are optional** (see the next section). **If a required item fails the document goes to `_Unrecognised` with the reason — never a guess.**

### The three amounts: total incl. GST, GST, amount excl. GST
* They are **never a reason to hold a document**: an invoice or credit note is filed in its folder even when one, two or all three cannot be found.
* They are written **exactly as printed**. A figure that is not found stays **empty** (never 0.00) and **the bot calculates nothing**, not even a missing GST from the other two. If the figures disagree they are kept as read, with no warning.
* The only exception: when the document itself says *GST free* (or *no GST*, *GST exempt*) and no GST amount is printed, GST is `0.00`.
* The **Remarks** column in the Record says `GST NOT FOUND` **only when all three are missing**; otherwise it is empty (it can also say `SITE NOT FOUND`).
* **Statements, PODs, account summaries and reports are never touched** by this rule: no amounts, no `0.00`, no remark.
* One guard: if none of the three is found **and** only the e-mail subject (not the page) says it is an invoice, the document waits (`Hold documents without amounts when the type came only from the subject`; set it to No to switch the guard off).

---

## 5a. If the tick, flags or categories do not appear in Outlook
Outlook versions and mailbox types differ in what they accept, so the bot **checks every change by reading it back from Outlook** and tries more than one way. To see exactly what YOUR Outlook allows:
1. In the panel (Setup 1) type the mailbox (for example `jiya`) in **Only one mailbox?** and a word from a test e-mail's subject (for example `BOTTEST`) in **Only touch e-mails whose subject contains**.
2. Press **🧪 Test Outlook marking on ONE e-mail**. It tries *mark as read*, *the tick* (three ways), *the red flag* and *a category* on that single e-mail, reads each back, **puts the e-mail back exactly as it was**, and prints what works.
3. If the tick cannot be set at all, the bot adds the category **Processed by bot** instead (setting `Tick fallback category`; type `none` to switch it off) and the run summary shows `first problem:` with Outlook's own words.
(The test refuses an e-mail that already has a flag.)

## 5b. Storage: backups and the quarterly archive

* **The database** (`data\bot.db`) holds every record and stays small: about **1 GB a year** at 2,500 documents a day. It is never emptied, so a re-sent invoice is always recognised as a duplicate.
* **Backups** are made once a day into `data\backups\bot_YYYYMMDD.db.gz`: **compressed (about 4 times smaller)** and checked before they are kept; 14 are kept (`Backup keep days`).
  Point `Backup folder` at the NAS to keep them off the bot PC. To restore: close the bot, run `Advanced\Restore_backup.bat` (or `python restore_backup.py`), choose the number. The database it replaces is kept as `bot.db.before_restore_<time>`.
* **Quarterly records archive:** on the first run after a quarter ends (Jan–Mar, Apr–Jun, Jul–Sep, Oct–Dec) the bot writes that quarter's records to a **read-only Excel file** `Reports\Archive\Archive_2026-07_to_2026-09.xlsx`
  (sheets: Record, Needs attention, Corrections, About). A quarter at 2,500 documents a day is about 17 MB and takes about 40 seconds, once. **Nothing is deleted from the database.**
  Set `Records archive folder` to move the archives (NAS, SharePoint); `Quarterly records archive` = No switches it off. **Open → Quarterly archives** opens the folder.

## 6. The settings workbook (`Bot_Settings.xlsx`)

Open it from the panel (**Open → Settings**). The **About** sheet explains every sheet. Most days nobody needs it: the panel changes the important things.

| Sheet | Change it to … |
|---|---|
| **Run** | limits, month pattern, workers, OCR, categories, handler e-mails, archive days, minutes-saved estimate … (each row explains itself) |
| **Keywords** / **BlockedWords** | subject words that mean *download* / *flag red* |
| **DocTypes** | where a type goes, its file tag and name pattern |
| **Senders** | exception senders (always download) and sender → supplier hints |
| **Companies** | mailbox name in Outlook, **folder name**, text to look for on documents |
| **Suppliers** | spelling clean-up, optional *Folder override* (MILK INVOICES) |
| **Sites** | the 157 sites (re-import with `Advanced\Import_sites.bat`) |
| **Labels** | the words each supplier prints next to invoice no., date, GST, total — *starting hints you improve from real results* |

Columns are found by **heading text**, so adding or re-ordering columns is safe. Alternatives in one cell are separated by `|`. **Close the file before a run.**

---

## 7. Advanced (for the person who looks after the bot)

The `Advanced\` folder has double-click versions of the same things (`Check_setup`, `Demo`, `Rehearsal`, `Run_live`, `Run_with_dates`, `Test_with_msg_files`,
`Schedule_every_30_minutes`, `Import_sites`, `Run_tests`, `Install_packages`). The same functions from a command window:

```
python run_bot.py                       one live run (automatic catch-up)
python run_bot.py --shadow              rehearsal
python run_bot.py --from 2026-10-01 --to 2026-10-05      e-mails that ARRIVED in this period
python run_bot.py --mailbox "Colyton"   only one mailbox (part of its name)
python run_bot.py --demo                sample e-mails in a simulated Outlook
python run_bot.py --msg-folder C:\Msgs  the whole bot on a folder of saved .msg e-mails
python check_setup.py                   is this PC ready?
python selftest.py                      automated tests (no Outlook needed)
python scorecard.py C:\Samples          grade the reader on real invoices / .msg e-mails
python diagnose.py file.pdf --subject "Invoice 123"      why was this file read this way?  (also works on a .msg)
python import_sites.py list.xlsx --old old_list.xlsx     (re)load the site list
python restore_backup.py                lists the database backups;  python restore_backup.py 1  restores the newest (the current database is kept)
python run_bot.py --archive-quarter 2026-Q3     write the read-only records archive of a quarter
python benchmark.py 2500                stress test (needs: pip install reportlab)
```

**Optional: OCR** (scanned PDFs and logo-only suppliers). Install **Tesseract OCR** (UB Mannheim Windows build) to its default folder
`C:\Program Files\Tesseract-OCR`, then press **Check my PC**: it should say *Tesseract OCR ready*. Without it the bot works normally; scanned PDFs go to `_Unrecognised` with the reason.

---

## 8. Troubleshooting

| Symptom | Cause / fix |
|---|---|
| `Python is not installed` | Install Python from python.org and tick **Add python.exe to PATH**, then double-click `Start_Bot.bat` again. |
| The panel does not open | Copy the address printed in the black window into your browser. Keep the window open. |
| `Could not connect to Outlook` | Classic Outlook must be installed, **open** and signed in (not *New Outlook*), under the same Windows user that runs the bot. |
| `mailbox not found in Outlook` | **Setup 1 → Find my Outlook mailboxes** and save. |
| `CANNOT START: Live mode needs the real SharePoint path` | Section 3. |
| `CANNOT START: Bot_Settings.xlsx cannot be read` / `… is open in Excel` | Close the workbook in Excel (or restore `Bot_Settings.backup.xlsx`). |
| `Not started: another run is in progress` | A run (maybe the scheduled one) is still going; a crashed run's leftover lock expires by itself. |
| Many files in `_Unassigned` | The site could not be decided, or its folder does not exist yet in SharePoint, or the root is not synced. The bot retries every run; the *reason* is in `To_Fix.xlsx` (pick the **Site folder** there). |
| A good invoice went to `_Unrecognised` | `python diagnose.py <file> --subject "<e-mail subject>"` shows why. Usual fix: add the supplier's wording to sheet **Labels**, or an alias to **Suppliers**. |
| `Outlook update problems` in a run | The mailbox is read-only for this user. Files are still filed and not re-processed; ask IT for edit rights. |
| Path too long | Windows limits paths to 260 characters unless long paths are enabled; shorten the sync folder path. |

---

## 9. What was tested, and what could not be

**Tested automatically (`python selftest.py`, 309 tests):** reading logic on many invoice layouts, all e-mail rules, routing, naming, duplicates, month folders,
live-path handling, To_Fix corrections, locking, failures and retries, shadow mode, parallel reading, the `.msg` reader, the Excel reports (also rendered and inspected),
the real Outlook adapter against a mock of Outlook's object model, the control panel (controller, web server, security checks, and its JavaScript driven against the real server),
a 2,500-document run (31 s on one core, < 100 MB), and the e-mails and PDFs you supplied.

**Could not be tested outside your environment — verify during the rehearsal period:**
* the live COM connection to classic Outlook (isolated in `invoicebot/outlook_io.py`; the panel's mailbox finder uses it),
* your real OneDrive / SharePoint sync folders and path lengths,
* Windows-specific parts (Task Scheduler button, batch files, Windows memory read-out, Tesseract on Windows, how the panel looks in your browser),
* accuracy on **all** your suppliers' real invoices — use **Rehearsal**, then tune sheet **Labels**.

---

## 10. Layout

```
OutlookInvoiceBot\
  Start_Bot.bat        ← the one thing to double-click
  Bot_Settings.xlsx  config.py  README.md  run_bot.bat (used by the schedule)
  Advanced\            helper scripts for the person who looks after the bot
  invoicebot\          the program     tests\  automated tests     demo_data\  sample files     docs\  guides
  app.py  run_bot.py  check_setup.py  selftest.py  scorecard.py  diagnose.py  import_sites.py  benchmark.py
  — created while running —
  data\bot.db  Reports\  logs\  Holding\  TestOutput\  Scratch\  Archive\
```
