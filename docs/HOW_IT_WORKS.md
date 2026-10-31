# Outlook Invoice Bot — How it works

Technical and process documentation · version 5.1.0

## Contents
1. [Purpose and scope](#1-purpose-and-scope)
2. [Architecture](#2-architecture)
3. [One run, step by step](#3-one-run-step-by-step)
4. [Decision rules](#4-decision-rules)
5. [The reading engine](#5-the-reading-engine)
6. [Where files are saved](#6-where-files-are-saved)
7. [Data model](#7-data-model)
8. [Reliability and safety design](#8-reliability-and-safety-design)
9. [Performance and resource protection](#9-performance-and-resource-protection)
10. [Technologies used](#10-technologies-used)
11. [Testing](#11-testing)
12. [Module reference](#12-module-reference)
13. [Known limits and possible next steps](#13-known-limits-and-possible-next-steps)
14. [The control panel](#14-the-control-panel)
15. [The Excel reports](#15-the-excel-reports)

---

## 1. Purpose and scope

**Problem.** The group (13 legal entities, 157 sites) receives about 2,500 supplier documents a day by e-mail into 13 Outlook mailboxes.
People downloaded each attachment, renamed it, saved it in the right SharePoint folder, and recorded it in Excel.

**Solution.** A Windows program that does this unattended, safely and traceably:

* finds the right e-mails by **subject keywords**, **exception senders** and **blocked words**;
* downloads PDF/Excel attachments, **reads** them (text, tables, OCR for scans), and **verifies** the figures;
* names them `Supplier_InvoiceNumber.pdf` (+ `_CN`, `_POD`) and files them into
  `<root>\<MONTH> <YEAR> EMAIL INVOICES\<COMPANY FOLDER>`;
* keeps statements in their own `_Statements` folder and parks account summaries and similar reports in `_Unrecognised\<type>` (three holding folders side by side: `_Unassigned`, `_Unrecognised`, `_Statements`);
* keeps a **database** record and produces **Excel** records for staff;
* marks e-mails (read + tick, red flag, categories) **only after the file is safely stored**;
* lets staff correct anything through one Excel sheet (`To_Fix.xlsx`).

**Out of scope (by design).** Paying invoices, posting to the accounting system, replying to suppliers, deleting or moving e-mails.

---

## 2. Architecture

```
  Start_Bot.bat ─► app.py ─► panel.py (local web page: buttons, log, setup)  ── starts ──┐
                        ┌──────────────────────────── run_bot.py (command line) ◄────────────┘
                        │ options: --shadow  --from/--to  --mailbox  --workers  --demo  --msg-folder  --report  --check │
                        └───────────────────────────────────────┬───────────────────────────────────────────┘
                                                                ▼
 Bot_Settings.xlsx ───► settings.py ───────────────► pipeline.py  (orchestrator: one complete run)
 (rules in Excel)       (loaded once per run,           │   │   │
 config.py              snapshot to the workers)        │   │   │
                                                        │   │   └──► router.py    where a file goes; safe copy-verify-rename
        ┌──────────── gateway interface ────────────────┘   │                    month/company folders, holding folders, archive
        ▼                                                   ▼
 outlook_io.py  ClassicOutlook (COM)           worker.py ──► extractor.py  reads ONE file
 fake_outlook.py  simulated Outlook                │            ├─ pdftext.py / xlsxdoc.py   words + positions
 (tests, demo)                                     │            ├─ labelreader.py / legacy.py label + table reading (amounts as printed)
                                                   │            ├─ suppliers.py   who issued it
        Reader PROCESSES (parallel)  ◄─────────────┘            ├─ doctypes.py    what it is
                                                                ├─ companies.py / sites.py   which company / site
                                                                └─ ocr.py   Tesseract for scans and logos
 db.py  SQLite (data/bot.db)  ◄── results ──  pipeline        exports.py  Excel: Record, To_Fix, Report, Shadow
 runlock.py  logs.py  memory.py  notify.py  naming.py
```

Design principles:

1. **Rules live in Excel, not in code.** Keywords, blocked words, folder names, labels, limits: staff edit them; no programmer needed.
2. **One file talks to Outlook** (`outlook_io.py`). Everything else is tested against a simulated Outlook with the same interface.
3. **Verify, then act.** Read → check what is required → write file → verify copy → record in the database → *only then* change the e-mail.
4. **Never guess.** If a figure cannot be proven, the document goes to `_Unrecognised` with the reason.
5. **Every loop is bounded** (e-mails, attachments, pages, retries, run time) and every failure is isolated to one e-mail or one file.
6. **Be able to rehearse.** Shadow mode runs the full pipeline and changes nothing.

---

## 3. One run, step by step

| # | Step | Detail |
|---|---|---|
| 1 | **Lock** | `data/run.lock` is created; a second run refuses to start. A lock whose process is gone (crash, power cut, restart), or older than *max run time + 15 min*, is treated as a crash leftover. |
| 2 | **Settings** | `Bot_Settings.xlsx` is read once into a snapshot. If it cannot be read, or *Destination mode* is invalid, the run **stops with a message** (no silent defaults). Live mode without a path stops too. |
| 3 | **House-keeping** | Old scratch files removed. |
| 4 | **Staff fixes** | Fixes typed into `To_Fix.xlsx` are applied (supplier, number, type, company folder); affected files move to their final folder; each change is logged. |
| 5 | **Waiting files** | Files in `_Unassigned` (folder missing, supplier on hold) are retried, up to 50 times each. |
| 6 | **Per mailbox** | (a) find the mailbox in Outlook; (b) compute the **window**; (c) scan e-mails newest-first (capped); (d) classify each; (e) process wanted e-mails in **batches of 25**. |
| 7 | **Per batch** | download attachments to `Scratch\` → read in parallel → decide the route → write the file (verified) → insert into SQLite → commit → *then* update Outlook → record the e-mail as handled. |
| 8 | **Windows** | A mailbox's window is marked complete only if the scan was not truncated, the run did not stop early, and no e-mail is waiting for a retry. Otherwise the same period is read again next run. |
| 9 | **Exports** | `Record_YYYY-MM.xlsx` for months that received a document, `To_Fix.xlsx` when something changed; nothing is written by an idle run. |
| 10 | **Clean-up** | Archive older than *Archive days* deleted; daily database backup; month-end warning if next month's folder is missing. |
| 11 | **Summary** | Printed and written to `logs/run_<time>.txt`. |

**Date window.** Explicit `--from/--to`: From 00:00 to To 23:59:59, both inclusive, never beyond *now*.
Otherwise automatic catch-up: from `end of last successful window − overlap (60 min)`, at most *Catch-up max days* (31) back;
a mailbox never read before starts *Default catch-up days* (2) back.

---

## 4. Decision rules

### 4.1 E-mail classification (first match wins)
1. Already recorded in the `emails` table → skip. (An e-mail with outcome `error` is retried up to 3 attempts, then `gave-up`.)
2. **Blocked word** in the subject (whole words: `overdue`, `urgent`, `reminder`, `remainder`, `stop supply`, `warning`; editable) →
   never downloaded; red flag if no flag exists; stays unread; FW:/RE: is handled the same way (red flag, no category). Optionally the first lines of the body are scanned too.
3. **Wanted** = a subject keyword matches (longest keyword wins; whole words so `inv` does not match "winning") or the sender is an **exception sender**
   (matching ignores spaces/punctuation: "Rarity Wholesale" ↔ `raritywholesale.com.au`):
   * PDF/Excel attachments → download and process;
   * only pictures ≥ 40 KB (hidden inline signature images ignored) → orange category *Image attachment*;
   * nothing attached → untouched.
4. FW:/RE: gets the red flag (never a category) when it is a document e-mail (or always, if *Red category applies to* = `all`).
5. Everything else → ignored.

Outlook changes are limited to: **read**, **follow-up flag** (set only when no flag exists), **category**. A flag set by a person is never overwritten.

### 4.2 Document type (first match wins)
1. a document **title** found in the first 15 lines of page 1 (keywords ≥ 7 letters, longest first, earliest on a tie — "Tax Invoice … credit note policy" is an invoice);
2. the e-mail **subject**; 3. the attachment **file name**; 4. an exception sender's default (Invoice).
Types and routes are defined in the **DocTypes** sheet.

### 4.3 Supplier identification (first that works)
1. an alias from the **Suppliers** sheet found (whole words) in the first 25 lines;
2. a printed issuer line ending in *Pty Ltd / Limited …* (customer blocks, address lines, titles, "Total inc GST" lines are rejected);
3. the sender hint from the **Senders** sheet;
4. the first plain name-like line (weak guess);
5. the **logo**, read by OCR and matched (spelling-tolerant) against all known supplier names.
The result is normalised through the Suppliers sheet (e.g. "BDD Australia Pty Ltd" → "Bega Cheese Limited").

### 4.4 Company, then site, then folder
`document → site → sender → mailbox` (editable; needed only when the layout contains `{COMPANY}`). Document = a company named in the top of the document; site = the company that owns the matched site;
sender = the company mailbox the e-mail was sent or forwarded from (strict match on a full mailbox name or address); mailbox = the company of the mailbox it arrived in. A *Folder override* on the supplier (e.g. `MILK INVOICES`) beats all three. Disagreements are written as a warning.

### 4.5 Site matching (decides the folder when the layout contains `{SITE}`)
account number → address fragments (deliver-to block first, then text) → site name/aliases (longest match; a short name inside a longer one is ignored) →
**"Pearl Energy <town>"** (a site whose town is unique in the list also answers to this name, because invoices and dockets are addressed that way; towns with two sites are never matched by town alone) →
recipient name of the e-mail → spelling-tolerant name match. Conflicting evidence (address says A, name says B) → *UNASSIGNED*, never a guess.
Fragments such as `19 balook` or `pacific highway south kempsey` are generated from the site list by `import_sites.py`.
**Shared towns:** if the only evidence is the generic name "Pearl Energy <town>" and another candidate site has the same town, the document is ambiguous and waits (found by the 2,500-document test: *Pearl Energy Wagga wagga* and *Mobil Wagga Wagga* are two sites of one company). A street address, an account number, or the specific brand name settles it.
An unresolved site sends the file to `_Unassigned`; the person picks the **Site folder** in `To_Fix.xlsx`. The real folder name of each site is the Sites column **SharePoint folder**, filled by
`run_bot.py --match-folders` (exact match ignoring capitals/spaces, then a similar spelling, with a report of sites without a folder and folders without a site).

### 4.6 Naming, duplicates, collisions
Pattern per type (sheet DocTypes). Any date in a name is written dd-mm-yyyy (setting `File name date format`). Illegal characters `<>:"/\|?*` removed, length capped. A missing number falls back to the original file name.
Duplicate = same SHA-1 content, or same supplier + type + number. Same name, different content → `_2`. Identical file already present → skipped.

---

## 5. The reading engine

1. **Open** — `pdfplumber` extracts every word with its position. Page 1 is read first (enough to identify most documents). Long documents are refused
   *before* reading all pages (default limit 15 pages) — a 31-page catalogue is rejected in about a second.
2. **Scanned?** — fewer than 40 characters on page 1 → OCR (Tesseract at 200 dpi, first 3 pages); words keep their coordinates so the same reader works.
3. **Lines** — words are grouped by baseline, so a large label and its small amount are one line.
4. **Label reader** — for invoice number, date, GST, ex-GST, total: the supplier's own labels (sheet **Labels**) first, then the general ones; value on the same line,
   or under the heading (table layouts, wrapped headings, totals rows); `Due Date` is never taken as the invoice date; amounts in brackets are negative.
5. **The three amounts are optional (since 5.6.0).** Total incl. GST, GST and amount excl. GST are read as printed and never calculated; a figure that is not found stays empty
   (the `consistent` / `believable` tests are used only to CHOOSE between two readings of the same page, e.g. a totals row against label matches, so an item-table line is not taken as the total).
   If the first reader found none of the three, the fallback reader is tried (as printed). GST is `0.00` only when the page says GST free and no GST amount is printed.
   `Remarks` = `GST NOT FOUND` only when all three are missing. Statements, PODs and other types never go through this. The guard: no amounts and the type came only from the subject → waits.
6. **Cross-checks** — invoice number vs e-mail subject; future dates rejected; credit notes forced negative; page 1 text vs company/site.
7. **Special cases** — PFD has its own reader for its totals box; Excel invoices are turned into positioned words (`xlsxdoc.py`) and read the same way.

Documents without amounts (POD, statements, summaries) read only the number and date. Duplicate detection by number is used only for types whose number is a real identifier (invoice, credit note, POD); a statement is a duplicate only if its content is identical.

---

## 6. Where files are saved

```
<root>\<layout>\Supplier_Number.pdf        layout = the setting 'Destination pattern', for example  {SITE}   or   {MONTH_FOLDER}\{COMPANY}
```
* `root`: test mode → `TestOutput\` next to the program; live mode → a typed `Live destination root` (sheet Run) or `LIVE_DESTINATION_ROOT` (`config.py`), otherwise **found automatically** (`autofind.py`: a folder `EMAIL INVOICES` inside `OFFICE INDIA` under the user profile / OneDrive, listing folder names only, time-limited, remembered in `data/destination.txt`).
* The month text is built from `Month folder pattern` and the e-mail **arrival** date (or the invoice date if *Month basis* = invoice).
* Writing is **copy → verify size → rename** (`name.part` → `name`). A failure leaves no half-written file.
* Live mode creates missing month / company / site folders like TEST (5.15.0), but never the `EMAIL INVOICES` folder itself: if it is missing the run stops before anything is written.
* SharePoint is reached through the **OneDrive sync client**: the bot writes into the synced folder and OneDrive uploads. No Azure app or API secret is needed.
  (A direct Microsoft Graph upload would only be needed if sync is not an option.)

---

## 7. Data model (SQLite, `data/bot.db`)

| Table | Purpose |
|---|---|
| `docs` | one row per document: ref (R000001…), status (Filed / Parked / Unassigned / Unrecognised), type, supplier (+how found), number, date, amount, GST, company folder (+source), site (+source), mailbox, e-mail id/subject/sender/received, original name, SHA-1, folder, file name, current path, archive path, warnings, reader method, pages, OCR flag, attempts, timestamps, fixed-by |
| `emails` | every e-mail acted on: outcome (`downloaded`, `blocked`, `image-only`, `no-attachment`, `category-only`, `error`, `gave-up`), attempts |
| `mailbox_runs` | window start/end and whether it was complete (drives the catch-up) |
| `runs` | summary of each run |
| `corrections` | each change made through `To_Fix.xlsx` (ref, field, old, new, who, when) |

SQLite is a single file, needs no server and survives power loss (WAL journal). It is backed up daily (`data/backups/bot_YYYYMMDD.db.gz`, 14 days): each backup is written with SQLite's online-backup call, checked with `PRAGMA integrity_check`, compressed (about 4× smaller) and read back in full before it is kept; `restore_backup.py` puts one back (the old database is kept as `.before_restore_<time>`).
Excel files are **views** generated from it.

---

## 8. Reliability and safety design

| Risk | Protection |
|---|---|
| Two runs at once | lock file with stale-lock takeover |
| Bad settings | settings error stops the run; never silently falls back to defaults |
| Wrong destination | live mode requires an explicit path; invalid mode rejected; test and live are separate |
| Half-written file | `.part` copy + size check + atomic rename |
| Overwriting | never; identical → skipped, different → `_2` |
| Lost attachment | the e-mail is ticked only if **all** its attachments were handled; otherwise it is retried (3 attempts) and the date window stays open |
| Crash between "file saved" and "e-mail ticked" | the e-mail is processed again; duplicates are recognised by content, so nothing is saved twice |
| No edit right on a mailbox | counted and reported; files are still filed and not re-processed |
| One dead folder / NAS | remembered for the run (no repeated timeouts); files go to `_Unassigned` |
| Giant or hostile files | page limit, size limit, scanned-page limit, per-file try/except, reader processes replaced by fresh ones every 50 documents each |
| A run killed in the middle (crash, power cut, restart) | the run lock records the process id; the next run sees the process is gone and takes the lock over at once (before 5.4.2 it waited up to 60 minutes); files already placed are recognised as identical, so nothing is duplicated; a warning records it |
| A reader process that hangs or crashes | every file has a time limit (240 s); on a timeout or a crash the reader processes are killed, the rest of the batch is read one by one in the main program, and the run goes on (nothing is lost; a warning is recorded) |
| Low memory | below 1,500 MB free the bot reads one file at a time; below 500 MB it stops cleanly and resumes next run |
| Run too long | stops cleanly after *Max run minutes*; the rest is picked up next time |
| Wrong guess | required items (number, date, supplier, type) must be found; a supplier that is a sentence or a link is rejected; odd characters are tidied; `_Unrecognised` instead of a guess |
| Unattended failures | full log per run; summary with problems; daily DB backup; 7-day archive of the original files (the e-mail itself is the permanent copy) |
| Deleting by mistake | the clean-up deletes only date-named folders inside `Archive\` |

**Shadow mode** executes the real pipeline with a throw-away in-memory database, a dry-run router (no file written anywhere) and a gateway whose
change methods are never called; it writes `Shadow_Report_<time>.xlsx` listing, per e-mail and attachment, what a real run would do.

---

## 9. Performance and resource protection

* Reading runs in **separate OS processes** (`ProcessPoolExecutor`, default CPU cores − 1, at most 4), so several documents are read at once and memory is returned
  when the program replaces them (after the number of documents set in `Recycle worker after documents`). The pool is started with `spawn` on every system and replaced by the program itself: Python's built-in `max_tasks_per_child` option is NOT used, because it deadlocked in Python 3.12 (found by a 500-document test and fixed in 5.4.1). Outlook (COM) is used only by the main process.
* Batches of 25 e-mails bound disk use (`Scratch\`) and memory.
* Measured here (one CPU core, synthetic one-page invoices): **2,500 documents in 31 s, 88 MB peak**. Real invoices are heavier (multi-page, tables, scans);
  budget roughly 0.2–1 s per document per core and measure on your own PC with a shadow run (the summary prints time and peak memory).
* The 20 MB, 31-page catalogue that took 40 s and 1 GB with a naive reader is now rejected in 1.1 s.

---

## 10. Technologies used

| Technology | Used for | Why |
|---|---|---|
| **Python 3.10+** | the whole program | readable, rich libraries, easy for the team to maintain |
| **pywin32 / COM (`win32com`)** | automating classic desktop Outlook | the standard supported way to read mailboxes and set flags/categories without cloud permissions |
| **pdfplumber** (built on **pdfminer.six**, **pypdfium2**) | PDF text with word coordinates; page rendering | coordinates allow reading table layouts and "label … value" pairs reliably |
| **Tesseract OCR** via **pytesseract** + **Pillow** | scanned PDFs; supplier logos | free, offline, runs on the PC |
| **openpyxl** | settings workbook, Record, To_Fix, reports, Excel invoices | no Excel installation needed, supports validation lists, styles, write-only mode for big sheets |
| **SQLite** (`sqlite3`, standard library) | the record, e-mail ledger, windows, corrections | zero administration, transactional, fast, single-file backup |
| **`concurrent.futures.ProcessPoolExecutor`** (`spawn`) | parallel document reading | uses all CPU cores, isolates crashes; recycled by the program, with a time limit per file |
| **OneDrive sync client** | delivering to SharePoint | the bot just writes files; no Azure app registration or API secret |
| **Windows Task Scheduler (`schtasks`)** | running every 30 minutes | built into Windows |
| **`difflib`** | spelling-tolerant matching of logos and site names | standard library |
| **`hashlib` (SHA-1)** | duplicate detection | content-based, independent of file name |
| **`ctypes`** | free-memory and peak-memory read-out on Windows | no extra package |
| **`argparse`, `dataclasses`, `re`, `logging`-style logger** | command line, data structures, text rules, run log | standard library |
| **`http.server`** (standard library) + **HTML / CSS / JavaScript** | the control panel: a local web page with big buttons | nothing to install, works offline, looks the same on every PC, fully testable |
| **`unittest`** | the automated tests | standard library, no extra install |
| **Node.js** *(development only)* | drives the control panel's JavaScript against the real server in the tests | not needed at the customer site |
| **LibreOffice** *(development only)* | renders the Excel reports to images so their look can be checked | not needed at the customer site |
| **reportlab** *(development only)* | generating the test PDFs and benchmark | not needed at the customer site |

External services used: **none** (everything runs on the bot PC; no data leaves the company except through OneDrive's normal sync).

---

## 11. Testing

`python selftest.py` runs **309 tests** in about 30 s, using a temporary folder and a simulated Outlook:

| File | Covers |
|---|---|
| `tests/test_reader.py` (19) | invoice layouts, label/table reading, arithmetic checks, date/amount formats, per-supplier labels, supplier-name traps, credit notes, hints |
| `tests/test_rules.py` (53) | keywords, blocked words, exception senders, naming, companies cascade, site fragments and matching, windows, month folders, live-path rules, safe placement, dry-run router, archive clean-up, locks, database, backups |
| `tests/test_outlook_com.py` (11) | the **real Outlook adapter** (`ClassicOutlook`) against a mock of Outlook's documented object model, including a full bot run through it |
| `tests/test_money_rule.py` (15) | the three optional amounts: blank not zero, nothing calculated, GST free, Remarks only when all three are missing, statements untouched, the guard, odd characters, supplier sanity |
| `tests/test_archive.py` (16) | compressed and verified backups, restore (including a damaged backup and a running bot), the quarterly archive (bounds, content, read-only, once only, failure never stops a run, command line) |
| `tests/test_autofind.py` (13) | the automatic search for the destination folder, runs with no path typed, the whole chain EMAIL INVOICES → company → site → folder |
| `tests/test_site_layout.py` (25) | the folder layout, real folder names, matching the site list to the folders, the "Pearl Energy <town>" names, site routing, the To_Fix site choice, live-mode missing site folders |
| `tests/test_panel.py` (25) | the control panel: every button's logic, destination/mailbox/schedule changes, the real web server and its security checks, and the page's JavaScript driven in Node against the real server |
| `tests/test_msg.py` (8) | the `.msg` reader (including damaged files), the `--msg-folder` mode, scorecard and diagnose on `.msg` e-mails |
| `tests/test_workflow.py` (73) | the whole bot against a simulated Outlook: every e-mail rule and flag/category, folders and names, credit-note signs, duplicates, re-runs, shadow mode, live path with missing folders then released, To_Fix corrections, locking, damaged settings, read-only mailbox, failed downloads (retry and give-up), partial attachment failure, parallel readers, Excel reports, handler e-mails, the `--demo` command |

Additional tools: `scorecard.py` (grade the reader on a folder of real invoices), `diagnose.py` (explain one file), `benchmark.py` (stress test), `--demo`.

Real documents used during development: 9 supplier e-mails (`.msg`) and PDFs from the business — account summary, two statements, payment reminders, stop-supply
and overdue notices, a pest-control invoice, and a 31-page catalogue. Findings from them (wrong supplier names, logo-only suppliers, exception-sender spelling,
file-name numbers mistaken for invoice numbers, the catalogue's cost) were turned into fixes and regression tests.

---

## 12. Module reference

| Module | Responsibility |
|---|---|
| `run_bot.py` | command line |
| `config.py` | the optional place for the SharePoint / holding paths |
| `app.py`, `invoicebot/panel.py`, `invoicebot/panel_ui.html` | the control panel (launcher, controller + secure local web server, the page) |
| `invoicebot/quarterly.py` | the quarterly read-only records archive (calendar quarters; deletes nothing) |
| `invoicebot/autofind.py` | finds the synced destination folder on this PC by itself |
| `invoicebot/mailboxes.py` | find the Outlook mailboxes, suggest the company for each, write the Companies sheet |
| `invoicebot/pipeline.py` | one complete run: windows, classification, batches, routing, Outlook updates, fixes, exports |
| `invoicebot/outlook_io.py` | `ClassicOutlook` (COM); data classes `MailInfo`, `AttInfo` |
| `invoicebot/fake_outlook.py` | simulated Outlook (tests, demo, benchmark) |
| `invoicebot/extractor.py` | read one file → supplier, type, fields, company, site, problems |
| `invoicebot/pdftext.py`, `xlsxdoc.py` | words with positions from PDF / Excel |
| `invoicebot/labelreader.py`, `legacy.py`, `parsing.py` | field reading, verification, fallback, date/money parsing |
| `invoicebot/suppliers.py` | who issued the document |
| `invoicebot/doctypes.py` | keywords, blocked words, exception senders, document type rules |
| `invoicebot/companies.py`, `sites.py` | company folder cascade; site matching and site-list import |
| `invoicebot/router.py` | month/company paths, safe placement, holding folders, archive |
| `invoicebot/db.py` | SQLite layer |
| `invoicebot/exports.py` | the Excel reports (Summary + Record + Needs attention, To_Fix, shadow) and reading `To_Fix.xlsx` |
| `invoicebot/settings.py` | the settings workbook, defaults, snapshot |
| `invoicebot/worker.py` | reader process entry point |
| `invoicebot/msgfile.py`, `msgtest.py` | testing only: read saved `.msg` e-mails and run the bot on them in a simulated Outlook |
| `invoicebot/ocr.py`, `memory.py`, `runlock.py`, `logs.py`, `naming.py`, `notify.py`, `paths.py`, `demo.py` | supporting functions |

---

## 13. Known limits and possible next steps

* **Classic Outlook only** (COM). If the company moves to New Outlook, the gateway can be replaced by a Microsoft Graph implementation without touching any other module.
* **Accuracy depends on the Labels / Suppliers tables** for new suppliers; the shadow period and `scorecard.py` are the way to tune them. Unreadable documents are never lost: they wait in `_Unrecognised`.
* Documents that are **pictures only** (PNG/JPG) are flagged for a person, not read.
* A first-time supplier is accepted by default; set *Hold first-time suppliers* = Yes to require confirmation.
* Possible additions: direct SharePoint upload through Microsoft Graph, export to the accounting system, a small dashboard on top of the SQLite database.

---

## 14. The control panel

`Start_Bot.bat` checks that Python and the packages are present (installing the packages itself on the first start), then runs `app.py`, which starts a
small web server on **127.0.0.1** (this PC only) and opens the browser. The page is one HTML file with plain JavaScript (no framework, no internet).

* **Controller** (`panel.Controller`) holds all the logic and is tested without any web server: start a run as a **separate process**
  (`python run_bot.py …`) and stream its output; stop it (killing the whole process tree on Windows and removing the stale lock); read the overview from the
  database; open folders and files; change the destination (checks that the folder exists); change settings through a safe editor with a rolling backup
  (`Bot_Settings.backup.xlsx`); detect the Outlook mailboxes (in its own process, so an Outlook problem can never freeze the panel); create or delete the Windows scheduled task.
* **Security:** only 127.0.0.1; a random secret token is required on every request; the `Host` header must be local (blocks DNS-rebinding); the server never answers
  cross-origin pre-flight requests; untrusted text (e-mail subjects in the log) is inserted into the page as plain text, never as HTML.
* **Safety of the data:** the page cannot corrupt a run, because a run is an independent process with its own lock. Closing the browser or the panel does not stop a run in progress
  (the *Stop* button does, safely).

## 15. The Excel reports

Built with openpyxl's **streaming writer** (a month of 55,000 rows: about 5 seconds, 45 MB) and inspected visually (rendered with LibreOffice during development).

| File | Content |
|---|---|
| `Record_YYYY-MM.xlsx` | **Record**: every document with filters, frozen headings, number formats, colour-coded status, banded rows and a clickable *Open folder*; **Needs attention**: only the Unassigned and Unrecognised documents with the reason |
| `To_Fix.xlsx` | *How to use* guide, *To fix* with yellow input cells and **drop-down lists** (company folders and document types come from the settings), status colours, *Open folder* links; sheets for image-only e-mails, overdue/urgent e-mails and new supplier names; a hidden *Lists* sheet feeds the drop-downs |
| `Report_<from>_<to>.xlsx` | the monthly report for any period |
| `Shadow_Report_<time>.xlsx` | rehearsal result: a Summary (KPI tiles and a table per mailbox), a colour-coded *Would do* sheet, blocked words, new suppliers, skipped e-mails |

All sheets are set to print one page wide on A4 landscape.

**When files are written.** The pipeline notes whether a run changed anything the Excel files show (a document, a fix, a released file, a recorded e-mail). `Record_YYYY-MM.xlsx` is rewritten only for months that received a document (or when it is missing); `To_Fix.xlsx` only when something changed (or when it is missing). A run with nothing new writes no Excel file (the run summary says so). Measured: a month of 55,000 rows takes about 10 s to write, so this saves about 10 s on every idle run at month-end.

### 4.5b Company first, then site (since version 5.3)
The reader returns the text it used for the site; the main program first decides the **company** (document → sender → mailbox, in the order of the `Company source order` setting without `site`),
then looks for the site **among that company's sites only** (`sites.match(..., only_company=...)`; a town only has to be unique inside the company, brands are recognised in the whole list).
Rules: a document-named company is trusted (no site of it fits → the file waits); a sender/mailbox company is a weak clue (no fit → all sites are tried, and the difference is recorded as a warning).

## 16. Storage, backups and the quarterly archive

Measured at 2,500 documents a day: the database grows about 1.2 KB per document (about 1 GB a year; 66 MB for 55,000 documents); a compressed backup is about 4× smaller than the database;
a monthly Record of 55,000 rows is 4.5 MB; a quarter of 227,500 rows is a 17 MB archive written in 42 s with flat memory (about 50 MB, because rows are streamed from the database through `db.DocSource` and never held in a list).

The archive is deliberately a **copy, not a move**. Deleting old rows would make the bot forget invoices it has already filed (so a re-sent invoice from four months ago would be filed twice) and would split the audit trail, in exchange for saving about 200 MB per quarter.
The database therefore keeps everything; the quarterly Excel file is the human-readable, read-only record that can be kept or handed over, and compressed backups keep the backup folder small.

## 17. Marking e-mails in Outlook (since 5.5.1)

`ClassicOutlook.mark_done / flag_red / add_category` never trust their own assignment: after saving, the e-mail is fetched again from Outlook and the value is compared.
The tick is tried as `FlagStatus = complete`, then `MarkAsTask(complete)`, then `FlagStatus` + `TaskCompletedDate`; the red flag as `FlagStatus = marked`, then `MarkAsTask(today)`.
If no tick works, the read status is still saved, the category in `Tick fallback category` is added, and the error text lists what each way answered (it appears as `first problem:` in the run summary).
`diagnose_marking` (panel button, `run_bot.py --test-outlook-marking --mailbox X --subject Y`) tries each change on one e-mail without a flag, shows the Outlook version, mailbox type and connection mode, and restores the e-mail.


## 5.8.0 additions (in short)

* New document type **Order Confirmation** (keywords: order confirmation, order confirmed, order acknowledgement/acknowledgment): treated like a statement, `Holding\_Order Confirmations\<Month Year>`, named `supplier_site_date_OC`. Blocked words still block it.
* Month folders (`October 2026`) inside `_Statements`, `_Order Confirmations` and `_Unrecognised`. Month = document date when within the month window of the arrival month, else the arrival month (`router.holding_month`).
* Statements and order confirmations are named `supplier_site_date_STM` / `_OC`; `NoSite` when no site is found, with the remark `STATEMENT MISSING DATA` and a highlighted row and the column *Statement missing data* in To_Fix.xlsx. Older Bot_Settings.xlsx files are read with the new pattern and the new type without editing (`doctypes.rule`).
* `run_bot.py --tidy-holding [--apply]` (`tidy.py`): list-only first, then moves and renames old files into month folders, updating the records.

## 5.7.0 additions (in short)
* Forwards/replies: downloaded, filed, e-mail left unread with the red flag, never ticked.
* Supplier clues: Suppliers sheet → SupplierIDs (ABN / e-mail domain / account name) → line ending in a legal ending → `Account Name:` → sender → logo → sender name.
* Month folder: invoice date within ±3 months (yellow / lilac / orange / light-blue rows); the bot builds next month's folders in the last two days and moves waiting future-dated invoices.
* To_Fix: *New sites* (move a document into the site you type) and *Sites by company*.
