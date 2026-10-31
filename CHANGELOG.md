# Changelog

## 5.20.0

Small, safe changes from the review of the _Unassigned folder (no new options, LIVE paths unchanged).

* **Site names suppliers print, added to the Sites sheet once** (added, never replacing anything):
  Mobil Hasting River + "Mobil Port Macquarie Hastings" (Asahi credit note, 128 Hastings River Dr);
  Mobil Moss Vale + "Mobil Mossvale" (Bega); Ampol Mittagong + "Pearl Energy Mittagong" (PFD, 65 Bowral Road);
  Pearl Energy Beresfield + "Pearl Energy Beresfield South" / "Yangan Drive" (Slick Chicks);
  Pearl Energy North Wyong + "1-10 Amy Cl Wyong" (Neverfail delivery address).
* **Fleet Card statements are placed by the merchant NAME at the top** ("MOBIL MOE", "Pearl Energy Tomago"). The address under it
  is the owner's postal address (Moe -> Kirk St, Burleigh Heads -> Robina, Jennings -> Glen Innes, Tomago -> Hexham) and pointed
  to another site, so these waited in _Unassigned. Exactly one site name must appear; otherwise the old rules decide.
* **Account statements whose title is a picture** (Neverfail: "Overdue / Current / Total Amount Due") are statements
  (HOLDINGS\_Statements), not invoices. A page titled TAX INVOICE is never changed by this.
* Left as they are on purpose: Melitta "Mobil Yandina" (not in the Sites sheet), SodaStream "Caltex Thornton, 2 Weakleys Dr"
  (not the listed Thornton address), Sydney Kebab "Camden Kebab and Pizza" (not a group site) - they wait for a person;
  scanned PDFs stay in _Unrecognised (no OCR).
* Tests: 499 (new: test_v520.py, using the payments team's real Sites sheet).

## 5.19.0

* **Fleet Card statements are filed in the Fleet Card folder whichever mailbox they arrive in.** Some still come to company
  mailboxes (the supplier has not yet moved them all to mstatement@), and Jiya@ reads every e-mail. Before, outside the Fleet card
  mailbox they became 'Other' or an invoice; now they go to `<Fleet card destination>\<SITE>\<report issue date>_STM.pdf` and
  the FleetCard Record, never to the invoice folders.
* **Pearl Investment fuel invoices are filed in FUEL INVOICES whichever mailbox they arrive in** (`<COMPANY>\<INVOICE NO>_<SITE>.pdf`
  + the PI Invoice list). Outside the fuel mailbox only a page that names Pearl Investment Family Trust (or its ABN) counts, so an
  ordinary supplier invoice is never taken for a fuel invoice.
* **Saved once.** The same statement from two mailboxes is a duplicate (same file, or same merchant number + report issue date);
  fuel invoices by their invoice number, as before.
* LIVE paths and folder rules are unchanged. Bot_Settings.xlsx: DocTypes notes updated.
* Tests: 495 (new: test_v519.py).

## 5.18.0

* **Tax invoices are no longer named _POD.** A page titled TAX INVOICE stays an invoice even when it also says
  "Delivery Note 433201 ..." (Lactalis) or carries a "Delivery Docket" copy on the same page (PacPlus Payne). A page without
  "Tax Invoice" that says Delivery Docket / Delivery Note is still a POD.
* **Street number ranges match.** "102-104 WOLLOMBI ROAD" also reads as "102 Wollombi Road" and "104 Wollombi Road", so the Metcash
  "Cover Sheet - POD" for FGRC Metro Cessnock finds Pearl Energy Cessnock and is filed under THREE PEARL ENERGY - no longer under
  SHRI HARI INVESTMENT (the mailbox it arrived in was the only clue left before).
* **Metcash Customer Activity Reports are filed in the site folder** (`<MONTH> EMAIL INVOICES\<COMPANY>\<SITE>\
  Metcash..._<date>_CAR.pdf`), no longer parked in _Unrecognised or taken for a statement: the subject "Customer Activity Report"
  wins over a PDF titled Statement. No amounts; never in the A/C Excel.
* **Security: "Client No: 9307(1)-SA" is read as 9307-SA.** Some alarm panels print the area number in brackets on the page with the
  arm/disarm events (page 1 holds only the periodic tests). Those events were dropped as an unknown client, so the site showed
  NO REPORT all week. Fixed for Hillcrest 9307-SA, Hampstead Gardens 5652-SA, Torrensville 5657-SA and Sheidow Park 5679-SA
  (tested on the real October reports).
* Bot_Settings.xlsx upgraded once: DocTypes 'Customer Activity Report' route = Company folder.
* Tests: 489 (new: test_v518.py, with the 4 real security reports).

## 5.17.0

* **Company + town finds the site.** When a ship-to or sold-to address names one of the group's companies and a town, and that
  company has exactly ONE site named after that town, that site is used - even when the brand or the street is printed
  differently from the Sites sheet. Example (Bega adjustment note): "BP Greta, ttf Pearl Energy No.3, 122 New England Hwy, GRETA"
  -> Pearl Energy Greta (listed as 122 High Street). Never without the company: Greta also has Ampol Greta (Shri Hari), so a
  block that only says Greta still waits for a person. Two sites of the company in that town -> it waits too.
* Settings unchanged (no new rows needed).
* Tests: 481 (new: test_v517.py).

## 5.16.0

* **MobiGo:** its Tax Invoice prints only "Order #" - that number (e.g. S506325) is now the invoice number (Labels row MobiGo:
  order # | order no | order number). A label with "#" may read the value under it.
* **Order documents are Order Confirmations, not _Unrecognised** (filed in `HOLDINGS\_Order Confirmations\<Month Year>\`,
  named Supplier_Site_Date_OC, never in the Record or the A/C Excel):
  * a page that says "THIS IS NOT AN INVOICE" / "not a tax invoice" / "Indicative total" (Fresho orders from Rarity Wholesale),
    even when the subject says Invoice or the sender is an exception sender;
  * a "Sales Order" (PFD) or an "Order # S00372" printout (MobiGo) when the page has no "Tax Invoice", "Invoice No/Date",
    credit note or statement and not the word "invoice" anywhere.
  A page titled TAX INVOICE always stays an invoice (also when it mentions its Sales Order No or carries a "not a tax invoice"
  remittance slip), and a statement that says "this is not a tax invoice" stays a statement. "Sales order" is NOT a keyword:
  subjects like "Sales Invoices for Sales Order" are still invoices.
* **Peters:** INV DATE 220726 is read as 22/07/2026. A date label written with "(ddmmyy)" in sheet Labels reads that compact
  format (Labels row Peters: inv date (ddmmyy)); a bare 6-digit number is never a date anywhere else, and the DUE DATE is never used.
* Dates written month first ("Thu Jul 23, 2026") and after a time ("Received: 9:46 AM Thu Jul 23, 2026") are read; "received" is a
  last-choice date label.
* Bot_Settings.xlsx upgraded once: Labels rows MobiGo and Peters (an existing row for them is completed, never replaced).
* Tests: 476 (new: test_v516.py with rebuilt MobiGo invoice + order, Fresho order, PFD sales order and Peters invoice).

## 5.15.0

* **The real LIVE folder.** `Live destination root` is now `C:\SHAREPOINT\pearlenergy.com.au\Patel Group - Documents\OFFICE INDIA\EMAIL INVOICES`
  (it was empty, and the automatic search only looked under the user folder, so the first LIVE run stopped with "SharePoint folder
  was not found"). The search now also looks in `C:\SHAREPOINT`. The path can still be changed in Bot_Settings.xlsx (sheet Run).
* **HOLDINGS and Reports sit inside EMAIL INVOICES** (`EMAIL INVOICES\HOLDINGS\...`, `EMAIL INVOICES\Reports\...`), made when
  missing. `Holding root` / `Reports root` blank = inside the live folder; a typed path still wins. TEST is unchanged (all local).
* **LIVE creates missing folders exactly like TEST**: month, company and site. An existing folder is always used, never made twice.
  The only folder never created is `EMAIL INVOICES` itself: if it is missing the run stops before anything is written.
* **Sold to before the bill-to.** Without a Ship to / Deliver to, the **Sold to** address decides the site (Bega / BDD Australia:
  "Invoice to" is the payer, "Sold to" is the real site). A sold-to that is not a known site waits in To_Fix (shown as the deliver-to,
  and learned when staff choose a site) - the payer is never used. A printed ship-to still always wins.
* **The amount payable wins.** The total words now start with total payable / total amount payable / amount payable, then the inc-GST
  totals (also "total incl taxes"), then amount due, and "balance due" only near the end (a paid invoice says 0.00). When a total
  label is followed by several amounts (Asahi: "Total AUD 1,125.56 94.08 1,219.64 116.80 1,336.44"), the amount that is the sum of
  the others is taken (1,336.44), never simply the first one; if nothing adds up the largest is taken with a note. Asahi row added
  to Labels. A supplier's own Labels row still comes first.
* **"Never read" really means never.** 5.14.0 removed only one of a company's two mailbox names, and Outlook's loose name matching then
  found the mailbox by the other name, so every mailbox was read. Now a mailbox set to Never read can never be reached by any name;
  the run summary says "set to Never read - not read". (A company is not switched off automatically: that would also switch off
  Pearl 1 when a test mailbox added to it is set back to Never read.) A company row without any mailbox name is skipped, not a crash.
  Setting a TEST mailbox to Never read also removes its TEST row.
* **New mailbox choice "📥 Read every e-mail (company from the invoice)"** (Setup step 2, any number of mailboxes): reads every
  e-mail of the Inbox and its sub-folders and files it exactly like a company mailbox; the company and site come from each document
  (then the sender), never from the mailbox. No TEST folder.
* **Bot_Settings.xlsx upgraded once** (the payments team's file is the base): live paths as above, Auto-find folder name
  EMAIL INVOICES, "Advice no" moved from the invoice DATE words to the invoice NUMBER words, total words in payable order,
  Asahi Labels row, Ampol Kelso + "17 Great Western Hwy" (19 Sydney Road kept), Jiya@ = Read every e-mail (TEST JIYA row removed).
  Mobil Singleton / Dunolly and the Bega "sold to" row are kept as written.
* Tests: 464 (new: test_v515.py, with rebuilt Bega and Asahi invoices and the payments team's settings file).

## 5.14.0

* **A product brand is never taken as the supplier.** The Suppliers sheet's "Name contains" is now only compared with the supplier's
  own header block (the lines above "Sold To / Deliver To" and the item table). Before, a Rum City Foods credit note with
  BRAND = SUNTORY in its item table was filed as Suntory Beverage & Food Australia.
* **A printed "T/A" trading name wins:** "Setlex Pty Ltd T/A" + "Rum City Foods" is recorded as Rum City Foods.
* **The sender's own e-mail domain is a strong clue** for e-mails sent straight from the supplier (never for FW/RE forwards, never
  Pearl's own or free-mail addresses): a domain in the supplier list gives that supplier's name, otherwise the words on the page
  that match the domain.
* **Rum City Foods** added to the supplier list (ABN 38 509 899 256, rumcityfoods.com.au); older settings files get the row.
* **Companies can be switched off in Setup step 2:** an Active tick box per company. An unticked company is NOT read by any run
  (also not the automatic ones) until it is ticked again. Safeguards: a confirmation when switching off more than one, the row says
  "switched OFF - not read", an orange banner at the top of the panel names them with a "Switch them all on again" button, and
  every run summary names them. E-mails of a switched-off company are not lost: its next run catches up (up to 31 days back).
* Note: a supplier now named from the supplier list instead of the Suppliers sheet can get the list's shorter name (e.g. "Rarity"
  instead of "Rarity Wholesale"), as agreed: the company's supplier list wins.
* Tests: 442 (new: test_v514.py, with a rebuilt Rum City credit note).

## 5.13.0

* **Setup step 2 shows what every mailbox is:** the drop-down offers the 13 companies, the special mailboxes
  (⛽ Fuel invoices, 💳 Fleet Card statements, 🛡 Security reports), 🧪 TEST mailbox and 🚫 Never read. "Leave as it is" is gone -
  nothing is left blank (a mailbox that fits nothing is set to Never read). P I Fuel Invoice, mstatement@ and Security show their role
  in purple with a link to their folder; office@ shows "Never read". Save these writes companies to the Companies sheet and the special
  roles to the SpecialMailboxes sheet (a mailbox is only ever in one of them). Only one mailbox per special role: the panel and the bot
  both refuse a second one. Making a mailbox "Never read" asks for confirmation.
* **HOLDINGS folder:** the holding folders (_Statements, _Order Confirmations, _Direct Debit Notices, _Account Activity Listings,
  _Unassigned, _Unrecognised) now live in OFFICE INDIA\MAILS\HOLDINGS. Reports stay in MAILS\Reports. An older settings file whose
  Holding root is still MAILS is switched once; files already in MAILS are not moved. TEST mode still keeps Holding local.
* **Fuel invoices of a customer that is not one of the companies** (e.g. SKM DIESEL PTY LTD) go to a folder named exactly as the
  Customer's company is printed: FUEL INVOICES\SKM DIESEL PTY LTD\<INVOICE NUMBER>_<SITE>.pdf (created when missing; an existing
  folder with another spelling is reused). They are not added to the Companies sheet.
* Tests: 437.

## 5.12.0

* **New panel layout:** Setup first (1 Where the documents go, 2 Outlook mailboxes, 3 Run automatically, 4 Security reports), then
  "Invoices: what would you like to do?", "Security: what would you like to do?", Needs your attention, Security, Open - and the run
  log last. New Open buttons for the fuel invoice list, the Fleet Card record and the security files.
* **TEST mode never writes into SharePoint:** in TEST, Holding, Reports, fuel invoices, Fleet Card statements and security files all
  go to folders next to the program (TestOutput\FUEL INVOICES, TestOutput\FLEETCARD STATEMENTS, TestOutput\SECURITY REPORTS),
  whatever paths the settings hold for LIVE.
* **Mailbox roles corrected (SpecialMailboxes sheet, older files are upgraded once):** P I Fuel Invoice = Fuel invoices,
  mstatement@ = Fleet card, Security = Security, office@ and serviceinvoices@ = never read.
* **Fuel invoices (P I Fuel Invoice):** the Inbox AND every folder under it are read (Outlook moves each company's invoices into its
  own sub-folder; companies without a sub-folder stay in the Inbox). Every e-mail with a PDF there is read, keyword or not. The
  invoice's **Customer** decides the company and the site (never the Reference); filed as
  FUEL INVOICES\<COMPANY>\<INVOICE NUMBER>_<SITE>.pdf. Listed in Reports\PI_Invoice_list_<MON>_<YEAR>.xlsx (one per month, by
  download date): Download Date, Contact, Invoice Number, Reference, Invoice Date filled by the bot; XERO Bill / ABA / QUICK FUEL
  left for staff - new rows are only ADDED at the bottom, rows already there are never changed, an invoice number is listed once.
  A Customer company that is not in the Companies sheet waits in _Unassigned. Tested on three real invoices.
* **Fleet Card statements (mstatement@):** FLEETCARD STATEMENTS\<SITE>\<report issue date>_STM.pdf, listed in
  Reports\FleetCard_Record_<MON>_<YEAR>.xlsx (Merchant No, Report Issue Date, Total payment, Site, Company, Status, File name;
  sheets Record and Needs attention).
* **Security bot (new, separate from the invoices):** its own buttons (Run security now, Security rehearsal, Security for dates), its
  own weekly schedule (Setup step 4) and its own settings file **Security_Settings.xlsx** (Company > Site > Client No > address >
  opening and closing time for every weekday; built from the October STORE OPEN & CLOSE workbook: 119 sites). It reads the weekly
  alarm reports of the Security mailbox (both PDF layouts, several sites per file) and fills the fortnightly
  SECURITY REPORTS\<MONTH YEAR>\SITE_DATA_<MON>_<YEAR>_<dd-mm>_TO_<dd-mm>.xlsx; the reports are kept in
  <MONTH YEAR>\Downloaded Reports. Rules: OPEN = first disarm (up to 3.5 h before opening counts), CLOSE = last arm (up to 4 h after
  closing counts); OPEN label yellow when more than 10 minutes late or missing, CLOSE label yellow when even 1 minute early or
  missing; NO REPORT and closed days highlight both; the Timing cell is never highlighted; remarks in red; seconds ignored; times
  used as printed. Reports with no arm/disarm or an unknown Client No are skipped (e-mail left untouched, Client No named in the
  summary); used e-mails are read + ticked. Never runs at the same moment as an invoice run (it waits). Tested on the five real
  reports: Tomingley (7 early closes, the 3:58 disarm on Thursday), Deeragun (Saturday open missing, Sunday closed 2 min early),
  Wangaratta (Saturday and Sunday opened late) - as the company's staff had filled them.
* **Quarterly archives** also for fuel invoices, Fleet Card statements and security events (Reports\Archive, read-only, nothing deleted).
* Commands: run_bot.py --security [--shadow] [--from --to]; run_bot.py --import-security-sites <workbook>.
* Tests: 428 (new: test_security.py on the real reports; fuel, Fleet Card and TEST-mode tests in test_v511.py).

## 5.11.0

* **Only the ship-to / deliver-to address decides.** When a document prints a deliver-to address, the bill-to is never used. A ship-to that is not a known site waits in _Unassigned (To_Fix shows the address as printed in the new column "Deliver-to as printed") - it is never filed under the bill-to site (the Unilever Preston / BP Orange North problem). Site names match in any word order and with small street-name typos ("BP NORTH ORANGE" = BP Orange North; the "hanharan" typo in the site list is corrected).
* **A waiting ship-to is filed by itself** when (a) the site is added to the Sites sheet, (b) staff create its folder in SharePoint (month > company > site: the bot finds PEARL ENERGY PRESTON inside its company folder), or (c) a person chooses the site once in To_Fix: the bot then remembers that supplier's deliver-to number / address for the next invoices. Learned sites are listed in To_Fix > Learned sites and can be removed there (Forget = YES). Run settings "Learn sites from staff fixes" and "Look for the site in the SharePoint folders".
* **LIVE folders:** a missing month or company folder is created (only when it does not exist); a missing SITE folder is NOT - the invoice waits in _Unassigned until staff create it. Older Bot_Settings.xlsx files are switched over once (Run settings "Live mode creates missing folders like test mode" = No, month Yes, company Yes, site No).
* **Special mailboxes (new sheet SpecialMailboxes):** P I Fuel Invoice = Fleet card, mstatement@ = Statements, Security = On hold (connected, nothing touched), office@ and serviceinvoices@ = Never. A special mailbox is never a company clue and is shown greyed out in "Find my Outlook mailboxes". A shared mailbox that Outlook does not list is opened by its address through the signed-in account.
* **Fleet Card statements** (Merchant Payment Advice, P I Fuel Invoice only): TEMPORARY DOWNLOADS\FLEETCARD STATEMENTS\<SITE>\<report issue date>_STM.pdf; the site folder is created only if missing; the merchant number is remembered for its site. Anywhere else they are handled as Other. Tested on the real 01_SEP_26 statement (Business Fuel Cards, Ampol Farleigh, 01-09-2026).
* **mstatement@ statements:** TEMPORARY DOWNLOADS\M STATEMENTS\<SITE>\<statement date>_STM.pdf.
* **Statements** now go to _Statements\<Month Year>\<COMPANY>\ (no site folder; the site stays in the file name). New document types with their own holding folders: **Direct Debit Notice** (_Direct Debit Notices\<COMPANY>\Supplier_date_DDN) and **Account Activity Listing** (Campbells / C-Store: _Account Activity Listings\<SITE>\Supplier_date_AAL). DocTypes has a new column "Folder layout". Old statements are not moved.
* **Holding and Reports in SharePoint:** new Run setting "Reports root"; the shipped settings point Holding to OFFICE INDIA\MAILS and the reports to MAILS\Reports. A configured folder that does not exist (unsynced) stops the run instead of creating an empty local copy. The database, logs and Scratch always stay next to the program. A OneDrive conflict copy of To_Fix.xlsx is reported in the run summary.
* **Record:** new columns "Downloaded by bot" and "Run marker" (the last e-mail of each mailbox in the latest run, bold with a box).
* **A/C Excel** in the company's MASTER format (NAME OF COMPANY ... REMARKS): filed invoices, credit notes and dockets for a period; dockets fill only the docket columns, credit notes negative, SHARE POINT opens the file. Panel button "Make the A/C Excel for these dates", command run_bot.py --ac --from --to.
* **Panel:** the tiles open their folders; the Unassigned / Unrecognised tiles also show the total still waiting; buttons for the latest A/C Excel and the Unassigned folder; nothing opens by itself.
* Supplier names: a greeting or a field label ("Dear Valued ...", "Billing To ...") is never taken as a supplier. Business Fuel Cards and Campbells added to SupplierIDs.
* The .msg test mode never writes into SharePoint (all output folders are switched to its own test folder).
* Tests: 26 new (test_v511.py), 412 in total.

## 5.10.0

* **LIVE mode now works like TEST mode for folders.** In LIVE a missing month, company or site folder is created by the bot, exactly as TEST does in TestOutput, so a document is no longer parked in _Unassigned only because a folder is missing. New Run setting "Live mode creates missing folders like test mode" (Yes); No returns to the three separate "Create ... folder in live mode" switches (all No). Older Bot_Settings.xlsx files get the new row at the next start.
* **No second folder beside the real one.** Before creating a folder, the bot looks for an existing one with the same name apart from capitals, extra spaces and punctuation (for example AMPOL  KELSO and AMPOL KELSO) and uses it. This also applies when a staff fix moves a document to a site.
* Tests: 4 new tests, 367 in total.

## 5.9.9

* **The ship-to / deliver-to address wins over the bill-to.** When a document shows both an "Invoice to" (bill-to) and a "Deliver to / Ship to" address (for example Unilever), the SHIP-TO block now decides the site AND the company. Before, both addresses counted equally, so two different sites meant "ambiguous" and a person had to decide, or the bill-to company sent the document to the wrong company folder. New Run setting "Ship-to address wins over bill-to" (Yes); No restores the old rule.
* Safe by design: if the ship-to block itself names two sites, or names a different company than the site belongs to, the careful older rules decide. A supplier account number set in the Accounts sheet still comes first. A document with only a bill-to address works as before.
* Tests: 10 new tests (two real Unilever-style PDFs with two address columns), 363 tests in total.

## 5.9.8

* **E-mails with a red flag are not read again.** Like the tick (flag complete), a red flag set by the bot or by a person now makes the bot skip the e-mail. New Run setting "Skip e-mails already flagged in Outlook" (Yes). Set it to No to read flagged e-mails again. The panel tile now reads "already ticked or flagged". Older Bot_Settings.xlsx files get the new row at the next start.

## 5.9.7

* **Multi-site suppliers keep one name.** Suppliers sheet: "Peters", "Peleguy" and "Nippy" now map to *Peters*, *Peleguy*, *Nippy's* for every site and every ABN, also when no ABN or e-mail from the list is on the page. Older Bot_Settings.xlsx files get the three rows at the next start.
* **The supplier list wins.** A supplier found by its ABN, e-mail domain or account name from the list is checked BEFORE the Suppliers sheet, and its name is no longer renamed by an older Suppliers-sheet alias (for example *Bega*, not *Bega Cheese Limited*). The Suppliers sheet still works for suppliers that are not on the list; its short name and folder override still apply.

## 5.9.6

* **A supplier that is not on the list is also caught when its name is printed LOW on the page.** If no "Pty Ltd / Pvt Ltd / Limited ..." name is found in the top lines, the rest of page 1 is searched for a legal-ending name printed next to the supplier's ABN, or after "Remit to / Pay to / Issued by / Trading as". Delivery and customer lines, your own companies and bank-box lines are never used; a Pty Ltd name with neither clue is not taken.

## 5.9.5

* **The company's supplier list is built in** (80 rows, 77 suppliers: name, ABN, e-mail domain). A document is now recognised from the supplier's ABN or e-mail domain alone, whatever the layout. Free-mail addresses (Gmail, Outlook, Yahoo) and e-mail-less rows are ABN only. Suppliers with several sites (Peleguy, Peters, Nippy's) have ONE name and one row per ABN: the site is worked out separately.
* An older Bot_Settings.xlsx gets the new SupplierIDs rows by itself at the next start (nothing already there is changed).
* AO Global is now named "AO Global" (from the list) instead of "AO Global Pty Ltd".
* New tests check every ABN and every domain in the list.

## 5.9.4

* **The run result explains every e-mail.** New tiles and summary lines: "Skipped: already done" (ticked in Outlook, or handled by the bot before), "E-mails with documents read". The mailbox line in the log now shows "already done N". Before, "E-mails read 609" next to 78 documents looked like most e-mails were not read.

## 5.9.3

* **E-mails that already have the tick are never read again.** Before, the bot trusted only its own database, so a new install, a new PC or a deleted `data` folder made it read and file every e-mail of the catch-up period again, ticked or not. Now an e-mail with the tick (flag complete) or the "Processed by bot" category is skipped and counted ("already ticked in Outlook"). Switch: Run sheet, `Skip e-mails already ticked in Outlook` (Yes).
* **A damaged or very slow file can no longer hold up a run.** Every file is read with a time limit (120 s); a file that takes longer, or crashes the reader, is skipped and listed as unrecognised for a person.
* **The run log shows progress:** "N e-mails found, M with documents to read" and "reading e-mails 1 to 25 of 140 (X documents done)".
* Damaged PDFs no longer print one "Data-loss while decompressing corrupted data" line per broken object.

## 5.9.2

* **Panel redesign (same buttons, calmer layout):** "Needs your attention (today)" now sits right under the four big buttons. The testing options (one mailbox, subject filter, demo, .msg files) are folded under "More options". Setup is a fold-out section that opens by itself until the bot is LIVE.
* A TEST / LIVE banner says where the documents go. A line shows whether automatic running is on. Messages stay visible when you scroll.
* Fixed: the date box and other hidden boxes were showing on start; the activity log is small until a run starts; the panel fits a phone or narrow window; keyboard focus is visible.

## 5.9.1

* **Outlook must really be open.** Before, a closed Outlook was silently started in the background, so the panel looked connected. Now the bot says "Outlook is not open" and does nothing.
* **Panel numbers are for TODAY** (waiting, unrecognised, filed, statements, handled). Open items from earlier days stay in To_Fix.xlsx.
* "Find my Outlook mailboxes" names the companies whose mailbox is NOT signed in on this PC (the run skips them there).
* Setup step numbers show as circles again (they were stretched ovals).
* Wording: FW:/RE: e-mails get a red flag (the Run-sheet note and demo data said "red category").

## 5.9.0

* **Order confirmations without a keyword in the subject:** an e-mail whose attachment FILE NAME contains "order" is opened; the document is filed only if the page itself is an Order Confirmation, otherwise it is ignored and the e-mail is noted so it is not opened again (setting `Check attachment names for order confirmations`).
* "ORDER DATE" is read as the document date. "CRM ID" is a last-resort number label (credit notes). "Tax" counts as the GST label when there is no GST line. Older Bot_Settings.xlsx files get these without editing.
* An invoice number printed under the "Tax Invoice" title with no label is read (not the ABN, not a date).
* The bare words "ABN", "Tax Invoice", "Invoice", "Tax" are never a supplier name.
* SupplierIDs: Statewide Oil (ABN 43 009 283 363), KC Supplies (ABN 85 105 637 901).
* Regression set `tests/layouts/`: the invoice layouts that failed on real e-mails (Dorries, KC Supplies, AO Global, Tip Top) are tested on every change.

## 5.8.2

* Invoice number printed BELOW its "Invoice #" label (columns side by side) is read. The supplier printed beside "Invoice # / Invoice To" labels is found. "Payment Due", "Terms", "Discount", "Date" can never be taken as a supplier name. Fixes AO Global Pty Ltd invoices.

## 5.8.1

* Supplier clue: when the page text runs "Bill to  Ship to  <Supplier Pty Ltd>" on one line (the supplier's address block printed to the right of the two customer columns), the name after the labels is the supplier. Fixes Dorries Anglers Choice Pty Ltd invoices going to _Unrecognised.

## 5.8.0

* New type **Order Confirmation**, treated like a statement: `Holding\_Order Confirmations`, named `supplier_site_date_OC`. A blocked word in the subject still blocks it.
* **Month folders** (`October 2026`) inside `_Statements`, `_Order Confirmations` and `_Unrecognised` (document date within the month window, else the arrival month).
* Statements and order confirmations are named **supplier_site_date_STM / _OC**. No site found: `NoSite` in the name, and the row is highlighted in To_Fix.xlsx with the new column **Statement missing data**.
* **One-time tidy**: `python run_bot.py --tidy-holding` lists what would move; add `--apply` to move and rename the old files.

## 5.7.1

* FW:/RE: e-mails get the red follow-up flag (no red category) instead (attachments still downloaded and filed; e-mail left unread, no tick).
* The image-only orange category and the "Processed by bot" fallback category stay as they were.
* To_Fix.xlsx, sheet "Sites by company": the Handler column is removed.
* Supplier clue: the supplier's own e-mail address printed on the document (hello@goldcoastcookies.com.au) names the supplier ("Gold Coast Cookies") when nothing else matches. Fixes Gold Coast Cookies invoices going to _Unrecognised.
* Tip Top Bakeries added to SupplierIDs (ABN 45 008 429 632, or the name anywhere on the page), so its documents are no longer named "A Division of George Weston Foods Limited". Older Bot_Settings.xlsx files get the row added automatically at the next run.

## 5.7.0 — forwards, dockets, suppliers without a name, the month window (from your real run)
* **FW: / RE: attachments are downloaded** whatever the subject. A forward/reply is left **UNREAD**, with **no tick** and the red category *Forwarded / Reply* (before: red category only, nothing downloaded). Unrecognised ones go to `_Unrecognised`. Blocked words still block. Switch: `Download attachments of FW/RE e-mails`.
* **Dockets** (`docket`, `del docket`, `delivery note`, `goods received`) are downloaded and filed like an invoice as `Supplier_Number_POD.pdf`, without amounts and without `GST NOT FOUND`. A page that says it is a delivery docket (and is not an invoice) is a POD even if the subject says invoice.
* **Supplier**: new sheets **SupplierIDs** (ABN / e-mail domain / account name — any one is enough; JP Bakery and Bonfect are pre-filled) and **LegalEndings** (Pty, Pvt, Private, Limited, Ltd …, one editable list). A name cut before its ending (`MCALLISTER & SONS PTY.` / `LIMITED`) is read; `Account Name:` in the bank box is used; the sender's display name is the last resort. Pty/Pvt/Limited/Ltd count as the same when names are compared. `Mc…` names are written `McAllister`.
* **Document type from the title only**: words inside the supplier's name (`Snackfood SERVICES`) no longer make an invoice; the title is also searched in the positioned lines (titles the plain text scrambles, e.g. Bonfect's `TAX INVOICE`).
* **Reading fixes found with your PDFs**: PDFs that deliver letters one by one (Bonfect totals) are joined, so total/GST/ex-GST are read (before: no total and a GST of 1.00); zero-width characters are removed; `26 Sept 2026` is a date (the Envisage invoice was taking the due date); a GST larger than the total allows is left empty; the fallback reader is trusted only when it found the total.
* **Site**: sheet **SiteAliases** (other names for a site; `Pearl Energy Wyong` / `10 Amy Close` → Pearl Energy North Wyong is pre-filled); within a known company the town plus a street named anywhere on the document (`Queanbeyan Service Station – Lowe St`) finds the site.
* **Duplicates / reprints** are skipped and the original row's Remarks says so. A **spreadsheet that arrives with an invoice** is filed beside it as `<invoice name>_breakdown.xlsx`.
* **Month window**: the month folder follows the invoice date within 3 months (older → its own month, yellow; outside the window → current month, lilac; future → waits in the current month, orange, and is **moved** into its month in the run that builds that month's folders; November lists it again in light blue, counted once). Notes only in Remarks. Settings `Month folder follows`, `Month window (months)`. The old `Month basis = invoice` still works.
* **Next month's folders are built by the bot** in the last two days of the month (folder structure copied, never files) and on the first run of a new month if missing (`Create next month's folders`).
* **To_Fix gets two sheets**: **New sites** (documents filed in the company folder with `SITE NOT FOUND`; type the site, the next run moves the file; a name not on the list creates nothing and Remarks says `SITE NAME NOT ON LIST`) and **Sites by company** (every company with its sites).

## 5.6.0 — the three-level path, and the amounts are optional
* **Month > company > site is the shipped layout** (`<root>\OCTOBER 2026 EMAIL INVOICES\ONE PEARL ENERGY\MOBIL BALLINA\file`), the shape of your SharePoint. If the site cannot be decided the document is filed in the
  company folder with the remark `SITE NOT FOUND` (setting `File in the company folder when the site is not found`). **Match my sites to the folders** now reads the month, then each company folder, then its site folders.
* **Total incl. GST, GST and amount excl. GST are optional.** Missing ones stay empty (never 0.00), nothing is calculated (the old "GST worked out as total minus ex-GST" is gone), disagreeing figures are kept as read,
  and none of it holds a document back. A GST-free invoice shows GST 0.00. New Record columns **Amount excl GST** and **Remarks** (`GST NOT FOUND` only when all three are missing). Statements, PODs and other types are never touched.
  Guard: no amounts and a type that came only from the subject waits (switchable).
* **Odd characters** (Symbol-font or non-breaking hyphens, invisible characters) are tidied in invoice numbers, supplier names and file names (no more `INV▢3121`).
* **A supplier that is a sentence or a link** ("Legal title of these goods remains in ...", "View online", "Limited") is rejected instead of naming the file.

## 5.5.1 — Outlook marking you can verify
* **Every Outlook change is checked by reading it back** from Outlook (a fresh copy of the e-mail), so a flag or category that Outlook accepts but does not keep is noticed instead of passing silently.
* The tick is tried three ways, the red flag two ways; a category that is not kept raises an error that says so.
* **Fallback:** if the tick cannot be set, the category `Processed by bot` is added (setting `Tick fallback category`, `none` switches it off) and the read status is kept.
* **New: Test Outlook marking on ONE e-mail** (panel Setup 1, or `run_bot.py --test-outlook-marking --mailbox jiya --subject BOTTEST`): tries read, tick (3 ways), red flag and category on one test e-mail, reads each back, shows Outlook's version and mailbox type, and puts the e-mail back as it was.

## 5.5.0 — back to company folders; the tick; names; reasons (from a real run)
* **Company folders again.** The shipped settings use `<root>\<MONTH> <YEAR> EMAIL INVOICES\<COMPANY>` (the layout of 5.1), not site folders. Site folders remain an option (Folder layout list, or `Destination pattern`). The site is always found and shown in the record.
  The automatic folder search now looks for `MAIL INVOICES` (setting `Auto-find folder name`).
* **The tick (flag complete) is set three ways and checked by reading it back** (`FlagStatus`, `MarkAsTask`, `FlagStatus` + completed date); the read status is always saved; if nothing works the run summary shows the exact reason ("first problem: ...") instead of only a count.
* **A failed or unrecognised document is still named `Supplier_Number`** when both were read (before: `date_originalname`). Documents where nothing could be read keep `dd-mm-yyyy_originalname`.
* **The run summary says why documents wait or are unrecognised** (counts per reason), so tuning does not need Excel.

## 5.4.3 — found by looking at a real inbox
* **Plural subjects are found**: "Sales Invoices for Sales Order" was ignored because "Invoices" is not the whole word "invoice". Keywords of 5+ letters now also match a plural (`Keywords also match plurals`, default Yes; short ones such as inv, cn, pod, stat, bill stay exact).
* **The sender can carry the company's folder name** (for example "Tweed Heads Energy No.1"), as well as its mailbox names.

## 5.4.2 — found by the batch tests (levels 5 to 7)
* **Shared towns never resolve by the generic name alone.** A document that says only "PEARL ENERGY WAGGA WAGGA" (no address) went to the site literally named that, although *Mobil Wagga Wagga* is in the same town and company.
  Now it waits for an address or the brand name. Nine towns in the list have this shape.
* **A killed run no longer blocks the next one.** The run lock only expired by age (up to 60 minutes after a crash or restart). It now records the process id and is taken over at once when that process is gone; a warning says so.
* A password-protected PDF now says so in the reason instead of an empty message.
* Measured with 2,500 generated documents (known answers, real site list, one core): 0 wrong values, 0 wrong sites, 83 s, 86 MB; a second day with 500 re-sent invoices: exactly 500 duplicates skipped, no growth in memory; killed mid-run and repeated: exactly one record per document.

## 5.4.1 — fixed: the parallel readers could freeze a busy run
* **A real hang was found and fixed.** With more than about 50 documents per reader, the readers were replaced through Python's `max_tasks_per_child` option, which deadlocks in Python 3.12: the run waited for ever.
  The program now replaces the readers itself, starts them with `spawn` on every system, and gives every file a time limit (240 s). A stuck or crashed reader is killed, the rest of the batch is read one by one, and the run continues (a warning says so; nothing is lost).
  Earlier tests used only a handful of documents (and the 2,500-document benchmark one reader), so they did not show it.
* **Fixed: a stale report.** A staff fix that left the file unrecognised did not rewrite the month's Record until another document arrived.
* Accuracy checked against known answers: 500 generated documents, 0 wrong values with one reader and with two, identical records.

## 5.4.0 — compressed backups, restore, quarterly records archive
* **Backups are compressed and verified** (`bot_YYYYMMDD.db.gz`, about 4× smaller; integrity check and full read-back before keeping). New `restore_backup.py` / `Advanced\Restore_backup.bat` (refuses while a run is in progress, keeps the replaced database).
* **Quarterly read-only records archive**: the first run after a calendar quarter writes `Archive_<first month>_to_<last month>.xlsx` (Record, Needs attention, Corrections, About). Nothing is deleted from the database.
  New settings `Quarterly records archive`, `Records archive folder`; command `run_bot.py --archive-quarter 2026-Q3`; panel button **Open → Quarterly archives**.
* **Reports stream from the database** instead of loading every row: a full quarter is written with about 50 MB of memory (was 485 MB).

## 5.3.1 — no Summary sheet; faster idle runs
* **Removed the Summary sheet** (KPI tiles, tables, charts) from `Record_YYYY-MM.xlsx` and `Report_<from>_<to>.xlsx`; they now hold **Record** and **Needs attention**. The setting `Minutes saved per document` is gone with it.
  (The Shadow report keeps its own short Summary: it is the headline of a rehearsal.)
* **Excel files are written only when something changed.** Before, the current month's report was rebuilt on every run, even an idle one (about 10 s at month-end for 55,000 rows).
  Measured: an idle run now takes 0.05 s, and 0.07 s with a month-end sized database. The run summary shows which Excel files were updated.

## 5.3.0 — company first, then site; the destination folder is found automatically
* **Two-step choice**: the bot decides the **company** first (document, sender, mailbox), then looks for the **site among that company's sites only**, then files into that site's folder.
  Two sites of one town that belong to different companies are no longer confused. A company named on the document is trusted; a weaker clue falls back to all sites and says so.
* **No path to type**: in LIVE mode the bot finds the synced `EMAIL INVOICES` folder (inside `OFFICE INDIA`, under the user profile / OneDrive) by itself, remembers it, and re-checks it at every run.
  A typed path still wins. New settings: `Find the destination folder automatically`, `Auto-find folder name`, `Auto-find parent folder`, `Choose the site within the company first`.

## 5.2.0 — site folders: the invoice goes to the folder of its site
* **Folder layout is a setting** (`Destination pattern`, or the Folder layout list in the panel): `{SITE}`, `{COMPANY}\{SITE}`, `{SITE}\{MONTH} {YEAR}`, `{MONTH_FOLDER}\{COMPANY}`.
  The workbook shipped with this version uses `{SITE}` = `<root>\<SITE FOLDER>`, the structure of `OFFICE INDIA\EMAIL INVOICES`.
* **Real folder names**: Sites sheet column *SharePoint folder*, filled by **Match my sites to the SharePoint folders** (`--match-folders`). It keeps `AMPOL  KELSO`'s double space, accepts a
  similar spelling (`Mobil Broken Hills` -> `MOBIL BROKEN HILL`) and reports sites without a folder and folders without a site.
* **Site finding**: a site also answers to "Pearl Energy <town>" when its town is unique (invoices and dockets are addressed that way). Towns with two sites are never matched by town alone.
* A document whose site cannot be decided waits in `_Unassigned`; `To_Fix.xlsx` has a **Site folder** drop-down with every real site folder. A missing site folder also waits (retried every run).
* New setting `Create site folder in live mode` (default No).

## 5.1.5 — rehearsal report with a proper Summary
* The Shadow report's Summary sheet now has KPI tiles (e-mails read, would be filed / parked / wait / unrecognised, red-flagged) and a per-mailbox table, in the same style as the monthly report.

## 5.1.4 — dates in file names are dd-mm-yyyy
* Every date written into a renamed file is now **dd-mm-yyyy** (`Supplier_12-08-2026_STM.pdf`, `03-10-2026_originalname.pdf` for unreadable files).
  The format is the Run-sheet cell `File name date format` (tokens dd, mm, yyyy, yy). Archive folders, backup files, logs and report file names are unchanged.

## 5.1.3 — the sender decides the company folder in a shared mailbox
* **Sender as a company clue**: an e-mail sent or forwarded by a company mailbox (*Pearl Energy Colyton*, *P & O Investment*, *Highland Energy 1*, *Shri Hari Investment* …)
  goes to that company's folder. Order: document → site → sender → mailbox (setting `Company source order`; switch off with `Use sender as company clue` = No).
  Workbooks from before this version get the sender clue automatically.
* **Fixed:** folder names were altered (`P & O INVESTMENT` became `P - O INVESTMENT`), so the real SharePoint folder would never have been found. Folder names are now used exactly as written;
  `&` is also kept in file names.

## 5.1.2 — three holding folders: _Unrecognised, _Unassigned and _Statements
* Statements now have **their own holding folder `Holding\_Statements`** next to `_Unassigned` and `_Unrecognised` (version 5.1.1 had filed them with the invoices; reversed on request).
  The three folders are created at the first run so they are visible. The folder name is the `Subfolder` cell of the Statement row in the DocTypes sheet, and `Route` = `Holding` /
  `Unrecognised` / `Company folder` chooses where each document type goes.
* Kept from 5.1.1: a statement is only a duplicate when its content is identical (its "number" is just a line item).

## 5.1.1 — statements filed like invoices; safe testing inside a mailbox that holds real e-mails
* **Statements now go to the same path as invoices, credit notes and PODs** (`<root>\<MONTH> <YEAR> EMAIL INVOICES\<COMPANY>`). Account summaries, activity reports and hire statements are still parked; each type's route is one cell in the DocTypes sheet.
* A statement is no longer compared by its "number" for duplicates (that number is just a line item), only by content.
* **Test filter**: `--subject TEXT` / the panel box "Only touch e-mails whose subject contains …". Every other e-mail is left completely alone
  and the catch-up window is not advanced, so a later normal run still handles them.

## 5.1.0 — one-click control panel and professional Excel reports
**New**
* **Control panel** (`Start_Bot.bat`): a local web page with big buttons — Run now, Rehearsal, Run for dates, Check my PC, demo, test with saved `.msg` e-mails, Stop,
  live log and result cards, "needs your attention", shortcuts to folders and files, and a Setup section: **find the Outlook mailboxes and match them to companies**,
  **switch TEST/LIVE with the SharePoint folder**, **turn automatic running on/off**.
* `Start_Bot.bat` installs the Python packages by itself on the first start; helper scripts moved into `Advanced\`.
* **Excel reports redesigned**: monthly report with a Summary page (KPI tiles, tables, native charts, estimated staff time saved), a filterable colour-coded Record with clickable links
  and a Needs-attention sheet; `To_Fix.xlsx` with a how-to sheet, drop-down lists and links; report for any period (`--report` / panel button); nicer shadow report and settings workbook.
* Mailbox detection (`invoicebot/mailboxes.py`) and a safe settings editor with a rolling backup.

**Fixed**
* PFD documents in a layout its special reader does not know now fall back to the general reader.
* A Companies row with an empty "Text to look for" could be picked for unrelated company names.
* Unclosed workbook/pipe handles (would keep `To_Fix.xlsx` locked on Windows / leak a handle per run).

## 5.0.0 — complete rebuild
**New**
* Month → company destination: `<root>\<MONTH> <YEAR> EMAIL INVOICES\<COMPANY FOLDER>`, built automatically; one place for the real SharePoint path (`config.py` or sheet *Run*).
* Test mode / live mode; live mode never invents SharePoint folders unless allowed; missing folders → `_Unassigned`, retried every run.
* All 13 mailboxes in one run, with automatic date-window catch-up, explicit From/To dates, overlap and caps.
* Email rules: blocked words (flag, never download), keywords, exception senders, image-only, FW/RE categories; existing flags are never overwritten.
* Document types with their own routing and names: invoice, credit note (negative), POD, statement, account summary, customer activity report, hire statement, other.
* SQLite database as the primary record; Excel monthly records, `To_Fix.xlsx` correction loop, From/To report.
* Parallel reader processes with recycling, memory watchdog, run time limit, page and size limits.
* **Shadow mode** (full rehearsal, changes nothing) with a detailed report.
* Demo mode with a simulated Outlook; automated tests; benchmark; scorecard; diagnose; setup checker.
* OCR for scanned PDFs and supplier logos; Excel invoices.
* Testing with saved e-mails: `.msg` reader, `run_bot.py --msg-folder`, `.msg` support in `scorecard.py` and `diagnose.py`.
* Site list importer (157 sites, renamed-site aliases, spelling variants, address fragments).

**Fixed (found with the real e-mails and PDFs)**
* Wrong supplier names ("Weekly Statement", "Follow-Up Report", "Australia", "Total inc GST"): alias table first, junk-title rejection, sender hints, logo OCR.
* Exception sender "Rarity Wholesale" now matches `raritywholesale.com.au` (spaces ignored).
* File name `Account_9014462` no longer mistaken for an invoice number.
* A 20 MB catalogue: 48 s and 1 GB → 1.1 s (page-count check before reading, logo OCR only on logo-sized images).
* `NOTE - CREDIT` titles are credit notes.
* Failed downloads keep the date window open; one failed attachment means the e-mail is not ticked.
* A damaged or locked settings file stops the run instead of silently using defaults.
