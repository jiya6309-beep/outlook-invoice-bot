"""Bot_Settings.xlsx: every business rule that staff may need to change, in Excel.

The program never needs editing for day-to-day changes. Each sheet is a small table:

    Run          options (paths, limits, switches)
    Keywords     subject words that mean "this email carries a document"  (inv, cn, pod, ...)
    BlockedWords subject words that mean "do NOT download, flag red"     (overdue, urgent, ...)
    DocTypes     what each document type means and where it goes
    Senders      exception senders (always download) and sender -> supplier hints
    Companies    the 13 companies: mailbox, folder name, text to look for on documents
    Suppliers    spelling clean-up for supplier names (and the optional MILK INVOICES folder)
    Sites        the site list (used for the record and the handler e-mails)
    Accounts     supplier account numbers -> site
    Labels       the words each supplier prints next to invoice number, date, GST, total

Columns are found by their HEADER TEXT, so adding or re-ordering columns is safe.
"""
import os
import re

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.properties import PageSetupProperties

from . import paths

UNASSIGNED = "UNASSIGNED"


class SettingsError(Exception):
    """Bot_Settings.xlsx cannot be used (locked, damaged, or a value is not allowed). The run stops with the message."""

SETTINGS_FILE_OVERRIDE = None          # tests may point this somewhere else

# ----------------------------------------------------------------------------------------
# Default content
# ----------------------------------------------------------------------------------------
LIVE_ROOT = r"C:\SHAREPOINT\pearlenergy.com.au\Patel Group - Documents\OFFICE INDIA\EMAIL INVOICES"   # 5.15.0: the real LIVE folder
LIVE_HOLDING = "HOLDINGS"          # LIVE: <Live destination root>\HOLDINGS when 'Holding root' is blank
LIVE_REPORTS = "Reports"           # LIVE: <Live destination root>\Reports when 'Reports root' is blank

RUN_DEFAULTS = [
    # key, value, notes
    ("Destination mode", "test", "test = write into the TestOutput folder next to the program.  live = write to the real SharePoint/NAS root below."),
    ("Find the destination folder automatically", "Yes", "Yes = when 'Live destination root' is empty (or does not exist on this PC), the bot looks for the synced SharePoint folder by itself: a folder 'Auto-find folder name' inside 'Auto-find parent folder', in C:\\SHAREPOINT and under your user profile / OneDrive."),
    ("Auto-find folder name", "EMAIL INVOICES", "The name of the destination folder to look for."),
    ("Auto-find parent folder", "OFFICE INDIA", "The folder it must be inside. Empty = any folder."),
    ("Test destination root", "", "Blank = the folder  TestOutput  next to the program."),
    ("Live destination root", LIVE_ROOT, "LIVE: the EMAIL INVOICES folder. The month folders (OCTOBER 2026 EMAIL INVOICES\\<COMPANY>\\<SITE>), HOLDINGS and Reports are made inside it when missing; an existing folder is always used, never made twice. This folder itself must exist (SharePoint synced): if it is missing the run stops and nothing is written. Change it here if the folder moves. Empty = the bot looks for it by itself."),
    ("Holding root", "", "Where _Statements, _Order Confirmations, _Direct Debit Notices, _Account Activity Listings, _Unassigned and _Unrecognised live. Blank (normal) = LIVE: <Live destination root>\\HOLDINGS (made when missing), TEST: the folder Holding next to the program. Type a path only to put them somewhere else; that folder (or the folder it sits in) must exist."),
    ("Reports root", "", "Where To_Fix.xlsx, the monthly Record, the A/C Excel and the quarterly archive are written. Blank (normal) = LIVE: <Live destination root>\\Reports (made when missing), TEST: the folder Reports next to the program. The database, logs and Scratch ALWAYS stay next to the program (a database in a synced folder can be corrupted)."),
    ("Month folder pattern", "{MONTH} {YEAR} EMAIL INVOICES", "Tokens: {MONTH}=OCTOBER  {Month}=October  {YEAR}=2026  {MM}=10.  Built automatically every month."),
    ("File name date format", "dd-mm-yyyy", "The date inside renamed files (statements, summaries, unreadable files). Tokens: dd mm yyyy yy.  dd-mm-yyyy gives 04-10-2026.  A slash / cannot be used in a file name."),
    ("Destination pattern", "{MONTH_FOLDER}\\{COMPANY}", "The folder layout under the destination root. Tokens: {SITE} = the site's SharePoint folder, {COMPANY} = the company folder, {MONTH_FOLDER} = e.g. OCTOBER 2026 EMAIL INVOICES, {MONTH} {YEAR} {MM}. Examples:  {SITE}   |   {COMPANY}\\{SITE}   |   {SITE}\\{MONTH} {YEAR}   |   {MONTH_FOLDER}\\{COMPANY}"),
    ("Choose the site within the company first", "Yes", "Yes = first decide the COMPANY (document, then sender, then mailbox), then look for the site among that company's sites only. Two sites in one town that belong to different companies are then never confused."),
    ("File in the company folder when the site is not found", "Yes", "Yes = when the layout is month > company > site and the site cannot be decided, the document is filed in the COMPANY folder (one level up) and the Remarks column says SITE NOT FOUND. No = it waits in _Unassigned until a person chooses the site."),
    ("Live mode creates missing folders like test mode", "Yes", "Yes (from 5.15.0) = LIVE creates every missing folder exactly like TEST: month, company and site (an existing folder is always used). No = only the folders switched on below are created."),
    ("Create site folder in live mode", "Yes", "Yes = the bot creates a missing site folder. No = a file whose site folder does not exist waits in _Unassigned (only used when 'Live mode creates missing folders like test mode' = No)."),
    ("Hold documents without amounts when the type came only from the subject", "Yes", "Yes = an invoice with none of the three amounts (total incl. GST, GST, amount excl. GST) waits when only the e-mail subject says it is an invoice; if the page itself says TAX INVOICE it is filed. A catalogue or an order confirmation cannot slip into a company folder."),
    ("Keywords also match plurals", "Yes", "Yes = a keyword of 5+ letters also matches its plural: invoice finds 'Sales Invoices', credit note finds 'Credit Notes', statement finds 'Statements'. Short keywords (inv, cn, pod, stat, bill) stay exact."),
    ("Tick fallback category", "Processed by bot", "If the tick (flag complete) cannot be set on your Outlook, this category is added to the e-mail instead, so staff can still see which e-mails the bot handled. Type none to switch it off."),
    ("Month basis", "arrival", "Which month's Excel record a document is listed in. arrival = the month the e-mail arrived (a 30 Sept e-mail processed on 1 Oct is listed in SEPTEMBER).  invoice = the date printed on the document. (Where the FILE goes is set by 'Month folder follows'.)"),
    ("Month folder follows", "window", "window = the month folder follows the INVOICE date, within 'Month window' months of the month the e-mail arrived: an older invoice goes to its own (older) month folder, a future-dated one waits in the current month folder and is moved when its month starts. Anything outside the window goes to the current month. arrival = always the month the e-mail arrived. invoice = always the invoice's own month."),
    ("Create next month's folders", "Yes", "Yes = in the last two days of a month the bot copies the folder structure (company and site folders, folders only) of the latest month to the next month's folder, and on the first run of a new month it builds the folder if it is missing. Staff then do not create month folders. Future-dated invoices that waited in the current month are moved into their month when its folder exists."),
    ("Month window (months)", "3", "How many calendar months back (and ahead) an invoice date may be and still get its own month folder."),
    ("Create month folder in live mode", "Yes", "Yes = the bot creates a month folder only if it does not exist yet (an existing one is always used). No = hold the files in _Unassigned until staff create it."),
    ("Create company folder in live mode", "Yes", "Yes = the bot creates a company folder only if it does not exist yet. No = hold the files if the company folder is missing."),
    ("Company source order", "document,site,sender,mailbox", "How the company folder is chosen, first match wins. document = a company name printed on the document; site = the company that owns the matched site; sender = the company mailbox the e-mail was sent/forwarded FROM; mailbox = the mailbox it arrived in."),
    ("Use sender as company clue", "Yes", "Yes = an e-mail sent or forwarded by a company mailbox (for example 'Pearl Energy Colyton') goes to that company's folder. Needed when one shared mailbox receives the invoices of all companies."),
    ("From date", "", "Optional. YYYY-MM-DD. Leave blank for automatic catch-up since the last successful run."),
    ("To date", "", "Optional. YYYY-MM-DD (inclusive)."),
    ("Default catch-up days", "2", "Used on the very first run of a mailbox."),
    ("Catch-up max days", "31", "A catch-up never goes back further than this. Use From/To dates for older mail."),
    ("Overlap minutes", "60", "Each run starts this much before the end of the previous successful run."),
    ("Max emails per mailbox per run", "5000", "Safety limit."),
    ("Max attachments per email", "20", "Safety limit."),
    ("Max run minutes", "45", "The run stops cleanly after this long; the rest is picked up next run."),
    ("Batch size (emails)", "25", "E-mails downloaded, read and filed per batch."),
    ("Workers", "auto", "Reader processes. auto = CPU cores minus 1 (maximum 4)."),
    ("Recycle worker after documents", "50", "The reader processes are replaced by fresh ones after this many documents each (keeps memory flat). Done by the program itself."),
    ("Min free RAM MB (use 1 worker)", "1500", "Below this much free memory the bot reads one document at a time."),
    ("Min free RAM MB (stop run)", "500", "Below this the run stops cleanly and resumes next time."),
    ("Max pages (invoice types)", "15", "Longer documents are parked as unrecognised (protects memory)."),
    ("Max attachment MB", "25", "Bigger files are not read; they go to _Unrecognised."),
    ("OCR enabled", "Yes", "Reads scanned PDFs and logos. Needs Tesseract installed (see README)."),
    ("OCR max pages", "3", "Pages of a scanned PDF that are read."),
    ("Tesseract path", "", "Blank = auto-detect (C:\\Program Files\\Tesseract-OCR\\tesseract.exe)."),
    ("Red category name", "Forwarded / Reply", "Outlook category given to FW:/RE: e-mails."),
    ("Check attachment names for order confirmations", "Yes", "Yes = an e-mail whose subject has no document word is still looked at when an attachment's FILE NAME contains 'order' (Order-SO106649.pdf). The PDF is opened and filed only if the page itself is an Order Confirmation; otherwise it is ignored (the e-mail is noted so it is not opened again). No = only subjects and senders decide."),
    ("Skip e-mails already ticked in Outlook", "Yes", "Yes = an e-mail that already has the tick (flag complete) or the 'Processed by bot' category is NOT read again, even when this copy of the bot has no record of it (a new install, a new PC, a deleted data folder). No = read every e-mail in the period."),
    ("Skip e-mails already flagged in Outlook", "Yes", "Yes = an e-mail that already has a red flag (set by the bot or by a person) is NOT read again. No = flagged e-mails are read like any other."),
    ("Learn sites from staff fixes", "Yes", "Yes = when a person chooses the site of a waiting ship-to document in To_Fix.xlsx, the bot remembers that supplier's deliver-to number / address (and a Fleet Card merchant number) for that site, so the next document goes there by itself. The learned list is shown in To_Fix.xlsx (sheet Learned sites) and can be removed there."),
    ("Look for the site in the SharePoint folders", "Yes", "Yes = a ship-to address that is not in the Sites sheet is also compared with the site FOLDERS that exist in the month folder (month > company > site). When staff create the folder (for example PEARL ENERGY PRESTON inside its company folder), the waiting documents are filed there by itself."),
    ("Use only the ship-to address", "Yes", "Yes = when a document has a ship-to / deliver-to address, ONLY that address decides the site and the company: the bill-to (invoice to) address and account numbers outside the ship-to block are never used. A ship-to address that is not in the Sites sheet makes the document WAIT in _Unassigned (reason in To_Fix.xlsx); add the site to the Sites sheet and the next run files it by itself. No = an unknown ship-to falls back to the other clues (the bill-to address), as before 5.11.0."),
    ("Ship-to address wins over bill-to", "Yes", "Yes = when a document shows both a bill-to (invoice to) and a ship-to / deliver-to address, the SHIP-TO address decides the site AND the company. No = the older rule (both addresses count; two different sites = a person decides)."),
    ("Download attachments of FW/RE e-mails", "Yes", "Yes = the PDF / Excel attachments of FW: and RE: e-mails are downloaded and filed even when the subject has no document word. The e-mail is then left UNREAD with a red flag and NO tick, so a person can look at it. No = the old behaviour (red flag only, nothing downloaded)."),
    ("Image category name", "Image attachment", "Outlook category for e-mails that only carry images."),
    ("Red category applies to", "document", "document = FW/RE e-mails that look like documents (keyword/exception sender/blocked word). all = every FW/RE e-mail."),
    ("Image min KB", "40", "Images smaller than this (logos, icons) are ignored."),
    ("Flag blocked emails without attachment", "Yes", "Red-flag overdue/urgent e-mails even when they have no attachment."),
    ("Scan body for blocked words", "No", "Also look for blocked words in the first lines of the e-mail body of document e-mails."),
    ("Hold first-time suppliers", "No", "Yes = a supplier name never seen before waits in _Unassigned until a person confirms it."),
    ("Email path mode", "off", "off | draft | send. Tells each site handler (Sites sheet) where new files were saved."),
    ("Archive folder", "", "Blank = the folder  Archive  next to the program."),
    ("Archive days", "7", "Copies of the original files are kept this long on the bot PC. 0 = keep none."),
    ("Report date basis", "invoice", "invoice | arrival. Used by the From/To report."),
    ("Backup folder", "", "Blank = data\\backups next to the program. Point it at the NAS for a safe copy."),
    ("Quarterly records archive", "Yes", "Yes = every three months the records of the quarter that just ended are written to a read-only Excel file. Nothing is deleted from the database."),
    ("Records archive folder", "", "Blank = Reports\\Archive next to the program. Point it at the NAS or a SharePoint folder to keep the archives off the bot PC."),
    ("Backup keep days", "14", "Daily database backups kept."),
    ("Max attempts for waiting files", "50", "A file waiting in _Unassigned is retried every run, up to this many times."),
    ("Settings version", "5.20.0", "Written by the bot. Tells it which one-time upgrades this file already has. Do not change."),
]

KEYWORDS = [
    ("invoice", "Invoice"), ("tax invoice", "Invoice"), ("invoiceno", "Invoice"), ("inv", "Invoice"),
    ("bill", "Invoice"), ("tba", "Invoice"), ("service", "Invoice"),
    ("adj", "Credit note"), ("credit note", "Credit note"), ("sales credit", "Credit note"),
    ("credit", "Credit note"), ("cn", "Credit note"), ("adjustment note", "Credit note"),
    ("credit memo", "Credit note"), ("pocredit", "Credit note"), ("total credit", "Credit note"),
    ("note credit", "Credit note"),
    ("stm", "Statement"), ("statement", "Statement"), ("stat", "Statement"),
    ("nutjob confectionery", "Statement"),
    ("hire statement", "Hire Statement"),
    ("order confirmation", "Order Confirmation"), ("order confirmed", "Order Confirmation"),
    ("order acknowledgement", "Order Confirmation"), ("order acknowledgment", "Order Confirmation"),
    ("customer activity report", "Customer Activity Report"),
    ("account summary", "Account Summary"),
    ("pod", "POD"), ("delivery docket", "POD"), ("proof of delivery", "POD"),
    ("direct debit notice", "Direct Debit Notice"), ("order receipt", "Other"), ("rma completed", "Other"),
    ("fleetcard merchant payment advice", "Fleet Card Statement"), ("merchant payment advice", "Fleet Card Statement"),
    ("account activity listing", "Account Activity Listing"), ("follow up report", "Other"),
]

BLOCKED = ["overdue", "urgent", "reminder", "remainder", "stop supply", "warning"]

DOCTYPES = [
    # type, route, subfolder, read amounts, sign, required fields, tag, name pattern, folder layout, notes
    ("Invoice", "Company folder", "", "Yes", "+", "number,date,total,gst", "INV", "{supplier}_{number}", "", "Filed to the company folder inside the month folder."),
    ("Credit note", "Company folder", "", "Yes", "-", "number,date,total,gst", "CN", "{supplier}_{number}_CN", "", "Amounts are stored negative."),
    ("POD", "Company folder", "", "No", "", "number", "POD", "{supplier}_{number}_POD", "", "Proof of delivery: no amounts."),
    ("Statement", "Holding", "_Statements", "No", "", "date", "STM", "{supplier}_{site}_{date}_STM", "{MONTH}\\{COMPANY}", "Kept in its own holding folder (HOLDINGS\\_Statements\\<month>\\<COMPANY>). Named supplier_site_date_STM; NoSite when no site was found (flagged in To_Fix)."),
    ("Order Confirmation", "Holding", "_Order Confirmations", "No", "", "date", "OC", "{supplier}_{site}_{date}_OC", "{MONTH}", "Treated like a statement: Holding\\_Order Confirmations\\<month>. Named supplier_site_date_OC; NoSite when no site was found (flagged in To_Fix)."),
    ("Customer Activity Report", "Company folder", "", "No", "", "date", "CAR", "{supplier}_{date}_CAR", "", "5.18.0: filed in the site folder like an invoice (Metcash Customer Activity Report, also when its PDF says Statement). No amounts; not in the A/C Excel."),
    ("Hire Statement", "Unrecognised", "Hire Statement", "No", "", "", "HIRE", "{supplier}_{date}_HIRE", "", "Parked on purpose."),
    ("Account Summary", "Unrecognised", "Account Summary", "No", "", "", "ACCSUM", "{supplier}_{date}_ACCSUM", "", "Parked on purpose."),
    ("Direct Debit Notice", "Holding", "_Direct Debit Notices", "No", "", "date", "DDN", "{supplier}_{date}_DDN", "{COMPANY}", "Direct debit notices (Tip Top ...): HOLDINGS\\_Direct Debit Notices\\<COMPANY>. Named supplier_date_DDN."),
    ("Account Activity Listing", "Holding", "_Account Activity Listings", "No", "", "date", "AAL", "{supplier}_{date}_AAL", "{SITE}", "Campbells / C-Store account activity listings: HOLDINGS\\_Account Activity Listings\\<SITE>. Named supplier_date_AAL."),
    ("Fuel Invoice", "Mailbox folder", "", "No", "", "number", "FUEL", "{number}_{site}", "{COMPANY}", "Fuel invoices from Pearl Investment Family Trust: <Destination folder of the Fuel invoices mailbox>\\<COMPANY>\\<INVOICE NUMBER>_<SITE>.pdf, whichever mailbox it came to (5.19.0; company and site from the invoice's Customer, never from the Reference)."),
    ("Fleet Card Statement", "Mailbox folder", "", "No", "", "date", "STM", "{date}_STM", "{SITE}", "Fleet Card Merchant Payment Advice: <Destination folder of the Fleet card mailbox>\\<SITE>\\<report issue date>_STM, whichever mailbox it came to (5.19.0). The same statement from two mailboxes is saved once."),
    ("Other", "Unrecognised", "", "No", "", "", "OTH", "{supplier}_{date}_OTH", "", "Direct debit notices, order receipts, ... left for a person."),
]

SENDERS = [
    # sender contains, always download, default type, supplier name, notes
    ("Rarity Wholesale", "Yes", "", "Rarity Wholesale", "Exception sender"),
    ("Best Burning Fire Fuels", "Yes", "", "Best Burning Fire Fuels", "Exception sender"),
    ("Sweet and Sour Australia", "Yes", "", "Sweet and Sour Australia Pty Ltd", "Exception sender"),
    ("Moorees", "Yes", "", "Moorees", "Exception sender - writes the site name in the subject"),
    ("Tip Top", "Yes", "", "Tip Top", "Exception sender - subject like DT:06.09.2026 C:1202316 INV:8024871400"),
    ("ccamatil", "No", "", "Coca-Cola Europacific Partners Australia Pty Ltd", "Hint only: names the supplier when the logo is an image"),
    ("jlborgo", "No", "", "Borgo's", "Hint only"),
    ("mobigodistributors", "No", "", "MobiGo Distributors", "Hint only"),
]

COMPANIES = [
    # name, text to look for on documents, folder name, mailbox (as in Outlook; | separates alternatives)
    ("PEARL ENERGY NO: 1 PTY LTD", "Pearl Energy No 1", "ONE PEARL ENERGY", "One@pearlenergy.com.au|One Pearl Energy"),
    ("PEARL ENERGY NO: 2 PTY LTD", "Pearl Energy No 2", "TWO PEARL ENERGY", "Two@pearlenergy.com.au|Two Pearl Energy"),
    ("PEARL ENERGY NO: 3 PTY LTD", "Pearl Energy No 3", "THREE PEARL ENERGY", "Three@pearlenergy.com.au|Three Pearl Energy"),
    ("PEARL ENERGY NO: 5 PTY LTD", "Pearl Energy No 5", "FIVE PEARL ENERGY", "five@pearlenergy.com.au|Five Pearl Energy"),
    ("SHRI HARI INVESTMENT PTY LTD", "Shri Hari Investment", "SHRI HARI INVESTMENT", "shrihari@pearlenergy.com.au|Shri Hari Investment"),
    ("GOPINATH INVESTMENT PTY LTD", "Gopinath Investment", "GOPINATH INVESTMENT", "Gopinath@pearlenergy.com.au|Gopinath Investment"),
    ("CENTRAL COAST ENERGY PTY LTD", "Central Coast Energy", "CENTRAL COAST ENERGY 1", "centralcoastenergy@pearlenergy.com.au|Central Coast Energy"),
    ("SHREE UMIYA INVESTMENT PTY LTD", "Shree Umiya Investment", "SHREE UMIYA INVESTMENT", "shreeumiya@pearlenergy.com.au|Shree Umiya Investment"),
    ("P & OINVESTMENT PTY LTD", "P & O Investment", "P & O INVESTMENT", "poinvestment@pearlenergy.com.au|P & O Investment"),
    ("COLYTON FUELS PTY LTD", "Colyton Fuels", "COLYTON FUELS", "pearl.colyton@pearlenergy.com.au|Pearl Energy Colyton"),
    ("RADHE KRISHNA INVESTMENT PTY LTD", "Radhe Krishna Investment", "RADHE KRISHNA INVESTMENT", "radhekrishna@pearlenergy.com.au|Radhe Krishna Investment"),
    ("HIGHLAND ENERGY 1 PTY LTD", "Highland Energy 1", "HIGHLAND ENERGY 1", "highlandenergy@pearlenergy.com.au|Highland Energy"),
    ("Tweed Heads Energy Pty Ltd", "Tweed Heads Energy", "TWEED HEADS ENERGY NO.1", "pearl.tweedheads@pearlenergy.com.au|Pearl Energy Tweed Heads"),
]

SUPPLIERS = [
    # name contains (| = alternatives), record name, short name, folder override
    ("PFD", "PFD Food Services Pty Ltd", "PFD", ""),
    ("Bhavani", "Bhavani Distributors", "Bhavani", ""),
    ("Elgas", "Elgas Limited", "Elgas", ""),
    ("Suntory", "Suntory Beverage & Food Australia Pty Ltd", "Suntory", ""),
    ("Bega|BDD Australia", "Bega Cheese Limited", "Bega", ""),
    ("Asahi", "Asahi Beverages", "Asahi", ""),
    ("Fresh To Go", "Fresh To Go Foods Pty Ltd", "FreshToGo", ""),
    ("The Distributors", "The Distributors", "TheDistributors", ""),
    ("Coca-Cola|Europacific", "Coca-Cola Europacific Partners Australia Pty Ltd", "CocaCola", ""),
    ("Rarity Wholesale", "Rarity Wholesale", "Rarity", ""),
    ("Borgo", "Borgo's", "Borgos", ""),
    ("MobiGo", "MobiGo Distributors", "MobiGo", ""),
    ("Lactalis", "Lactalis Australia Pty Ltd", "Lactalis", ""),
    # suppliers with several sites (one ABN each): ONE name for all of them, whatever site or ABN the document shows
    ("Peters", "Peters", "Peters", ""),
    ("Peleguy", "Peleguy", "Peleguy", ""),
    ("Nippy", "Nippy's", "Nippys", ""),
]

ADDED_SUPPLIERS = SUPPLIERS[-3:]            # appended to an older Bot_Settings.xlsx at the next start

LEGAL_ENDINGS = ["Pty", "Pvt", "Private", "Proprietary", "Limited", "Ltd", "LLP", "LLC", "Inc", "Incorporated", "Corp",
                 "Corporation", "P/L", "PLC"]

SUPPLIER_IDS = [
    # supplier (record name), ABN, e-mail domain, text after "Account Name" - ANY ONE of them found on the page is enough
    ("JP Bakery", "27065303264", "jpbakery.com.au", ""),
    ("Bonfect Snackfood Services", "30648891829", "bonfect.com.au", "Bonfect Snackfood Services"),
    ("Statewide Oil", "43009283363", "statewideoil.com.au", "Statewide Oil"),
    ("KC Supplies", "85105637901", "kcsupplies.com.au", "KC Supplies"),
    ("Tip Top Bakeries", "45008429632", "", "Tip Top Bakeries"),                # "A Division of George Weston Foods": the name is Tip Top
    # --- the company's supplier list (80 rows: ABN, e-mail domain). One name can have several rows (one ABN each). Free-mail and e-mail-less rows are ABN only.
    ('Essentials Intl', '18660934396', 'essentialsint.com', ""),
    ('SRC', '59153320800', '', ""),
    ("Windybank's Bait", '12818906218', 'windybanks.com.au', ""),
    ('Supagas', '50074008496', 'supagas.com.au', ""),
    ('IPL', '28632453135', 'iplretail.com.au', ""),
    ('Lactalis', '56072928879', 'lactalis.com', ""),
    ('Unilever', '66004050828', 'unilever.com', ""),
    ('Sodastream', '81104202093', 'pepsico.com', ""),
    ("Nippy's", '26007841492', 'nippys.com.au', ""),
    ('Asahi', '51004243994', 'asahi.com.au', ""),
    ('Youfoodz', '72160909893', 'youfoodz.com', ""),
    ('Coke', '68076594119', 'ccep.com', ""),
    ('The Dist', '87002994774', 'the-distributors.com.au', ""),
    ('Bells Pure Ice', '75653679040', 'bgl.com.au', ""),
    ('Ezymart', '84608615961', 'ezymartdistribution.com.au', ""),
    ('Mobigo', '87169618660', 'mobigodistributors.com.au', ""),
    ('Tasty Track', '82688861001', 'tastytrackdistributors.com.au', ""),
    ('Sugar Shock', '17680481690', 'mathtechindustries.com', ""),
    ('Direct Bakery', '26181647750', '', ""),
    ('Quinnies Coffee', '64456811098', 'quinniescoffeeco.com.au', ""),
    ('AO Global', '42089269778', 'aoglobal.net.au', ""),
    ('Peters', '11154314913', 'peters.com.au', ""),
    ('PFD', '29006972381', 'pfdfoods.com.au', ""),
    ('Suntory Oceania', '73060091536', 'suntory.com', ""),
    ('P N & J H', '23001246868', 'bestburningfirefuels.com.au', ""),
    ('Melitta', '18002133726', 'melitta.com', ""),
    ('Packedging', '34684579637', '', ""),
    ('Bega', '68083019390', 'bega.com.au', ""),
    ('Tweed Bait', '71010917674', 'tweedbait.com.au', ""),
    ('Bidfood', '33000228231', 'bidfood.com.au', ""),
    ('Red Funnel', '21000633689', 'redfunnel.com.au', ""),
    ('The Real Dist', '56775317664', 'realdist.com.au', ""),
    ('The Market Grocer', '18115674774', 'themarketgrocer.com', ""),
    ('Neverfail', '43003559519', 'neverfail.com.au', ""),
    ('Peleguy', '81606668713', 'peleguy.com.au', ""),
    ("Thompson's Bakery", '58659320171', 'thompsonspies.com', ""),
    ('Lidcombe Ice', '42093371969', 'lidcombeice.com.au', ""),
    ('Nutjob', '97201145378', '', ""),
    ('The Green Centre', '55659690856', 'thegreencentre.com.au', ""),
    ('BOC', '95000029729', 'boc.com', ""),
    ('Gold Coast Cookies', '66641150254', '', ""),
    ('Peleguy', '98832539665', 'peleguy.com.au', ""),
    ('Smart Gadget World', '21573313001', '', ""),
    ('C-Store', '62080064099', '', ""),
    ('Elgas Swap N Go', '85002749260', 'elgas.com.au', ""),
    ('Enviro Chemicals', '49053846361', 'envirochemicals.com.au', ""),
    ('Simply Fresh', '71168312409', 'simplyfreshdist.com.au', ""),
    ('Biosafe Pack', '19681063641', 'biosafepack.com.au', ""),
    ('Spiced Tea', '18630456067', 'spicedteachai.com', ""),
    ('Sweet N Sour', '15133285711', 'sweetandsour.com.au', ""),
    ("Sugar D's", '58466013325', '', ""),
    ('Charcoal Kings', '35623207885', 'standardandpartnersltd.com', ""),
    ('Rarity', '15100064775', 'raritywholesale.com.au', ""),
    ('Coffs Cleaner', '12143117024', 'coffscleaningsupplies.com.au', ""),
    ('Peters', '22000802040', 'peters.com.au', ""),
    ('Cool Blue Ice', '82628579671', 'coolblueice.com.au', ""),
    ('Menval', '72151153730', 'sweetground.com.au', ""),
    ("Garlo's Pies", '30095033971', 'garlospies.com.au', ""),
    ('Slick Chicks', '17078782537', 'slickchicks.com.au', ""),
    ('DJays', '14103055296', 'djaysgourmet.com.au', ""),
    ('Shayona Dist', '74674660087', 'shayonadistributors.com.au', ""),
    ('More Ice', '84057938708', 'moreice.com.au', ""),
    ('Veridia', '96662613407', 'veridia.com.au', ""),
    ("Nippy's", '39645758336', 'nippys.com.au', ""),
    ("Walkley's Finest", '50644404848', 'walkleysfinest.com.au', ""),
    ('Ribo Firewood', '37115064401', 'ribo.com.au', ""),
    ('Farmer Bake', '57651670585', 'farmersbakeplace.com.au', ""),
    ("Jim's Jerky", '54659051233', 'jimsjerky.com.au', ""),
    ('YG Dist', '95674663864', 'ygdistributors.com.au', ""),
    ('Paloma Foods', '76680049889', 'nutsandsweets.com.au', ""),
    ('Send It', '98652223535', 'senditenergy.com.au', ""),
    ('Tandoor & Tadka', '69644700514', '', ""),
    ('J&L Borgo', '67000492071', 'jlborgo.com.au', ""),
    ('Aussie Ice', '11167816079', 'aussieice.com.au', ""),
    ('Repco', '97097993283', 'gpcasiapac.com', ""),
    ('The Cafe Dist', '31611310000', 'thecafedist.com.au', ""),
    ('Coast2Coast', '54161162830', 'coast2coastdistributors.com.au', ""),
    ("Business Fuel Cards", "99161721106", "fleetcard.com.au", ""),
    ("Campbells", "57000226399", "", ""),
    ("Rum City Foods", "38509899256", "rumcityfoods.com.au", ""),
]

SITE_ALIASES = [
    # other ways a document names a site | the site as written in the Sites sheet
    ("Pearl Energy Wyong|10 Amy Close|1-10 Amy Close|Amy Close Wyong", "Pearl Energy North Wyong"),
]

ADDED_KEYWORDS = [("del docket", "POD"), ("docket", "POD"), ("delivery note", "POD"), ("goods received", "POD"),
                  ("goods received note", "POD"), ("delivery receipt", "POD"),
                  ("order confirmation", "Order Confirmation"), ("order confirmed", "Order Confirmation"),
                  ("order acknowledgement", "Order Confirmation"), ("order acknowledgment", "Order Confirmation")]
# (not 'sales order' as a keyword: subjects like 'Sales Invoices for Sales Order' are invoices - see doctypes.ORDER_TITLE)
MISSING_SITE = "STATEMENT MISSING DATA: site name not found in this document"
ADDED_LABELS = {"invoice_no": ["crm id"], "date": ["order date", "received"], "gst": ["tax"],         # always last: every other label wins first
                "total": ["total incl gst", "total inclusive gst", "total including gst", "total incl taxes", "total including taxes",
                          "total inc taxes"]}
# 5.15.0: the order of the TOTAL words of the '*' row. The amount payable (inc GST) always wins over any other total; a supplier's
# own Labels row still comes first. Words a person added that are not in this list keep their place after these; 'total' is last.
PAYABLE_ORDER = ["total payable", "total amount payable", "amount payable",
                 "total inc gst", "total incl gst", "total inclusive gst", "total including gst", "total incl taxes",
                 "total including taxes", "total inc taxes",
                 "total amount due", "total due", "amount due",
                 "invoice total", "total invoice", "grand total", "total value", "total price", "total credit", "credit total", "total aud",
                 "balance due"]           # last before 'total': a paid invoice shows 'Balance due 0.00'


def payable_sorted(words_list):
    """The total words in PAYABLE_ORDER first (in that order), then any other words in their own order, the bare 'total' last."""
    known = {w: i for i, w in enumerate(PAYABLE_ORDER)}
    first = sorted([w for w in words_list if w.lower() in known], key=lambda w: known[w.lower()])
    rest = [w for w in words_list if w.lower() not in known and w.lower() != "total"]
    return first + rest + [w for w in words_list if w.lower() == "total"]
OLD_STATEMENT_PATTERN = "{supplier}_{date}_STM"                  # sheets made before 5.8.0: read as the new pattern

ASAHI_LABELS = ("Asahi", "", "", "", "total gst", "net total ex gst", "total payable")     # 5.15.0: the totals row ends with Total Payable
MOBIGO_LABELS = ("MobiGo", "", "order #|order no|order number", "", "", "", "")              # 5.16.0: its Tax Invoice prints only 'Order #'
PETERS_LABELS = ("Peters", "", "", "inv date (ddmmyy)", "", "", "")      # 5.16.0: INV DATE 220726 = 22/07/2026 ('(ddmmyy)' = that format)

LABELS = [
    ("*",
     "deliver to|delivery address|ship to|invoice to|bill to|sold to|customer",
     "invoice number|invoice no|invoice nbr|invoice #|inv no|document no|document number|doc no|"
     "transaction number|transaction no|credit note no|credit note number|credit memo no|"
     "adjustment note no|pod no|pod number|delivery no|docket no|tax invoice|invoice|credit note|"
     "reference no|reference number|inv|crm id",
     "invoice date|transaction date|trans date|statement date|date of issue|issue date|tax invoice date|order date|date",
     "total gst|gst total|gst value|gst amount|total tax|gst included|tax included|gst|tax",
     "net total ex gst|total excl gst|total ex gst|excl gst value|amount excl gst|total before gst|net amount|"
     "net line total|total taxable|taxable amount|sales for tax|merchandise amount|sub total|net total",
     "|".join(PAYABLE_ORDER) + "|total"),
    ("Supas", "deliver to", "tax invoice", "invoice date", "gst", "total excl gst", "total inc gst"),
    ASAHI_LABELS,
    MOBIGO_LABELS,
    PETERS_LABELS,
    ("Bidfood", "deliver to", "document no", "invoice date", "gst value", "excl gst value", "total value"),
    ("Bonfect", "customer", "transaction number", "transaction date", "gst", "taxable amount", "total inc gst"),
    ("Coca-Cola|Coke", "deliver to", "tax invoice", "date", "total gst", "total taxable", "total price"),
    ("Direct Bakery", "invoice to", "invoice#", "date", "gst", "sub total", "total inclusive gst"),
    ("The Distributors", "deliver to acc", "document no", "transaction date", "gst", "taxable amount", "total inc gst"),
    ("Ezymart", "invoice to", "reference no", "date", "tax included", "sub total", "total aud"),
    ("Essentials International", "", "invoice number", "invoice date", "", "", ""),
    ("Chauncy", "invoice to|bill to", "invoice no", "invoice date", "tax gst", "sales for tax", "invoice total"),
    ("Suntory", "deliver to", "invoice no", "date", "total gst", "total excl gst", "total incl gst"),
    ("Lidcombe Ice", "deliver to", "inv", "date", "gst", "merchandise amount", "invoice total"),
]

SPECIAL_MAILBOXES = [
    # mailbox (as in Outlook, | between names), role, destination folder (LIVE mode), active, notes
    ("P I Fuel Invoice|pifuelinvoice@pearlenergy.com.au", "Fuel invoices", r"C:\SHAREPOINT\pearlenergy.com.au\Patel Group - Documents\OFFICE INDIA\FUEL INVOICES", "Yes",
     "Shared mailbox (opened through office@). Fuel invoices -> <Destination folder>\\<COMPANY>\\<INVOICE NUMBER>_<SITE>.pdf, listed in Reports\\PI_Invoice_list_<MON>_<YEAR>.xlsx. Read in the invoice run. The mailbox is never a company clue."),
    ("mstatement@pearlenergy.com.au", "Fleet card", r"C:\SHAREPOINT\pearlenergy.com.au\Patel Group - Documents\OFFICE INDIA\TEMPORARY DOWNLOADS\FLEETCARD STATEMENTS", "Yes",
     "Fleet Card Merchant Payment Advice -> <Destination folder>\\<SITE>\\<report issue date>_STM.pdf, listed in Reports\\FleetCard_Record_<MON>_<YEAR>.xlsx. Read in the invoice run."),
    ("Security|security@pearlenergy.com.au", "Security", r"C:\SHAREPOINT\pearlenergy.com.au\Patel Group - Documents\OFFICE INDIA\TEMPORARY DOWNLOADS\SECURITY REPORTS", "Yes",
     "Shared mailbox with the weekly alarm reports. Read ONLY by the security run (its own buttons and schedule) -> <Destination folder>\\<MONTH YEAR>\\Downloaded Reports + SITE_DATA_... fortnightly Excel. Settings: Security_Settings.xlsx."),
    ("office@pearlenergy.com.au", "Never", "", "Yes", "The account the shared mailboxes are opened through. Its own inbox is NEVER read."),
    ("serviceinvoices@pearlenergy.com.au", "Never", "", "Yes", "Not read."),
]
SPECIAL_ROLES = ("fuel invoices", "fleet card", "security", "statements", "on hold", "never", "read all")
READ_ALL_TEXT = "Read every e-mail"        # the Role text of a "read every e-mail" mailbox (5.15.0)
TEST_FOLDERS = {"fuel invoices": "FUEL INVOICES", "fleet card": "FLEETCARD STATEMENTS", "security": "SECURITY REPORTS",
                "statements": "M STATEMENTS"}

SITE_HEADERS = ["Site name", "Company", "Aliases", "Address fragments", "Handler", "Handler email",
                "Active", "Source", "Notes", "SharePoint folder"]

SHEETS = {
    "Run": (["Setting", "Value", "Notes"], RUN_DEFAULTS, [38, 34, 110]),
    "Keywords": (["Keyword", "Document type", "Active"], [(k, t, "Yes") for k, t in KEYWORDS], [34, 28, 10]),
    "BlockedWords": (["Word", "Active"], [(w, "Yes") for w in BLOCKED], [30, 10]),
    "DocTypes": (["Document type", "Route", "Subfolder", "Read amounts", "Sign", "Required fields",
                  "File tag", "Name pattern", "Folder layout", "Notes"], DOCTYPES, [26, 18, 26, 14, 8, 24, 10, 30, 22, 60]),
    "Senders": (["Sender contains", "Always download", "Default document type", "Supplier name", "Notes"],
                SENDERS, [30, 16, 22, 46, 70]),
    "Companies": (["Company name", "Text to look for", "Folder name", "Mailbox (as in Outlook)", "Active"],
                  [c + ("Yes",) for c in COMPANIES], [36, 28, 30, 62, 10]),
    "Suppliers": (["Name contains", "Record name", "Short name", "Folder override"], SUPPLIERS, [30, 50, 18, 24]),
    "Sites": (SITE_HEADERS, [], [34, 36, 36, 38, 14, 30, 8, 18, 40, 34]),
    "SiteAliases": (["Other names for the site (| between them)", "Site name"], SITE_ALIASES, [70, 36]),
    "SupplierIDs": (["Supplier (record name)", "ABN", "E-mail domain", "Account name contains"], SUPPLIER_IDS, [40, 18, 26, 36]),
    "LegalEndings": (["Legal ending", "Active"], [(e, "Yes") for e in LEGAL_ENDINGS], [24, 10]),
    "SpecialMailboxes": (["Mailbox (as in Outlook)", "Role", "Destination folder", "Active", "Notes"], SPECIAL_MAILBOXES, [46, 14, 70, 10, 90]),
    "Accounts": (["Supplier name contains", "Account number", "Site name"], [], [28, 20, 34]),
    "Labels": (["Supplier name contains", "Site / deliver-to label", "Invoice number label", "Invoice date label",
                "GST amount label", "Ex-GST / taxable amount label", "Total (incl. GST) label"], LABELS,
               [28, 30, 40, 30, 30, 36, 40]),
}

ABOUT_TITLE = "Outlook Invoice Bot  -  settings"
ABOUT_INTRO = ("The bot reads this file at the start of every run. Save and CLOSE it before a run. "
               "Most days nobody needs to open it: the control panel (Start_Bot.bat) changes the important ones for you.")
ABOUT_TABLE = [
    # sheet, what it is for, how to change it
    ("Run", "Options: test/live, the SharePoint folder, limits, switches. Each row explains itself.", "Change the value in column B."),
    ("Keywords", "Words in an e-mail SUBJECT that mean 'download the attachment' (inv, cn, pod, statement ...).", "Add a row. Set Active = No to switch a word off."),
    ("BlockedWords", "Words in the SUBJECT that mean 'do not download, flag red' (overdue, urgent, reminder, stop supply, warning).", "Add a row."),
    ("DocTypes", "What each document type is: where it goes, whether amounts are read, credit notes negative, the file-name pattern.", "Rarely changed."),
    ("Senders", "'Always download' rows = exception senders (Rarity, Moorees ...). Other rows only help to name a supplier.", "Add a row for a new exception sender."),
    ("Companies", "The 13 companies: Outlook mailbox name, the SharePoint FOLDER name, and the text printed on documents.", "Use Setup > Find my Outlook mailboxes in the control panel."),
    ("Suppliers", "Spelling clean-up for supplier names. 'Folder override' sends a supplier to another folder (e.g. MILK INVOICES).", "Add a row when a supplier name looks wrong."),
    ("SiteAliases", "Other names a document uses for a site (for example 'Pearl Energy Wyong' = Pearl Energy North Wyong).", "Add a row: the names, then the site exactly as in the Sites sheet."),
    ("SupplierIDs", "For suppliers whose name is not printed (or is printed oddly): the ABN, the e-mail domain or the account name identifies them. Any ONE is enough.", "Add a row for a supplier that lands in _Unrecognised with 'supplier not recognised'."),
    ("LegalEndings", "Words that end a company name (Pty, Pvt, Private, Limited, Ltd ...). A line ending in one of them is taken as the supplier name; they are treated as the same when names are compared.", "Add a row for another ending."),
    ("Sites", "The site list: used to find the site of a document, and (column SharePoint folder) the real folder name of each site.", "Fill the SharePoint folder column with Setup > Match my sites to the folders; re-import the list with Advanced\\Import_sites.bat."),
    ("SpecialMailboxes", "Mailboxes that are NOT one company: Fleet card (P I Fuel Invoice), Statements (mstatement@), On hold (Security), Never (office@). Role and Destination folder per mailbox.", "Change the Role or the Destination folder. Never = the bot does not read it."),
    ("Accounts", "A supplier's account number for a site (optional).", "Add a row."),
    ("Labels", "The words each supplier prints next to invoice no., date, GST and total. Starting hints that you improve from real results.", "Add the wording of a supplier the bot cannot read yet."),
]
ABOUT_NOTES = [
    "Columns are found by their HEADING, so adding or re-ordering columns is safe.",
    "Several alternatives in one cell are separated by |   for example  document no|doc no",
    "Folder structure the bot builds:   <root> \\ <MONTH> <YEAR> EMAIL INVOICES \\ <company folder> \\ <site folder> \\ <file>",
]
TAB_COLORS = {"SiteAliases": "548235", "SupplierIDs": "548235", "LegalEndings": "2F5597", "Run": "C00000", "Keywords": "2F5597", "BlockedWords": "C00000", "DocTypes": "2F5597", "Senders": "2F5597",
              "SpecialMailboxes": "C00000", "Companies": "548235", "Suppliers": "548235", "Sites": "548235", "Accounts": "7F7F7F", "Labels": "ED7D31"}

# ----------------------------------------------------------------------------------------
# Small text helpers shared by the whole program
# ----------------------------------------------------------------------------------------

def norm(text):
    """Lower-case letters and digits only:  'NO: 1' and 'No1' compare equal."""
    return re.sub(r"[^a-z0-9]", "", str(text).lower())


def words(text):
    """Lower-case words separated by single spaces:  'Mobil Taree  South' -> 'mobil taree south'."""
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9]+", " ", str(text).lower())).strip()


def has_phrase(text_words, phrase):
    """True if the phrase appears in the text as whole words (both already passed through words())."""
    return bool(phrase) and (" " + phrase + " ") in (" " + text_words + " ")


def yes(value, default=False):
    v = str(value or "").strip().lower()
    if not v:
        return default
    return v in ("yes", "y", "true", "1", "on")


# ----------------------------------------------------------------------------------------
# The workbook
# ----------------------------------------------------------------------------------------

def settings_file():
    return SETTINGS_FILE_OVERRIDE or os.environ.get("BOT_SETTINGS") or paths.p("Bot_Settings.xlsx")


def create_default_file(path=None, sites=None):
    """Write a fresh settings workbook. `sites` = optional list of Sites-sheet rows (lists)."""
    path = path or settings_file()
    wb = Workbook()
    ws = wb.active
    ws.title = "About"
    ws.sheet_properties.tabColor = "1F3864"
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABC", (18, 100, 52)):
        ws.column_dimensions[col].width = w
    ws.append([ABOUT_TITLE])
    ws["A1"].font = Font(bold=True, size=18, color="FFFFFF")
    for col in "ABC":
        ws[f"{col}1"].fill = PatternFill("solid", fgColor="1F3864")
    ws.row_dimensions[1].height = 34
    ws.append([ABOUT_INTRO])
    ws["A2"].font = Font(italic=True, color="595959")
    ws.append([])
    ws.append(["Sheet", "What it is for", "How to change it"])
    for c in ws[4]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="2F5597")
    for sheet, what, how in ABOUT_TABLE:
        ws.append([sheet, what, how])
        ws.cell(ws.max_row, 1).font = Font(bold=True, color=TAB_COLORS.get(sheet, "000000"))
        for c in (2, 3):
            ws.cell(ws.max_row, c).alignment = Alignment(wrap_text=True, vertical="top")
    ws.append([])
    for note in ABOUT_NOTES:
        ws.append([note])
        ws.cell(ws.max_row, 1).font = Font(italic=True, color="595959")
    for name, (headers, rows, widths) in SHEETS.items():
        sh = wb.create_sheet(name)
        sh.append(headers)
        for cell in sh[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
            cell.alignment = Alignment(vertical="center")
        for i, w in enumerate(widths):
            sh.column_dimensions[get_column_letter(i + 1)].width = w
        data = sites if (name == "Sites" and sites is not None) else rows
        for row in data:
            sh.append(list(row))
        sh.freeze_panes = "A2"
        sh.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(sh.max_row, 2)}"
        sh.sheet_properties.tabColor = TAB_COLORS.get(name, "7F7F7F")
        for col, h in enumerate(headers, start=1):          # Yes/No drop-downs
            if h in ("Active", "Always download", "Read amounts"):
                dv = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True)
                sh.add_data_validation(dv)
                dv.add(f"{get_column_letter(col)}2:{get_column_letter(col)}2000")
    for sheet in wb.worksheets:                              # every sheet prints one page wide, A4 landscape
        sheet.page_setup.orientation, sheet.page_setup.paperSize = "landscape", 9
        sheet.page_setup.fitToWidth, sheet.page_setup.fitToHeight = 1, 0
        sheet.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    run = wb["Run"]
    for r in range(2, run.max_row + 1):
        run.cell(r, 3).alignment = Alignment(wrap_text=True, vertical="top")
        if run.cell(r, 1).value == "Destination mode":
            dv = DataValidation(type="list", formula1='"test,live"', allow_blank=False)
            run.add_data_validation(dv)
            dv.add(f"B{r}")
    for r in range(2, wb["Accounts"].max_row + 200):          # keep leading zeros in account numbers
        wb["Accounts"].cell(r, 2).number_format = "@"
    wb.save(path)
    return path


def _cell_text(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).strip()


_cache = None


def _defaults_in_memory():
    return {name: [dict(zip(headers, [str(c) for c in row])) for row in rows]
            for name, (headers, rows, _w) in SHEETS.items()}


def _load():
    path = settings_file()
    if not os.path.exists(path):
        create_default_file(path)
    try:
        wb = load_workbook(path, data_only=True)
    except Exception as e:                  # NEVER fall back silently to defaults: that could send files to the wrong place
        raise SettingsError(f"{os.path.basename(path)} cannot be read ({type(e).__name__}: {e}).\n"
                            "   Close it in Excel (and check it is a valid .xlsx file), then run again.")
    defaults = _defaults_in_memory()
    tables = {}
    for name, (headers, _rows, _w) in SHEETS.items():
        if name not in wb.sheetnames:
            tables[name] = defaults[name]
            continue
        sh = wb[name]
        file_headers = [_cell_text(c) for c in next(sh.iter_rows(min_row=1, max_row=1, values_only=True), [])]
        index = {h: i for i, h in enumerate(file_headers) if h}
        rows = []
        for raw in sh.iter_rows(min_row=2, values_only=True):
            cells = [_cell_text(c) for c in raw]
            rec = {h: (cells[index[h]] if h in index and index[h] < len(cells) else "") for h in headers}
            if rec[headers[0]]:
                rows.append(rec)
        tables[name] = rows
    wb.close()
    if not tables["Labels"]:
        tables["Labels"] = defaults["Labels"]
    return tables


def edit(change):
    """Change Bot_Settings.xlsx from the program (used by the control panel).
    `change(workbook)` edits an openpyxl workbook; a rolling backup (Bot_Settings.backup.xlsx) is kept."""
    import shutil
    path = settings_file()
    if not os.path.exists(path):
        create_default_file(path)
    try:
        wb = load_workbook(path)
    except Exception as e:
        raise SettingsError(f"{os.path.basename(path)} cannot be opened ({type(e).__name__}). Close it in Excel and try again.")
    change(wb)
    try:
        shutil.copyfile(path, path.replace(".xlsx", ".backup.xlsx"))
        wb.save(path)
    except PermissionError:
        raise SettingsError(f"{os.path.basename(path)} is open in Excel. Close it and try again.")
    finally:
        wb.close()
    reset()


SETTINGS_VERSION = "5.20.0"


def _version_tuple(text):
    return tuple(int(x) for x in re.findall(r"\d+", str(text or ""))[:3]) or (0,)


def _sheet_rows(ws):
    """(header index by name, list of (row number, {header: cell}))"""
    heads = {str(c.value or "").strip(): i + 1 for i, c in enumerate(ws[1]) if c.value}
    return heads


def _migrate(book, from_version=""):
    """Every one-time step the file does not have yet (by its Settings version)."""
    done = _migrate_before_5_15(book, from_version)
    if _version_tuple(from_version) < (5, 15, 0):
        done += _migrate_5_15(book)
    if _version_tuple(from_version) < (5, 16, 0):
        done += _migrate_5_16(book)
    if _version_tuple(from_version) < (5, 18, 0):
        done += _migrate_5_18(book)
    if _version_tuple(from_version) < (5, 19, 0):
        done += _migrate_5_19(book)
    if _version_tuple(from_version) < (5, 20, 0):
        done += _migrate_5_20(book)
    return done


SITE_ADDITIONS_5_20 = [            # (site, aliases to add, address fragments to add) - from the _Unassigned review of October 2026
    ("Mobil Hasting River", ["Mobil Port Macquarie Hastings"], []),   # Asahi: 'MOBIL PORT MACQUARIE HASTINGS, 128 Hastings River Dr'
    ("Mobil Moss Vale", ["Mobil Mossvale"], []),                       # Bega sold-to 'Mobil Mossvale'
    ("Ampol Mittagong", ["Pearl Energy Mittagong"], []),               # PFD ship-to 'PEARL ENERGY MITTAGONG, 65 Bowral Road'
    ("Pearl Energy Beresfield", ["Pearl Energy Beresfield South"], ["yangan drive beresfield"]),   # Slick Chicks
    ("Pearl Energy North Wyong", [], ["amy cl wyong", "1 10 amy"]),     # Neverfail: delivery address '1-10 Amy Cl Wyong'
]


def _migrate_5_20(book):
    """5.20.0: names and addresses suppliers print for some sites (added, never replaced)."""
    done = []
    if "Sites" not in book.sheetnames:
        return done
    ws = book["Sites"]
    h = _sheet_rows(ws)
    nc, ac, fc = h.get("Site name"), h.get("Aliases"), h.get("Address fragments")
    if not (nc and ac and fc):
        return done
    for site, aliases, frags in SITE_ADDITIONS_5_20:
        for r in range(2, ws.max_row + 1):
            if norm(ws.cell(r, nc).value or "") != norm(site):
                continue
            for col, add in ((ac, aliases), (fc, frags)):
                have = _alternatives(ws.cell(r, col).value or "")
                new = [a for a in add if norm(a) not in [norm(x) for x in have]]
                if new:
                    ws.cell(r, col).value = "|".join(have + new)
                    done.append(f"Sites '{site}': + {', '.join(new)}")
            break
    return done


def _migrate_5_19(book):
    """5.19.0: the Notes of Fleet Card Statement / Fuel Invoice no longer say 'only in that mailbox'."""
    done = []
    if "DocTypes" not in book.sheetnames:
        return done
    ws = book["DocTypes"]
    h = _sheet_rows(ws)
    tc, nc = h.get("Document type"), h.get("Notes")
    notes = {"fleet card statement": "Fleet Card Merchant Payment Advice: <Destination folder of the Fleet card mailbox>\\<SITE>\\<report issue date>_STM, whichever mailbox it came to (5.19.0). The same statement from two mailboxes is saved once.", "fuel invoice": "Fuel invoices from Pearl Investment Family Trust: <Destination folder of the Fuel invoices mailbox>\\<COMPANY>\\<INVOICE NUMBER>_<SITE>.pdf, whichever mailbox it came to (5.19.0; company and site from the invoice's Customer, never from the Reference)."}
    if tc and nc:
        for r in range(2, ws.max_row + 1):
            key = str(ws.cell(r, tc).value or "").strip().lower()
            if key in notes and "ONLY" in str(ws.cell(r, nc).value or ""):
                ws.cell(r, nc).value = notes[key]
                done.append(f"DocTypes '{ws.cell(r, tc).value}': notes updated (any mailbox)")
    return done


def _migrate_5_18(book):
    """5.18.0: Customer Activity Reports (Metcash) are filed in the site folder, no longer parked in _Unrecognised."""
    done = []
    if "DocTypes" not in book.sheetnames:
        return done
    ws = book["DocTypes"]
    h = _sheet_rows(ws)
    tc, rc, fc, qc, nc = (h.get("Document type"), h.get("Route"), h.get("Subfolder"), h.get("Required fields"), h.get("Notes"))
    if not (tc and rc):
        return done
    for r in range(2, ws.max_row + 1):
        if str(ws.cell(r, tc).value or "").strip().lower() == "customer activity report" \
                and str(ws.cell(r, rc).value or "").strip().lower() == "unrecognised":
            ws.cell(r, rc).value = "Company folder"
            if fc:
                ws.cell(r, fc).value = None
            if qc and not ws.cell(r, qc).value:
                ws.cell(r, qc).value = "date"
            if nc:
                ws.cell(r, nc).value = "5.18.0: filed in the site folder like an invoice (Metcash Customer Activity Report, also when its PDF says Statement). No amounts; not in the A/C Excel."
            done.append("DocTypes 'Customer Activity Report': filed in the site folder (was parked in _Unrecognised)")
    return done


def _migrate_5_16(book):
    """5.16.0: Labels rows for MobiGo (the Order # is the invoice number) and Peters (INV DATE as DDMMYY). A row already there
    for the supplier is completed, never replaced."""
    done = []
    if "Labels" not in book.sheetnames:
        return done
    ws = book["Labels"]
    h = _sheet_rows(ws)
    sc = h.get("Supplier name contains")
    if not sc:
        return done
    heads = ["Supplier name contains"] + _LABEL_HEADERS
    for row in (MOBIGO_LABELS, PETERS_LABELS):
        name = row[0]
        found = next((r for r in range(2, ws.max_row + 1)
                      if norm(name) in [norm(a) for a in _alternatives(ws.cell(r, sc).value or "")]), None)
        if found is None:
            ws.append(list(row))
            done.append(f"Labels row '{name}'")
            continue
        for head, value in zip(heads[1:], row[1:]):
            col = h.get(head)
            if value and col:
                have = _alternatives(ws.cell(found, col).value or "")
                add = [v for v in _alternatives(value) if v.lower() not in [x.lower() for x in have]]
                if add:
                    ws.cell(found, col).value = "|".join(add + have)       # the new words first: they are the reason for this row
                    done.append(f"Labels '{name}': {head} + {', '.join(add)}")
    return done


def _run_rows(book):
    run = book["Run"]
    return run, {str(run.cell(r, 1).value or "").strip().lower(): r for r in range(2, run.max_row + 1)}


def _migrate_5_15(book):
    """5.15.0: the LIVE folder is OFFICE INDIA\\EMAIL INVOICES, with HOLDINGS and Reports inside it; LIVE creates missing folders like
    TEST; the payable total wins; small corrections agreed with the payments team (Labels, Ampol Kelso, the Jiya@ test mailbox)."""
    done = []
    run, rows_ = _run_rows(book)
    notes = {k.lower(): n for k, _v, n in RUN_DEFAULTS}

    def setv(key, value, why):
        r = rows_.get(key.lower())
        if r is None:
            run.append([key, value, notes.get(key.lower(), "")])
            done.append(f"Run '{key}' = {value or '(blank)'} ({why})")
            return
        if str(run.cell(r, 2).value or "").strip() != value:
            run.cell(r, 2).value = value or None
            done.append(f"Run '{key}' = {value or '(blank)'} ({why})")
        if key.lower() in notes:
            run.cell(r, 3).value = notes[key.lower()]

    def value(key):
        r = rows_.get(key.lower())
        return str(run.cell(r, 2).value or "").strip() if r else ""

    live = value("Live destination root")
    if not live or norm(live) in (norm(OLD_MAILS), norm(OLD_MAILS.rsplit("\\", 1)[0] + "\\MAIL INVOICES")):
        setv("Live destination root", LIVE_ROOT, "the real EMAIL INVOICES folder")
    for key in ("Holding root", "Reports root"):
        if value(key).lower().startswith(OLD_MAILS.lower()):
            setv(key, "", "now inside the EMAIL INVOICES folder")
        elif key.lower() in notes and rows_.get(key.lower()):
            run.cell(rows_[key.lower()], 3).value = notes[key.lower()]
    if value("Auto-find folder name").upper() in ("", "MAIL INVOICES"):
        setv("Auto-find folder name", "EMAIL INVOICES", "the real folder name")
    setv("Live mode creates missing folders like test mode", "Yes", "LIVE makes missing folders exactly like TEST")
    setv("Create site folder in live mode", "Yes", "LIVE makes missing site folders")
    for key in ("Find the destination folder automatically",):
        if rows_.get(key.lower()):
            run.cell(rows_[key.lower()], 3).value = notes[key.lower()]

    if "Labels" in book.sheetnames:                                          # the '*' row: payable first, 'Advice no' is a number
        ws = book["Labels"]
        h = _sheet_rows(ws)
        sc, nc, dc, tc = (h.get("Supplier name contains"), h.get("Invoice number label"), h.get("Invoice date label"),
                          h.get("Total (incl. GST) label"))
        if sc and nc and dc and tc:
            have_asahi = False
            for r in range(2, ws.max_row + 1):
                key = str(ws.cell(r, sc).value or "").strip()
                if "asahi" in key.lower():
                    have_asahi = True
                if key not in ("*", "default", "DEFAULT"):
                    continue
                dates = _alternatives(ws.cell(r, dc).value or "")
                advice = [d for d in dates if norm(d) == "adviceno"]
                if advice:
                    ws.cell(r, dc).value = "|".join(d for d in dates if norm(d) != "adviceno")
                    nums = _alternatives(ws.cell(r, nc).value or "")
                    if not any(norm(n) == "adviceno" for n in nums):
                        ws.cell(r, nc).value = "|".join(nums + advice[:1])
                    done.append("Labels '*': 'Advice no' moved from the invoice DATE words to the invoice NUMBER words (last)")
                totals = _alternatives(ws.cell(r, tc).value or "")
                have = {t.lower() for t in totals}
                new = payable_sorted(totals + [t for t in ADDED_LABELS["total"] if t not in have])
                if new != totals:
                    ws.cell(r, tc).value = "|".join(new)
                    done.append("Labels '*': the TOTAL words now start with total payable / amount payable (the amount payable wins)")
            if not have_asahi:
                ws.append(list(ASAHI_LABELS))
                done.append("Labels row 'Asahi' (Total Payable)")

    if "Sites" in book.sheetnames:                                           # Ampol Kelso: the address printed on Bega invoices too
        ws = book["Sites"]
        h = _sheet_rows(ws)
        nc, fc = h.get("Site name"), h.get("Address fragments")
        if nc and fc:
            for r in range(2, ws.max_row + 1):
                if norm(ws.cell(r, nc).value or "") == "ampolkelso":
                    frags = _alternatives(ws.cell(r, fc).value or "")
                    add = [f for f in ("17 great western", "great western hwy kelso") if f not in [x.lower() for x in frags]]
                    if add:
                        ws.cell(r, fc).value = "|".join(frags + add)
                        done.append("Sites 'Ampol Kelso': address 17 Great Western Hwy added (19 Sydney Road kept)")

    if "Companies" in book.sheetnames and "SpecialMailboxes" in book.sheetnames:   # Jiya@: read every e-mail, no TEST JIYA folder
        ws, sp = book["Companies"], book["SpecialMailboxes"]
        h = _sheet_rows(ws)
        cc, mc = h.get("Company name"), h.get("Mailbox (as in Outlook)")
        if cc and mc:
            for r in range(ws.max_row, 1, -1):
                if norm(ws.cell(r, cc).value or "") == norm("TEST - Jiya@pearlenergy.com.au"):
                    box = str(ws.cell(r, mc).value or "").split("|")[0].strip() or "Jiya@pearlenergy.com.au"
                    ws.delete_rows(r)
                    if not any(norm(box) in [norm(a) for a in str(sp.cell(k, 1).value or "").split("|")] for k in range(2, sp.max_row + 1)):
                        sp.append([box, READ_ALL_TEXT, "", "Yes", "Every e-mail is read (Inbox + sub-folders); the company and site come "
                                   "from each document, never from this mailbox."])
                    done.append(f"{box}: TEST mailbox -> Read every e-mail (no TEST JIYA folder)")
    return done


def _migrate_before_5_15(book, from_version=""):
    """One-time changes when an older Bot_Settings.xlsx meets a newer version (each step only once, by version). Values a person typed are only changed where 5.11.0
    deliberately changes the rule (folder creation, the document type of three keywords); everything is listed in the run summary."""
    done = []
    old = _version_tuple(from_version)
    if old >= (5, 12, 0):
        return done + _migrate_5_13(book)
    if old >= (5, 11, 0):
        return done + _migrate_5_12(book) + _migrate_5_13(book)
    run = book["Run"]
    current = {str(run.cell(r, 1).value or "").strip().lower(): r for r in range(2, run.max_row + 1)}
    for key, value in (("Live mode creates missing folders like test mode", "No"), ("Create month folder in live mode", "Yes"),
                       ("Create company folder in live mode", "Yes"), ("Create site folder in live mode", "No")):
        r = current.get(key.lower())
        if r and str(run.cell(r, 2).value or "").strip() != value:
            run.cell(r, 2).value = value
            done.append(f"Run '{key}' = {value}")
    if "Keywords" in book.sheetnames:
        ws = book["Keywords"]
        h = _sheet_rows(ws)
        kc, tc = h.get("Keyword"), h.get("Document type")
        have = set()
        if kc and tc:
            for r in range(2, ws.max_row + 1):
                kw = str(ws.cell(r, kc).value or "").strip().lower()
                have.add(kw)
                want = {"direct debit notice": "Direct Debit Notice", "fleetcard merchant payment advice": "Fleet Card Statement"}.get(kw)
                if want and str(ws.cell(r, tc).value or "").strip().lower() == "other":
                    ws.cell(r, tc).value = want
                    done.append(f"Keyword '{kw}' = {want}")
            for kw, typ in (("merchant payment advice", "Fleet Card Statement"), ("account activity listing", "Account Activity Listing"),
                            ("direct debit notice", "Direct Debit Notice")):
                if kw not in have:
                    row = [""] * max(h.values())
                    row[kc - 1], row[tc - 1] = kw, typ
                    if h.get("Active"):
                        row[h["Active"] - 1] = "Yes"
                    ws.append(row)
                    done.append(f"Keyword '{kw}'")
    if "DocTypes" in book.sheetnames:
        ws = book["DocTypes"]
        h = _sheet_rows(ws)
        if "Folder layout" not in h:
            col = ws.max_column + 1
            ws.cell(1, col).value = "Folder layout"
            ws.cell(1, col).font = Font(bold=True, color="FFFFFF")
            ws.cell(1, col).fill = PatternFill("solid", fgColor="1F4E78")
            ws.column_dimensions[get_column_letter(col)].width = 22
            h["Folder layout"] = col
        tcol = h.get("Document type", 1)
        names = {}
        for r in range(2, ws.max_row + 1):
            names[str(ws.cell(r, tcol).value or "").strip().lower()] = r
        layouts = {"statement": "{MONTH}\\{COMPANY}", "order confirmation": "{MONTH}"}
        for typ, lay in layouts.items():
            r = names.get(typ)
            if r and not str(ws.cell(r, h["Folder layout"]).value or "").strip():
                ws.cell(r, h["Folder layout"]).value = lay
                done.append(f"DocTypes '{typ}' folder layout {lay}")
        headers = SHEETS["DocTypes"][0]
        for d in DOCTYPES:
            if d[0].lower() not in names:
                row = [""] * ws.max_column
                for name, value in zip(headers, d):
                    if name in h:
                        row[h[name] - 1] = value
                ws.append(row)
                done.append(f"DocTypes row '{d[0]}'")
    if "Sites" in book.sheetnames:
        ws = book["Sites"]
        for row in ws.iter_rows(min_row=2):
            for c in row:
                if isinstance(c.value, str) and "hanharan" in c.value.lower():
                    c.value = re.sub("hanharan", "hanrahan", c.value, flags=re.IGNORECASE)
                    done.append("Sites: 'hanharan' corrected to 'hanrahan'")
    return done + _migrate_5_12(book) + _migrate_5_13(book)


OLD_MAILS = r"C:\SHAREPOINT\pearlenergy.com.au\Patel Group - Documents\OFFICE INDIA\MAILS"


def _migrate_5_13(book):
    """5.13.0: the holding folders move into OFFICE INDIA\\MAILS\\HOLDINGS (only when Holding root is still the old MAILS folder;
    files already there are not moved). The Reports root stays MAILS\\Reports."""
    run = book["Run"]
    for r in range(2, run.max_row + 1):
        if str(run.cell(r, 1).value or "").strip().lower() == "holding root":
            value = str(run.cell(r, 2).value or "").strip().rstrip("\\")
            if value.lower() == OLD_MAILS.lower():
                run.cell(r, 2).value = OLD_MAILS + "\\HOLDINGS"
                return ["Run 'Holding root' = ...\\MAILS\\HOLDINGS (files already in MAILS are left where they are)"]
    return []


def _migrate_5_12(book):
    """5.12.0: P I Fuel Invoice = Fuel invoices (FUEL INVOICES), mstatement@ = Fleet card (FLEETCARD STATEMENTS),
    Security = Security (SECURITY REPORTS). The Fuel Invoice document type is added by the DocTypes step."""
    done = []
    if "SpecialMailboxes" not in book.sheetnames:
        return done
    ws = book["SpecialMailboxes"]
    want = {norm(n): row for row in SPECIAL_MAILBOXES for n in row[0].split("|")}
    seen = set()
    for r in range(2, ws.max_row + 1):
        names = [norm(n) for n in str(ws.cell(r, 1).value or "").split("|") if n.strip()]
        hit = next((want[n] for n in names if n in want), None)
        if not hit or hit[1].lower() == "never":
            if hit:
                seen.add(hit[0])
            continue
        seen.add(hit[0])
        if str(ws.cell(r, 2).value or "").strip().lower() != hit[1].lower():
            ws.cell(r, 2).value, ws.cell(r, 3).value, ws.cell(r, 5).value = hit[1], hit[2], hit[4]
            done.append(f"SpecialMailboxes: {hit[0].split('|')[0]} = {hit[1]}")
    for row in SPECIAL_MAILBOXES:
        if row[0] not in seen:
            ws.append(list(row))
            done.append(f"SpecialMailboxes: {row[0].split('|')[0]} added")
    if "DocTypes" in book.sheetnames:
        dt = book["DocTypes"]
        h = _sheet_rows(dt)
        names = {str(dt.cell(r, h.get("Document type", 1)).value or "").strip().lower() for r in range(2, dt.max_row + 1)}
        for d in DOCTYPES:
            if d[0].lower() not in names:
                row = [""] * dt.max_column
                for name, value in zip(SHEETS["DocTypes"][0], d):
                    if name in h:
                        row[h[name] - 1] = value
                dt.append(row)
                done.append(f"DocTypes row '{d[0]}'")
    return done


def upgrade_workbook():
    """Add to an older Bot_Settings.xlsx what a newer version needs: the sheets it does not have yet (with their default rows) and the missing Run
    settings. Nothing that exists is changed. A backup (Bot_Settings.backup.xlsx) is kept. Returns a list of what was added ([] = nothing to do)."""
    path = settings_file()
    if not os.path.exists(path):
        return []
    try:
        wb = load_workbook(path, read_only=True)
        missing_sheets = [n for n in SHEETS if n not in wb.sheetnames]
        have = set()
        if "Run" in wb.sheetnames:
            have = {str(r[0] or "").strip().lower() for r in wb["Run"].iter_rows(min_row=2, max_col=1, values_only=True)}
        have_sup = set()
        if "Suppliers" in wb.sheetnames:
            have_sup = {str(r[0] or "").strip().lower() for r in wb["Suppliers"].iter_rows(min_row=2, max_col=1, values_only=True)}
        have_ids = set()
        if "SupplierIDs" in wb.sheetnames:
            have_ids = {str(r[0] or "").strip().lower() for r in wb["SupplierIDs"].iter_rows(min_row=2, max_col=1, values_only=True)}
        wb.close()
    except Exception:
        return []
    try:
        version = ""
        wb = load_workbook(path, read_only=True)
        if "Run" in wb.sheetnames:
            for r in wb["Run"].iter_rows(min_row=2, max_col=2, values_only=True):
                if str(r[0] or "").strip().lower() == "settings version":
                    version = str(r[1] or "").strip()
        wb.close()
    except Exception:
        version = ""
    migrate = _version_tuple(version) < _version_tuple(SETTINGS_VERSION)
    missing_ids = [r for r in SUPPLIER_IDS if r[0].lower() not in have_ids] if "SupplierIDs" not in missing_sheets and have_ids else []
    missing_run = [(k, v, n) for k, v, n in RUN_DEFAULTS if k.lower() not in have]
    missing_sup = [r for r in ADDED_SUPPLIERS if r[0].lower() not in have_sup] if "Suppliers" not in missing_sheets else []
    if not missing_sheets and not missing_run and not missing_ids and not missing_sup and not migrate:
        return []
    added = []

    def change(book):
        for name in missing_sheets:
            headers, rows_, widths = SHEETS[name]
            sh = book.create_sheet(name)
            sh.append(headers)
            for cell in sh[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="1F4E78")
            for i, w in enumerate(widths):
                sh.column_dimensions[get_column_letter(i + 1)].width = w
            for row in rows_:
                sh.append(list(row))
            sh.freeze_panes = "A2"
            sh.sheet_properties.tabColor = TAB_COLORS.get(name, "7F7F7F")
            added.append(f"sheet {name}")
        for row in missing_ids:
            book["SupplierIDs"].append(list(row))
            added.append(f"SupplierIDs row '{row[0]}'")
        for row in missing_sup:
            book["Suppliers"].append(list(row))
            added.append(f"Suppliers row '{row[0]}'")
        for k, v, n in missing_run:
            if k.lower() == "settings version":
                continue
            book["Run"].append([k, v, n])
            added.append(f"Run setting '{k}'")
        if migrate:
            added.extend(_migrate(book, version))
            set_run_value(book, "Settings version", SETTINGS_VERSION)

    try:
        edit(change)
    except SettingsError:
        return []                                      # open in Excel: try again at the next run
    return added


def set_run_value(wb, key, value):
    """Set a row of the Run sheet (the row is added if the workbook is an older version)."""
    ws = wb["Run"]
    for r in range(2, ws.max_row + 1):
        if str(ws.cell(r, 1).value or "").strip().lower() == key.lower():
            ws.cell(r, 2).value = value
            return
    ws.append([key, value, ""])


def get():
    global _cache
    if _cache is None:
        _cache = _load()
    return _cache


def reset():
    global _cache
    _cache = None


def snapshot():
    """A plain, picklable copy of all tables (handed to the reader processes)."""
    return {k: [dict(r) for r in v] for k, v in get().items()}


def use_snapshot(snap):
    global _cache
    _cache = snap


def rows(name):
    return get().get(name, [])


# ----------------------------------------------------------------------------------------
# Run options
# ----------------------------------------------------------------------------------------

def run_value(key, default=""):
    key_l = key.strip().lower()
    for r in rows("Run"):
        if r["Setting"].strip().lower() == key_l:
            return r["Value"].strip() or default
    for k, v, _n in RUN_DEFAULTS:                       # the sheet may be an older version
        if k.lower() == key_l:
            return str(v).strip() or default
    return default


def run_bool(key, default=False):
    return yes(run_value(key), default)


def run_int(key, default=0):
    try:
        return int(float(run_value(key, str(default))))
    except ValueError:
        return default


def _config_value(name):
    try:
        import config                                   # lives in the project folder
    except Exception:
        try:
            import importlib.util
            spec = importlib.util.spec_from_file_location("config", os.path.join(paths.project_dir(), "config.py"))
            config = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(config)
        except Exception:
            return ""
    return str(getattr(config, name, "") or "").strip()


def destination_mode():
    mode = run_value("Destination mode", "test").lower()
    if mode not in ("test", "live"):
        raise SettingsError(f"'Destination mode' in Bot_Settings.xlsx (sheet Run) must be  test  or  live  - it says '{mode}'.")
    return mode


def destination_root():
    """The folder that contains the monthly folders. ('' = live mode but no path yet.)"""
    if destination_mode() == "live":
        # >>> The real SharePoint path is read from here:  Bot_Settings.xlsx (Run) or config.py <<<
        explicit = run_value("Live destination root") or _config_value("LIVE_DESTINATION_ROOT")
        if not run_bool("Find the destination folder automatically", True):
            return explicit
        if explicit and os.path.isdir(explicit):
            return explicit
        from . import autofind
        found = autofind.cached()
        if explicit and not found:
            return explicit                       # typed but missing: the run stops with a clear message (nothing is written)
        return found
        from . import autofind                    # no path typed anywhere: the folder the bot found by itself
        return autofind.cached()
    return run_value("Test destination root") or paths.p("TestOutput")


def is_test_mode():
    try:
        return destination_mode() != "live"
    except SettingsError:
        return True


def _inside_live_root(name, fallback):
    root = ""
    try:
        root = destination_root()
    except SettingsError:
        pass
    return os.path.join(root, name) if root else paths.p(fallback)


def holding_root():
    """LIVE: the Holding root setting, blank = <Live destination root>\\HOLDINGS. TEST: ALWAYS the local Holding folder next to the
    program, so a test never writes into SharePoint."""
    if is_test_mode():
        return paths.p("Holding")
    return run_value("Holding root") or _config_value("HOLDING_ROOT") or _inside_live_root(LIVE_HOLDING, "Holding")


def reports_root():
    """LIVE: the Reports root setting, blank = <Live destination root>\\Reports. TEST: ALWAYS the local Reports folder."""
    if is_test_mode():
        return paths.p("Reports")
    return run_value("Reports root") or _config_value("REPORTS_ROOT") or _inside_live_root(LIVE_REPORTS, "Reports")


def check_output_root(path, what):
    """A configured output folder (Holding, Reports, a special mailbox's destination) must exist, or at least the folder it sits in.
    An unsynced SharePoint folder must never be re-created as an empty local copy."""
    if not path or os.path.isdir(path) or os.path.isdir(os.path.dirname(os.path.normpath(path))):
        return
    raise SettingsError(f"The {what} folder does not exist on this PC:\n   {path}\n"
                        "   Check that OneDrive / SharePoint has synced it (or correct the path in Bot_Settings.xlsx, sheet Run).")


def scratch_root():
    return paths.p("Scratch")


def archive_root():
    return run_value("Archive folder") or paths.p("Archive")


# ----------------------------------------------------------------------------------------
# Suppliers, companies, labels (small lookups used all over the program)
# ----------------------------------------------------------------------------------------

def _alternatives(cell):
    return [a.strip() for a in str(cell).split("|") if a.strip()]


def supplier_row_for(text):
    """The Suppliers row whose 'Name contains' alias appears (as whole words) in the text, or None."""
    t = words(text)
    best = None
    for r in rows("Suppliers"):
        for alt in _alternatives(r["Name contains"]):
            a = words(alt)
            if a and has_phrase(t, a) and (best is None or len(a) > best[0]):
                best = (len(a), r)
    return best[1] if best else None


def legal_endings():
    """The active legal endings (lower case), from the LegalEndings sheet."""
    out = [r["Legal ending"].strip().lower() for r in rows("LegalEndings")
           if r.get("Legal ending", "").strip() and yes(r.get("Active", ""), default=True)]
    return out or [e.lower() for e in LEGAL_ENDINGS]


def strip_endings(name):
    """A name without its legal endings, as words:  'Envisage Holdings Pty Ltd' and 'ENVISAGE HOLDINGS PVT. LIMITED' -> 'envisage holdings'."""
    endings = {words(e) for e in legal_endings()}
    toks = words(name).split()
    while toks and (toks[-1] in endings or (len(toks) > 1 and " ".join(toks[-2:]) in endings)):
        toks = toks[:-2] if (len(toks) > 1 and " ".join(toks[-2:]) in endings) else toks[:-1]
    return " ".join(toks)


def supplier_key(name):
    """Two spellings of a supplier that differ only in the legal ending have the same key."""
    return strip_endings(name) or words(name)


def supplier_id_row_for(text):
    """The SupplierIDs row whose ABN, e-mail domain or account name appears in the text (any ONE is enough), or None."""
    low = str(text).lower()
    w = words(text)
    for r in rows("SupplierIDs"):
        digits = re.sub(r"\D", "", r.get("ABN", ""))
        if len(digits) == 11 and re.search(r"(?<!\d)" + r"\s?".join(digits) + r"(?!\d)", low):
            return r
        domain = r.get("E-mail domain", "").strip().lower().lstrip("@")
        if domain and re.search(r"(?<![a-z0-9-])" + re.escape(domain), low):
            return r
        acct = words(r.get("Account name contains", ""))
        if acct and has_phrase(w, acct):
            return r
    return None


def site_alias_names(site_name):
    """Extra names (words) that the SiteAliases sheet gives to a site."""
    out = []
    for r in rows("SiteAliases"):
        if words(r.get("Site name", "")) == words(site_name):
            out += [words(a) for a in _alternatives(r.get("Other names for the site (| between them)", "")) if words(a)]
    return out


def default_short(name):
    cleaned = re.sub(r"\b(pty|pvt|private|proprietary|ltd|limited|p/l|inc|incorporated|llp|llc|plc)\b\.?", " ", name, flags=re.IGNORECASE)
    cleaned = re.sub(r"[^A-Za-z0-9& ]", "", cleaned).strip()
    return re.sub(r"\s+", "-", cleaned)[:24] or "Supplier"


def canonical_supplier(name):
    """Apply the Suppliers sheet. Returns (record name, short name, folder override)."""
    r = supplier_row_for(name)
    if name and supplier_key(name) in {supplier_key(i["Supplier (record name)"]) for i in rows("SupplierIDs") if i.get("Supplier (record name)")}:
        # a name from the supplier list (ABN / e-mail domain) is kept as it is: an older Suppliers-sheet alias must not rename it
        return name, ((r["Short name"] if r and r["Short name"] else None) or default_short(name)), (r["Folder override"] if r else "")
    if r and r["Record name"]:
        return r["Record name"], (r["Short name"] or default_short(r["Record name"])), r["Folder override"]
    return name, default_short(name), ""


_FIELDS = ["site", "invoice_no", "date", "gst", "ex", "total"]
_LABEL_HEADERS = ["Site / deliver-to label", "Invoice number label", "Invoice date label",
                  "GST amount label", "Ex-GST / taxable amount label", "Total (incl. GST) label"]


def labels_for(printed_name, supplier):
    """Label words for this supplier: its own row first, then the '*' row."""
    names = [words(printed_name), words(supplier)]
    specific, default = [], None
    for r in rows("Labels"):
        key = r["Supplier name contains"].strip()
        if key in ("*", "default", "DEFAULT"):
            default = default or r
            continue
        alts = [words(a) for a in _alternatives(key)]
        hit = [a for a in alts if a and any(has_phrase(n, a) for n in names)]
        if hit:
            specific.append((max(len(a) for a in hit), r))
    specific.sort(key=lambda item: -item[0])
    ordered = [r for _, r in specific] + ([default] if default else [])
    result = {}
    for field, header in zip(_FIELDS, _LABEL_HEADERS):
        seen, out, tail = set(), [], []
        for r in ordered:
            for alt in _alternatives(r.get(header, "")):
                if alt.lower() not in seen:
                    seen.add(alt.lower())
                    (tail if r is default else out).append(alt)
        for extra in ADDED_LABELS.get(field, []):                 # words added in code, for Bot_Settings.xlsx files made before they existed
            if extra.lower() not in seen:
                seen.add(extra.lower())
                tail.append(extra)
        result[field] = out + (payable_sorted(tail) if field == "total" else tail)
    return result
