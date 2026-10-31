"""
=====================================================================================
 NORMALLY NOT NEEDED (5.15.0): the real LIVE folder is already filled in in Bot_Settings.xlsx
=====================================================================================

Everything else (keywords, companies, labels, month-folder pattern, ...) lives in
Bot_Settings.xlsx. The two paths below can be set EITHER here OR in Bot_Settings.xlsx
(sheet "Run", rows "Live destination root" and "Holding root"). If both are filled in,
the value in Bot_Settings.xlsx wins.

The bot starts in TEST mode: it writes into the folder  TestOutput  next to the program.
To go live later:
    1. fill in LIVE_DESTINATION_ROOT below (remove the leading '#'), and
    2. in Bot_Settings.xlsx, sheet "Run", set  "Destination mode"  to  live.

The bot then builds the final path itself, every single time:

    <LIVE_DESTINATION_ROOT> \\ <MONTH> <YEAR> EMAIL INVOICES \\ <COMPANY FOLDER> \\ <file>
    e.g. ...\\OFFICE INDIA\\EMAIL INVOICES \\ OCTOBER 2026 EMAIL INVOICES \\ ONE PEARL ENERGY \\ <SITE> \\ Supplier_12345.pdf
    (HOLDINGS and Reports are made inside the same EMAIL INVOICES folder)

so nothing has to be changed when a new month starts.
"""

# --------------------------------------------------------------------------------------
# >>>>>>>>>>>>>>>>>>>>>>  PUT THE REAL SHAREPOINT PATH HERE (LATER)  <<<<<<<<<<<<<<<<<<<<<
#
# The folder that CONTAINS the monthly folders ("OCTOBER 2026 EMAIL INVOICES", ...).
# Use the path of the library as it appears on the bot PC after OneDrive sync, e.g.:
#
# LIVE_DESTINATION_ROOT = r"C:\SHAREPOINT\pearlenergy.com.au\Patel Group - Documents\OFFICE INDIA\EMAIL INVOICES"
#
LIVE_DESTINATION_ROOT = ""
# --------------------------------------------------------------------------------------

# --------------------------------------------------------------------------------------
# OPTIONAL: where the holding folders (_Unassigned, _Unrecognised\Statement, ...) live.
# Leave empty: LIVE uses <EMAIL INVOICES>\HOLDINGS, TEST the folder "Holding" next to the program.
#
# HOLDING_ROOT = r"C:\SHAREPOINT\pearlenergy.com.au\Patel Group - Documents\OFFICE INDIA\EMAIL INVOICES\HOLDINGS"
#
HOLDING_ROOT = ""
# --------------------------------------------------------------------------------------
