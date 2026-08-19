"""
Settings for sku_pod_tracker.py.

This file is yours -- sku_pod_tracker.py reads it and never overwrites it, so
replacing the script will not wipe your sheet IDs again.
"""

import os

# Path to the service account JSON key.
# GitHub Actions sets SKU_POD_SERVICE_ACCOUNT; on your PC the fallback is used.
SERVICE_ACCOUNT_FILE = os.environ.get(
    "SKU_POD_SERVICE_ACCOUNT",
    r"C:\Users\tbd20\Downloads\SKU POD\service_account.json")

# One Google Sheet per month.
#   key   = the label shown in the dashboard month dropdown
#   value = the sheet ID, i.e. the part of the URL between /d/ and /edit
#           https://docs.google.com/spreadsheets/d/1AbC...XyZ/edit#gid=0
#                                                  ^^^^^^^^^^^ this
MONTH_SOURCES = {
    "August 2026": "1DLQEmPUC0tGWvfmuftXsxnj2d-p8ywQ79mdJQ2o9B14",
    # "September 2026": "PUT_SEPTEMBER_SHEET_ID_HERE",
}

# Tab name inside each monthly sheet.
WORKSHEET_NAME = "SKU POD ENTRY"

# Row holding the real column names. Row 1 is the merged Auto/Manual banner.
HEADER_ROW = 2

OUTPUT_XLSX = "sku_pod_tracker.xlsx"
OUTPUT_HTML = "sku_pod_dashboard.html"

# Parties on open POs: POD wins over GRN for these. Everyone else is GRN first.
# Matching ignores spaces and punctuation, so "METRO C&C" also matches
# "Metro C & C" -- but not "METRO C AND C". Check the spelling in your sheet.
OPEN_PO_PARTIES = ["NB", "METRO C&C", "RELIANCE SIGNATURE", "RELIANCE SMART"]

# Columns the GRN quantity is read from, first non-blank wins.
GRN_PRIORITY = ["NET GRN QTY", "GRN Qty", "Auto GRN Qty"]

# GRN Config = No means the account never produces a GRN, so a blank GRN there
# is not a shortfall. True falls back to POD for those rows.
GRN_CONFIG_NO_FALLS_BACK_TO_POD = True

# Invoices still awaiting GRN are always listed separately. Set True to also
# keep them out of the headline fill rate.
EXCLUDE_PENDING_GRN_FROM_FILL_RATE = False
