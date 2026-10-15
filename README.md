# pod-tracker

Everything that runs against the ops team's monthly **"<Month YYYY> - SKU POD MASTER TRACKER"** Google Sheets
(owner hitesh@; one spreadsheet per month, tab `SKU POD ENTRY` = one row per invoice x item with POD and GRN
status). All three jobs use the same Google service account (`podashboard@…iam.gserviceaccount.com`, secret
`GOOGLE_SERVICE_ACCOUNT`); a new month's sheet is picked up as soon as it is shared with that account.

| job | workflow | what it does | needs on the sheet |
|---|---|---|---|
| dashboard (`sku_pod_tracker.py`, `dashboard.py`, `config.py`) | `refresh.yml`, every 3 h | builds the SKU POD dashboard (GitHub Pages) | Viewer |
| supabase (`pod_sheet_loader.py`) | `refresh.yml`, every 3 h | loads `SKU POD ENTRY` into `public.sku_pod_entry` for Birbal's Order to Cash report; replaces a month when its sheet changed | Viewer |
| grn_tab (`grn_tab_writer.py`) | `grn_tab.yml`, daily 07:15 IST | rewrites the `GRN DATA` tab of the current + previous month's sheets from Supabase | **Editor** (granted to `sheets-loader@…`, secret `GRN_TAB_SERVICE_ACCOUNT`) |

Secrets: `GOOGLE_SERVICE_ACCOUNT` (podashboard@'s key JSON), `GRN_TAB_SERVICE_ACCOUNT` (sheets-loader@'s key JSON, the
account the tracker owner shares as Editor; the grn_tab job falls back to `GOOGLE_SERVICE_ACCOUNT` when it is unset),
`SUPABASE_DB_URL` (session pooler host; runners are IPv4 only, `PGSSLMODE=require`).

## GRN DATA tab (since 2026-10-14)

Until then the ops team pasted the tab every day from a consolidation sheet ("grnpy") built from the per-partner
GRN Google Sheets that the LlamaIndex schedulers write. Every feed also lands in Supabase (one table per partner,
3-hourly), so `grn_tab_writer.py` writes the tab from there and the sheet-side consolidation is retired.

* Layout is unchanged where the tracker reads it: `Part_1` (PO; the standing-PO parties Nature's Basket, Reliance
  Signature / Smart, Metro C&C and Milkbasket get the GRN number instead, so they match by invoice only -- the same PO
  carries many invoices),
  `Part_2` (our invoice number as the partner quotes it), `Part_3` (partner item id), `GRN Qty`, `GRN Amt`,
  `SKU Name` (the tab's own formula, re-written per row, with a fallback through the ERP description), `Source Sheet`
  (now the Supabase feed name). New reference columns: `ERP item`, `ERP description`, `Date`.
* One row per Part_1 x Part_2 x Part_3: exact duplicate lines (a GRN mailed twice) count once, distinct GRNs are
  summed. Column choices per feed were verified against the hand-built tab (`grn_tab_writer.py`, `FEED_SQL`).
* Only lines dated from 21 days before the tracker's month (`GRN_TAB_LOOKBACK_DAYS`) are written: ~60k rows for a
  month instead of the 429k-row full history (the workbook was at Google's 10 M cell ceiling); the freed rows are
  given back to the workbook.
* Rows whose Source is not a Supabase feed (today only the Flipkart Minutes portal export, `FLIPKART QUICK`) are
  carried over unchanged, so a manual paste for such a source survives. `--no-carry` drops them.
* Runs are logged to `public.workflow_logs` as `grn_tab:<Month YYYY>`.

```
python grn_tab_writer.py --dry-run                  # counts per feed, writes nothing
python grn_tab_writer.py --month "October 2026"
python grn_tab_writer.py --sheet-id <id> --month "October 2026"   # a test copy
```

Local runs read `.env` next to the script (`SUPABASE_DB_URL`, `SUPABASE_DB_SSLMODE=verify-full`,
`SUPABASE_DB_SSLROOTCERT`, optionally `SKU_POD_SERVICE_ACCOUNT` or `GRN_TAB_TOKEN_JSON`).
