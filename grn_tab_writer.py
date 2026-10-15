#!/usr/bin/env python3
"""Rewrite the "GRN DATA" tab of the monthly "<Month YYYY> - SKU POD MASTER TRACKER" sheets from Supabase.

Until 2026-10-14 the ops team pasted this tab by hand every day from a consolidation sheet ("grnpy") that was
itself built from the per-partner GRN Google Sheets the LlamaIndex schedulers write. Every one of those feeds
also lands in Supabase (one table per partner, loaded 3-hourly), so this script writes the tab straight from
there and the sheet-side consolidation can be retired.

What the tracker needs from the tab (its only readers are the "SKU POD ENTRY" formulas Auto GRN Qty, col P,
and GRN Amt, col W):
    A Part_1   partner PO number, for every feed. The tracker matches by invoice first and falls back to the PO
               (owner's rule, 2026-10-15). For the standing-PO parties (Nature's Basket, Reliance Signature /
               Smart, Metro C&C, Milkbasket) the same PO carries many invoices, so the fallback can only pick one
               of that PO's receipts -- the receipts that quote no invoice are written first so that it finds
               those. (The old paste had the GRN number here for Nature's Basket.)
    B Part_2   our invoice number as the partner quotes it (blank where the feed has none: Zepto, Blinkit HOT,
               Amazon)
    C Part_3   the partner's item id (SKU code / ASIN / article / EAN / product no; Reliance: item description)
               -- the tab's own SKU Name formula maps it to the ERP name through DATA VALIDATION L:N
    D GRN Qty  accepted / received quantity, summed per (Part_1, Part_2, Part_3)
    E GRN Amt  the feed's line value where it carries one, else 0 (Nature's Basket, Reliance, Milkbasket: 0,
               as before)
    F SKU Name the tracker's own formula, re-written per row (falls back to the ERP description in col I
               through DATA VALIDATION E:F when the partner id is not in the item list)
    G Source   the Supabase feed (zepto_grn, hot_grn, instamart_grn, ...)
    H, I, J    ERP item no + description (Birbal's item map) and the date (GRN date; Zepto: PO date; else the
               day the line was loaded) -- new, for reference only
The matching key (Part_2 + SKU Name, else Part_1 + SKU Name) is unchanged, so the tracker's formulas are not
touched. Exact duplicate lines (the same GRN mailed twice) are counted once, the way the order-to-cash bridge
does; distinct GRNs for one PO x invoice x item are summed, as the old consolidation's SUMIFS did.

Rows are limited to lines dated from LOOKBACK_DAYS before the tracker's month (a GRN cannot precede its
invoice by much), so a month's tab carries ~60-100k rows instead of the 429k the full history had reached
-- the workbook was at 10.2 M cells. Rows whose Source is not one of the feeds below (today: the
"FLIPKART QUICK" portal export, which has no Supabase feed yet) are carried over unchanged, so a manual paste
for such a source survives the rewrite; pass --no-carry to drop them.

Which sheets: every tracker the account can edit for the current month and MONTHS_BACK months before it
(default 1: the ops team keeps editing the previous month as late PODs arrive). A tracker the account can only
read is reported and skipped -- the owner has to share it with edit rights (the same share the POD loader
needs, one role higher).

    python grn_tab_writer.py                        # current + previous month
    python grn_tab_writer.py --month "October 2026"
    python grn_tab_writer.py --sheet-id <id> --month "October 2026"   # e.g. a test copy
    python grn_tab_writer.py --dry-run [--dump rows.json]             # query only

Environment: SUPABASE_DB_URL (+ PGSSLMODE / SUPABASE_DB_SSLMODE, SUPABASE_DB_SSLROOTCERT),
SKU_POD_SERVICE_ACCOUNT (service account key file -- the account the tracker owner shares as Editor: since
2026-10-15 that is sheets-loader@natural-system-459204-j9.iam.gserviceaccount.com, secret GRN_TAB_SERVICE_ACCOUNT;
the POD loader's podashboard@ key only reads) or GRN_TAB_TOKEN_JSON (an OAuth token.json with the spreadsheets +
drive.readonly scopes, used instead when set), GRN_TAB_LOOKBACK_DAYS (default 21), GRN_TAB_MONTHS_BACK (default 1).
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import logging
import os
import re
import socket
import sys
import time

import httplib2
import psycopg2
from google_auth_httplib2 import AuthorizedHttp
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

try:
    from dotenv import load_dotenv
    load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), '.env'), override=False)
except ImportError:
    pass

log = logging.getLogger('grn-tab')

TAB = 'GRN DATA'
NAME_RE = re.compile(r'^\s*([A-Za-z]+ 20\d\d)\s*-\s*SKU POD MASTER TRACKER\s*$', re.I)
HEADER = ['Part_1', 'Part_2', 'Part_3', 'GRN Qty', 'GRN Amt', 'SKU Name', 'Source Sheet', 'ERP item', 'ERP description', 'Date']
NOTE_COL = 'L'                                   # a cell outside the data columns for the "written at" note
LOOKBACK_DAYS = int(os.environ.get('GRN_TAB_LOOKBACK_DAYS', '21'))
MONTHS_BACK = int(os.environ.get('GRN_TAB_MONTHS_BACK', '1'))
CHUNK = 20000                                    # rows per Sheets API write
IST = dt.timezone(dt.timedelta(hours=5, minutes=30))

# Source names the hand-pasted tab used for the feeds below; rows with these are ours and are replaced.
LEGACY_SOURCES = {'Sheet1', 'hotgrn', 'instamartgrn', 'nbgrn', 'hyperpuregrn', 'Amazon', 'bbalertgrn', 'mraws',
                  'reliancegrn', 'mbgrn', 'mrgrn', 'bbalertgrn_2', 'mrdonot', 'zeptogrn', 'amazon', 'nb'}

# One arm per feed. Column choices were checked against the hand-built tab on 2026-10-14 (60 random keys per
# feed: qty and amount reproduced for every feed that carries them). `code` is what Birbal's item map is keyed
# on (ref_platform_item_map.platform_sku_code) -- the same as item except for Reliance, whose Part_3 is the
# description the ops team mapped by hand while the article number keys the item map.
FEED_SQL = """
with raw as (
  select 'zepto_grn' as feed, id, po_no as po, null::text as inv, upper(sku) as item, lower(sku) as code, 'Zepto' as map_platform,
         null::text as grn_no,
         case when po_date ~ '^\\d{1,2} [A-Za-z]+ \\d{4}$' then to_date(po_date, 'FMDD FMMonth YYYY')
              when po_date ~ '^[A-Za-z]{3}-\\d{4}$' then to_date(po_date, 'Mon-YYYY')
              when po_date ~ '^\\d{4}-\\d{2}-\\d{2}' then left(po_date, 10)::date end as grn_date,
         coalesce(grn_quantity, 0) as qty, coalesce(grn_quantity, 0) * coalesce(landing_cost, 0) as amt, created_at::date as loaded
    from public.zepto_grn
  union all
  select 'amazon_grn', id, po_number, nullif(invoice_number, ''), asin, asin, 'Amazon', null, invoice_date,
         coalesce(received_quantity, 0), coalesce(amount_received, 0), created_at::date
    from public.amazon_grn
  union all
  select 'hot_grn', id, po_number, null, item_code, item_code, 'Blinkit', null, null::date,
         coalesce(quantity_grn, 0), coalesce(total_grn_amount, 0), created_at::date
    from public.hot_grn
  union all
  select 'hyperpure_grn', id, po_number, vendor_invoice_number, product_no, product_no, 'Blinkit ZHPL', null, grn_date,
         coalesce(grn_qty, 0), coalesce(amount, 0), created_at::date
    from public.hyperpure_grn
  union all
  select 'instamart_grn', id, po_number, invoice_no, sku_code, sku_code, 'Instamart', grn_no, grn_date,
         coalesce(recv_qty, 0), coalesce(total_inr, 0), created_at::date
    from public.instamart_grn
  union all
  select 'nb_grn', id, po_no, invoice_no, article_code, article_code, 'Nature''s Basket', grn_no, grn_date,
         coalesce(accepted_quantity, 0), 0, created_at::date
    from public.nb_grn
  union all
  select 'bb_alert_grn', id, po_no, invoice_no, sku_code, sku_code, 'BigBasket', grn_no, grn_timestamp,
         coalesce(accepted_quantity, 0), coalesce(accepted_value_incl_tax, 0), created_at::date
    from public.bb_alert_grn
  union all
  select 'bb_net_grn', id, pono, invoiceno, skucode, skucode, 'BigBasket', grnno, invoicedate,
         coalesce(quantity, 0), coalesce(totalvalue, 0), created_at::date
    from public.bb_net_grn
  union all
  select 'mraws_grn', id, po_number, vendor_invoice_number, sku, sku, 'More Retail', null, grn_date,
         coalesce(rcv_qty, 0), coalesce(net_value, 0), created_at::date
    from public.mraws_grn
  union all
  select 'mrgrn_grn', id, po_number, vendor_invoice_number, sku, sku, 'More Retail', null, grn_date,
         coalesce(rcv_qty, 0), coalesce(net_value, 0), created_at::date
    from public.mrgrn_grn
  union all
  -- Reliance Signature / Smart stores (+ Metro C&C on the same feed): item lines live in raw_data. The same
  -- exclusions as the order-to-cash bridge: GRNs that are also in milkbasket_grn, last year's GRNs.
  select 'reliance_grn', g.id, g.po_number, g.vendor_invoice_number, g.raw_data->>'item_description', g.raw_data->>'article', 'Milkbasket',
         g.grn_number, g.grn_date, coalesce((g.raw_data->>'accepted_qty')::numeric, 0), 0, g.created_at::date
    from public.reliance_grn g
   where not exists (select 1 from public.milkbasket_grn mb where mb.grn_number = g.grn_number or mb.vendor_invoice_number = g.vendor_invoice_number)
     and g.grn_date >= '2026-04-01' and coalesce(g.vendor_invoice_number, '') !~ '^26'
  union all
  -- Milkbasket GRN PDFs sometimes list another vendor's lines (Nutri Being / "Thf" products); the old
  -- consolidation kept Mimansa rows only, but the vendor name is blank or "Reliance Retail" on many of ours.
  select 'milkbasket_grn', id, po_number, vendor_invoice_number, article, article, 'Milkbasket', grn_number, grn_date,
         coalesce(accepted_qty, 0), 0, created_at::date
    from public.milkbasket_grn
   where coalesce(vendor_name, '') !~* '^nutri'
  union all
  select 'flipkart_cb_grn', id, po_number, null, item_description, item_description, 'Flipkart', null, expected_delivery_date,
         coalesce(grn_quantity, 0), coalesce(total_amount, 0), created_at::date
    from public.flipkart_cb_grn
  union all
  select 'flipkart_ppr_grn', id, null, coalesce(supplier_invoice_number, invoice_number), description_of_goods, description_of_goods, 'Flipkart', null, invoice_date,
         coalesce(quantity, 0), coalesce(amount, 0), created_at::date
    from public.flipkart_ppr_grn
  union all
  select 'fatema_grn', id, po_number, null, coalesce(fsn_no, item_code), coalesce(fsn_no, item_code), 'Flipkart', bill_no, bill_date,
         coalesce(quantity, 0), coalesce(amount, 0), created_at::date
    from public.fatema_grn
  union all
  select 'slveggies_grn', id, po_number, invoice_number, fsn_number, fsn_number, 'Flipkart', bill_number, bill_date,
         coalesce(grn_quantity, received_quantity, 0), coalesce(amount, 0), created_at::date
    from public.slveggies_grn
),
lines as (
  select feed, id, nullif(trim(po), '') as po, nullif(trim(inv), '') as inv, trim(item) as item, code, map_platform, grn_no, grn_date, qty, amt, loaded,
         row_number() over (partition by feed, po, inv, grn_no, item, grn_date, qty order by id) as dup_seq
    from raw
   where coalesce(nullif(trim(po), ''), nullif(trim(inv), '')) is not null and nullif(trim(item), '') is not null
),
keyed as (
  -- one row per Part_1 (PO) x Part_2 (invoice) x Part_3 (item), the key the old consolidation used; an invoice
  -- received on two GRNs is one row (the tracker's MATCH would otherwise see only the first GRN). The tracker
  -- matches by invoice first and falls back to the PO (owner's rule, 2026-10-15), so Part_1 is the PO for every
  -- feed -- also for the standing-PO parties, where the fallback can only ever pick one of the PO's receipts.
  select feed, po, inv, item,
         sum(qty) as qty, sum(amt) as amt, max(grn_date) as grn_date, min(loaded) as loaded,
         max(coalesce(r.new_no, m.erp_item_no)) as erp_item_no
    from lines l
    left join public.ref_platform_item_map m on m.platform = l.map_platform and lower(m.platform_sku_code) = lower(l.code) and m.erp_item_no is not null
    left join warehouse.item_rename_map r on r.old_no = m.erp_item_no
   where l.dup_seq = 1
     and (%(since)s::date is null or coalesce(l.grn_date, l.loaded) >= %(since)s::date)
   group by feed, po, inv, item
)
select k.feed, k.po, k.inv, k.item, k.qty, k.amt, k.erp_item_no, i."Description" as erp_description, coalesce(k.grn_date, k.loaded) as ref_date
  from keyed k
  left join public.bc_item_card i on i."No" = k.erp_item_no
 -- within a PO the receipts that quote no invoice come first: they are what the PO fallback should find
 order by k.feed, k.po, (k.inv is null) desc, k.inv, k.item
"""

MANAGED_FEEDS = ['zepto_grn', 'amazon_grn', 'hot_grn', 'hyperpure_grn', 'instamart_grn', 'nb_grn', 'bb_alert_grn', 'bb_net_grn',
                 'mraws_grn', 'mrgrn_grn', 'reliance_grn', 'milkbasket_grn', 'flipkart_cb_grn', 'flipkart_ppr_grn', 'fatema_grn', 'slveggies_grn']


# ------------------------------------------------------------------ google
def credentials():
    tok = os.environ.get('GRN_TAB_TOKEN_JSON')
    scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.readonly']
    if tok:
        from google.oauth2.credentials import Credentials
        from google.auth.transport.requests import Request
        c = Credentials.from_authorized_user_file(tok)
        if not c.valid:
            c.refresh(Request())
        return c, 'oauth token %s' % tok
    from google.oauth2 import service_account
    path = os.environ.get('SKU_POD_SERVICE_ACCOUNT', r'C:\Users\tbd20\Downloads\SKU POD\service_account.json')
    c = service_account.Credentials.from_service_account_file(path, scopes=scopes)
    return c, c.service_account_email


HTTP_TIMEOUT = 600                               # a structural change on a 7 M-cell workbook makes Sheets recalculate
                                                 # before it answers; the client library's default 60 s is not enough


def service(name, version, creds):
    return build(name, version, http=AuthorizedHttp(creds, http=httplib2.Http(timeout=HTTP_TIMEOUT)), cache_discovery=False)


def call(req, what='request', attempts=6):
    """Execute a Google API request, retrying the 5xx / 429 the Sheets API throws under load and dropped
    connections. Callers that are not idempotent (the grid trim) re-read the sheet before each attempt."""
    for i in range(attempts):
        try:
            return req.execute()
        except HttpError as e:
            if e.resp.status not in (429, 500, 502, 503, 504) or i == attempts - 1:
                raise
            why = 'HTTP %s' % e.resp.status
        except (socket.timeout, TimeoutError, ConnectionError, httplib2.HttpLib2Error) as e:
            if i == attempts - 1:
                raise
            why = type(e).__name__
        wait = 5 * (2 ** i)
        log.warning('%s: %s, retry in %ds', what, why, wait)
        time.sleep(wait)


def discover(drive):
    """{label: {id, name, modifiedTime, canEdit}} for every tracker the account can see."""
    out, token = {}, None
    while True:
        res = call(drive.files().list(q="mimeType='application/vnd.google-apps.spreadsheet' and trashed=false and name contains 'SKU POD MASTER TRACKER'",
                                      fields='nextPageToken, files(id,name,modifiedTime,capabilities(canEdit))', pageSize=100, pageToken=token,
                                      includeItemsFromAllDrives=True, supportsAllDrives=True), 'drive list')
        for f in res.get('files', []):
            m = NAME_RE.match(f['name'])
            if not m:
                continue
            label = m.group(1).title()
            if label in out:
                raise SystemExit(f"two sheets are named for {label}: {out[label]['id']} and {f['id']} -- rename or unshare one")
            out[label] = {'id': f['id'], 'name': f['name'], 'modifiedTime': f['modifiedTime'], 'canEdit': f.get('capabilities', {}).get('canEdit', False)}
        token = res.get('nextPageToken')
        if not token:
            return out


def month_start(label):
    return dt.datetime.strptime(label, '%B %Y').date()


def wanted_months(today, months_back):
    first = today.replace(day=1)
    out = []
    for _ in range(months_back + 1):
        out.append(first.strftime('%B %Y'))
        first = (first - dt.timedelta(days=1)).replace(day=1)
    return out


# ------------------------------------------------------------------ database
def connect():
    dsn = os.environ['SUPABASE_DB_URL']
    kw = dict(sslmode=os.environ.get('SUPABASE_DB_SSLMODE') or os.environ.get('PGSSLMODE') or 'require')
    if os.environ.get('SUPABASE_DB_SSLROOTCERT'):
        kw['sslrootcert'] = os.environ['SUPABASE_DB_SSLROOTCERT']
    return psycopg2.connect(dsn, **kw)


def fetch(conn, since):
    with conn.cursor() as cur:
        cur.execute("set statement_timeout = '900000'")
        cur.execute(FEED_SQL, {'since': since})
        return cur.fetchall()


def log_run(conn, label, status, rows, msg='', started=None):
    """public.workflow_logs, as every other sheet job writes it (freshness check = one query)."""
    try:
        with conn.cursor() as cur:
            cur.execute("select 1 from information_schema.tables where table_schema='public' and table_name='workflow_logs'")
            if cur.fetchone():
                cur.execute("insert into public.workflow_logs (workflow, source, started_at, ended_at, status, details, rows_written, processed, created_at) "
                            "values (%s, %s, %s, now(), %s, %s::jsonb, %s, %s, now())",
                            ('pod-tracker:grn_tab', 'grn_tab:%s' % label, started or dt.datetime.now(dt.timezone.utc), status,
                             json.dumps({'message': msg[:500], 'rows_written': rows}), rows, rows))
        conn.commit()
    except Exception as e:                                    # noqa: BLE001
        conn.rollback()
        log.warning('workflow_logs not written: %s', str(e)[:100])


# ------------------------------------------------------------------ the tab
_DIGITS = re.compile(r'^[1-9][0-9]{0,14}$|^0$')


def ident(v):
    """Partner ids as the pasted tab had them: an all-digit id becomes a number (DATA VALIDATION's Item No
    column is numeric, and XLOOKUP will not match text to a number); anything else stays text."""
    if v is None:
        return ''
    s = str(v).strip()
    return int(s) if _DIGITS.match(s) else s


def num(v):
    if v is None:
        return 0
    f = float(v)
    return int(f) if f.is_integer() else round(f, 4)


def sku_formula(row):
    r = row
    return ("=IFERROR(IFERROR(XLOOKUP(C{r},'DATA VALIDATION'!L:L,'DATA VALIDATION'!N:N),XLOOKUP(C{r},'DATA VALIDATION'!M:M,'DATA VALIDATION'!N:N)),"
            "IFERROR(XLOOKUP(I{r},'DATA VALIDATION'!E:E,'DATA VALIDATION'!F:F),\"\"))").format(r=r)


def build_rows(db_rows):
    out = []
    for feed, po, inv, item, qty, amt, erp_no, erp_desc, ref_date in db_rows:
        # Part_1 and Part_3 typed like the pasted tab (numeric ids as numbers); the invoice number stays text
        out.append([ident(po), inv or '', ident(item), num(qty), num(amt), '', feed,
                    erp_no or '', erp_desc or '', ref_date.isoformat() if ref_date else ''])
    return out


def read_carry(svc, sid):
    """Rows of the current tab whose Source is not one of ours (e.g. a portal export pasted by hand)."""
    keep, start, step = [], 2, 50000
    managed = set(MANAGED_FEEDS) | LEGACY_SOURCES
    while True:
        res = call(svc.spreadsheets().values().get(spreadsheetId=sid, range=f"'{TAB}'!A{start}:J{start + step - 1}", valueRenderOption='UNFORMATTED_VALUE'), 'read tab')
        vals = res.get('values', [])
        if not vals:
            break
        for r in vals:
            r = list(r) + [''] * (10 - len(r))
            src = str(r[6]).strip()
            if src and src not in managed and (str(r[0]).strip() or str(r[1]).strip() or str(r[2]).strip()):
                inv = r[1] if not isinstance(r[1], float) else ident(r[1])
                keep.append([ident(r[0]), '' if inv is None else inv, ident(r[2]), num(r[3]) if r[3] not in ('', None) else 0,
                             num(r[4]) if r[4] not in ('', None) else 0, '', src, r[7], r[8], r[9] if isinstance(r[9], str) else ''])
        if len(vals) < step:
            break
        start += step
    return keep


def write_tab(svc, sid, rows, note):
    meta = call(svc.spreadsheets().get(spreadsheetId=sid, fields='sheets(properties(sheetId,title,gridProperties))'), 'sheet meta')
    props = {s['properties']['title']: s['properties'] for s in meta.get('sheets', [])}
    if TAB not in props:
        r = call(svc.spreadsheets().batchUpdate(spreadsheetId=sid, body={'requests': [{'addSheet': {'properties': {'title': TAB, 'gridProperties': {'rowCount': 1000, 'columnCount': 12}}}}]}), 'add tab')
        props[TAB] = r['replies'][0]['addSheet']['properties']
    p = props[TAB]
    gid, grid = p['sheetId'], p['gridProperties']
    need_rows, need_cols = len(rows) + 1 + 50, max(12, grid.get('columnCount', 0))
    reqs = []
    if grid['rowCount'] < need_rows:
        reqs.append({'appendDimension': {'sheetId': gid, 'dimension': 'ROWS', 'length': need_rows - grid['rowCount']}})
    if grid.get('columnCount', 0) < need_cols:
        reqs.append({'appendDimension': {'sheetId': gid, 'dimension': 'COLUMNS', 'length': need_cols - grid.get('columnCount', 0)}})
    if reqs:
        call(svc.spreadsheets().batchUpdate(spreadsheetId=sid, body={'requests': reqs}), 'grow grid')
    call(svc.spreadsheets().values().clear(spreadsheetId=sid, range=f"'{TAB}'!A1:Z"), 'clear tab')
    # values (RAW keeps ids exactly as sent), then the SKU Name formulas (USER_ENTERED), in chunks
    call(svc.spreadsheets().values().update(spreadsheetId=sid, range=f"'{TAB}'!A1", valueInputOption='RAW',
                                            body={'values': [HEADER], 'majorDimension': 'ROWS'}), 'header')
    for i in range(0, len(rows), CHUNK):
        chunk = rows[i:i + CHUNK]
        call(svc.spreadsheets().values().update(spreadsheetId=sid, range=f"'{TAB}'!A{i + 2}", valueInputOption='RAW', body={'values': chunk}), f'rows {i}')
        call(svc.spreadsheets().values().update(spreadsheetId=sid, range=f"'{TAB}'!F{i + 2}", valueInputOption='USER_ENTERED',
                                                body={'values': [[sku_formula(i + 2 + j)] for j in range(len(chunk))]}), f'formulas {i}')
        log.info('  %d / %d rows written', min(i + CHUNK, len(rows)), len(rows))
    call(svc.spreadsheets().values().update(spreadsheetId=sid, range=f"'{TAB}'!{NOTE_COL}1", valueInputOption='RAW', body={'values': [[note]]}), 'note')
    reqs = [{'updateSheetProperties': {'properties': {'sheetId': gid, 'gridProperties': {'frozenRowCount': 1}}, 'fields': 'gridProperties.frozenRowCount'}},
            {'repeatCell': {'range': {'sheetId': gid, 'startRowIndex': 0, 'endRowIndex': 1}, 'cell': {'userEnteredFormat': {'textFormat': {'bold': True}}}, 'fields': 'userEnteredFormat.textFormat.bold'}}]
    call(svc.spreadsheets().batchUpdate(spreadsheetId=sid, body={'requests': reqs}), 'format')
    trim_grid(svc, sid, gid, need_rows)
    # read back
    got = call(svc.spreadsheets().values().get(spreadsheetId=sid, range=f"'{TAB}'!A:A", majorDimension='COLUMNS'), 'verify').get('values', [[]])[0]
    if len(got) - 1 != len(rows):
        raise RuntimeError(f'wrote {len(rows)} rows but the tab now has {len(got) - 1}')


def trim_grid(svc, sid, gid, need_rows, attempts=4):
    """Give the rows below the data back to the workbook (it was at Google's cell ceiling). Not idempotent on a
    dropped connection, so the row count is re-read before every attempt and nothing is deleted twice."""
    for i in range(attempts):
        meta = call(svc.spreadsheets().get(spreadsheetId=sid, fields='sheets(properties(sheetId,gridProperties(rowCount)))'), 'grid size')
        row_count = next(s['properties']['gridProperties']['rowCount'] for s in meta['sheets'] if s['properties']['sheetId'] == gid)
        if row_count <= need_rows:
            return
        try:
            call(svc.spreadsheets().batchUpdate(spreadsheetId=sid, body={'requests': [
                {'deleteDimension': {'range': {'sheetId': gid, 'dimension': 'ROWS', 'startIndex': need_rows, 'endIndex': row_count}}}]}), 'trim', attempts=1)
            log.info('  grid trimmed from %d to %d rows', row_count, need_rows)
            return
        except (socket.timeout, TimeoutError, ConnectionError, httplib2.HttpLib2Error, HttpError) as e:
            if i == attempts - 1:
                raise
            log.warning('trim: %s, checking the sheet and retrying', type(e).__name__)
            time.sleep(10 * (i + 1))


# ------------------------------------------------------------------ main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--month', help='only this tracker, e.g. "October 2026" (with --sheet-id: the month that sets the date window)')
    ap.add_argument('--sheet-id', help='write this spreadsheet instead of the trackers found on Drive (needs --month)')
    ap.add_argument('--since', help='override the date window start, YYYY-MM-DD; "all" = every line')
    ap.add_argument('--no-carry', action='store_true', help='do not carry over rows of sources that are not in Supabase')
    ap.add_argument('--dry-run', action='store_true', help='query only, write nothing')
    ap.add_argument('--dump', help='with --dry-run: write the rows to this JSON file')
    ap.add_argument('-v', action='store_true')
    a = ap.parse_args()
    logging.basicConfig(level=logging.DEBUG if a.v else logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
    if a.sheet_id and not a.month:
        sys.exit('--sheet-id needs --month')
    today = dt.datetime.now(IST).date()

    targets = []                                              # (label, sheet_id, since)
    def since_for(label):
        if a.since == 'all':
            return None
        if a.since:
            return dt.date.fromisoformat(a.since)
        return month_start(label) - dt.timedelta(days=LOOKBACK_DAYS)

    svc = None
    if a.sheet_id:
        targets.append((a.month.title(), a.sheet_id, since_for(a.month.title())))
    else:
        months = [a.month.title()] if a.month else wanted_months(today, MONTHS_BACK)
        try:
            cr, who = credentials()
            svc = service('sheets', 'v4', cr)
            found = discover(service('drive', 'v3', cr))
        except Exception as e:                                # noqa: BLE001
            if not a.dry_run:
                raise
            log.warning('no Google credentials (%s): dry run without looking at Drive', str(e)[:120])
            found, who = {}, '?'
        for label in months:
            f = found.get(label)
            if not f:
                log.warning('%s: no tracker named for it is shared with %s', label, who)
                if a.dry_run:
                    targets.append((label, None, since_for(label)))
                continue
            if not f['canEdit']:
                log.error('%s: %s (%s) is shared read-only with %s -- ask the owner for edit rights%s', label, f['name'], f['id'], who,
                          '' if a.dry_run else '; skipped')
                if not a.dry_run:
                    continue
            targets.append((label, f['id'], since_for(label)))
    if not targets:
        sys.exit('nothing to write')

    conn = connect()
    try:
        for label, sid, since in targets:
            started = dt.datetime.now(dt.timezone.utc)
            db_rows = fetch(conn, since)
            rows = build_rows(db_rows)
            per_feed = {}
            for r in rows:
                per_feed[r[6]] = per_feed.get(r[6], 0) + 1
            log.info('%s: %d rows from Supabase (lines dated from %s): %s', label, len(rows), since or 'the beginning',
                     ', '.join(f'{k} {v}' for k, v in sorted(per_feed.items(), key=lambda kv: -kv[1])))
            if a.dry_run:
                if a.dump:
                    json.dump(rows, open(a.dump, 'w'), default=str)
                    log.info('rows dumped to %s', a.dump)
                continue
            if svc is None:
                cr, who = credentials()
                svc = service('sheets', 'v4', cr)
            try:
                carry = [] if a.no_carry else read_carry(svc, sid)
                if carry:
                    srcs = {}
                    for r in carry:
                        srcs[r[6]] = srcs.get(r[6], 0) + 1
                    log.info('%s: carrying over %d rows of sources not in Supabase: %s', label, len(carry), srcs)
                note = ('Rewritten from Supabase %s IST by pod-tracker/grn_tab_writer.py: %d GRN lines dated from %s (%d carried over from sources not in Supabase). '
                        'Do not paste here -- it is replaced every morning.'
                        % (dt.datetime.now(IST).strftime('%d %b %Y %H:%M'), len(rows), since or 'the beginning', len(carry)))
                write_tab(svc, sid, rows + carry, note)
                log.info('%s: written https://docs.google.com/spreadsheets/d/%s', label, sid)
                log_run(conn, label, 'success', len(rows) + len(carry), 'supabase %d, carried %d, since %s' % (len(rows), len(carry), since), started)
            except Exception as e:                            # noqa: BLE001
                log.exception('%s: failed', label)
                log_run(conn, label, 'failed', 0, str(e), started)
                raise
    finally:
        conn.close()


if __name__ == '__main__':
    main()
