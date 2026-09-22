"""
SKU POD sheet loader
====================
Loads the "SKU POD ENTRY" tab of every monthly "<Month YYYY> - SKU POD MASTER TRACKER" Google Sheet into
public.sku_pod_entry. These sheets are the only place POD (proof of delivery) status/qty and the T1/T2 city
tier live; Birbal's warehouse.o2c_report joins them onto the order-to-cash bridge by invoice + item.

The ops team edits the sheets in place (POD qty arrives days after the invoice), so every run REPLACES a
month's rows in one transaction -- unlike the GRN loaders, which are append-only.

Months are discovered by file name among the spreadsheets the service account can see, so a new month is
picked up as soon as its sheet is shared with the service account. A month whose sheet has not changed
since its last load (Drive modifiedTime) is skipped; --force reloads it anyway.

Run:
    python pod_sheet_loader.py                 # every month found
    python pod_sheet_loader.py --month "August 2026"
    python pod_sheet_loader.py --dry-run
Env: SUPABASE_DB_URL, SKU_POD_SERVICE_ACCOUNT (path to the service account json key)
"""
import argparse, csv, datetime as dt, io, json, logging, os, re, sys, time

import psycopg2
from psycopg2.extras import execute_values
from google.oauth2 import service_account
from googleapiclient.discovery import build

try:
    from dotenv import load_dotenv
    load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), '.env'), override=False)
except ImportError:
    pass

log = logging.getLogger('pod-loader')
TAB = 'SKU POD ENTRY'
HEADER_ROW = 2                       # row 1 is the merged Auto/Manual banner
NAME_RE = re.compile(r'^\s*([A-Za-z]+ 20\d\d)\s*-\s*SKU POD MASTER TRACKER\s*$', re.I)
TABLE = 'public.sku_pod_entry'

# sheet header -> (column, type). Headers not listed land only in raw_data.
COLUMNS = {
    'PO. No.': ('po_no', 'text'), 'Warehouse': ('warehouse', 'text'), 'Invoice No.': ('invoice_no', 'text'),
    'PO Date': ('po_date', 'date'), 'Posting Date': ('posting_date', 'date'), 'Display Name': ('display_name', 'text'),
    'Customer location name': ('customer_location_name', 'text'), 'Ship-to City': ('ship_to_city', 'text'),
    'Item/Account': ('item_no', 'text'), 'Description': ('description', 'text'),
    'Parent Descripition': ('parent_description', 'text'), 'Is Cancel': ('is_cancel', 'text'),
    'GRN Config': ('grn_config', 'text'), 'PO Qty': ('po_qty', 'numeric'), 'Invoice Qty.': ('invoice_qty', 'numeric'),
    'Auto GRN Qty': ('auto_grn_qty', 'numeric'), 'Auto GDN QTY': ('auto_gdn_qty', 'numeric'), 'GRN Qty': ('grn_qty', 'numeric'),
    'GDN QTY': ('gdn_qty', 'numeric'), 'NET GRN QTY': ('net_grn_qty', 'numeric'), 'Invoice Amt': ('invoice_amt', 'numeric'),
    'PO Amt': ('po_amt', 'numeric'), 'GRN Amt': ('grn_amt', 'numeric'), 'GDN amt': ('gdn_amt', 'numeric'),
    'POD Qty': ('pod_qty', 'numeric'), 'Inv - POD Diff': ('inv_pod_diff', 'numeric'), 'POD Status': ('pod_status', 'text'),
    'GRN Status': ('grn_status', 'text'), 'Delivery/GRN Date': ('delivery_grn_date', 'date'),
    'POD & GRN Qty Diff.': ('pod_grn_qty_diff', 'numeric'), 'T1/T2': ('tier', 'text'), 'RTV DOC': ('rtv_doc', 'text'),
    'Return QTY': ('return_qty', 'numeric'), 'Return Reason': ('return_reason', 'text'), 'CN AMOUNT': ('cn_amount', 'numeric'),
    'Unit Price': ('unit_price', 'numeric'), 'SRO No': ('sro_no', 'text'), 'GDN %': ('gdn_pct', 'numeric'),
    'Concat': ('concat_key', 'text'), 'CATAGORY': ('category', 'text'),
}
COLS = [c for c, _ in COLUMNS.values()]
TYPES = dict(COLUMNS.values())

DDL = f"""
create table if not exists {TABLE} (
    id bigint generated always as identity primary key,
    month text not null,                 -- "August 2026", from the sheet name
    month_start date not null,
    sheet_id text not null,
    sheet_row integer not null,
    {', '.join(f'{c} {t}' for c, t in COLUMNS.values())},
    raw_data jsonb not null,
    loaded_at timestamptz not null default now()
);
alter table {TABLE} add column if not exists sheet_modified timestamptz;   -- the sheet's Drive modifiedTime at load
create index if not exists sku_pod_entry_inv_item on {TABLE} (invoice_no, item_no);
create index if not exists sku_pod_entry_month on {TABLE} (month_start);
alter table {TABLE} enable row level security;
revoke all on {TABLE} from anon, authenticated;
comment on table {TABLE} is 'LIVE POD. "SKU POD ENTRY" tab of the monthly SKU POD MASTER TRACKER Google Sheets (ops team, edited in place). One row per invoice x item; each load REPLACES the month. Loaded by pod_sheet_loader.py. Source of POD status/qty and the T1/T2 tier for warehouse.o2c_report.';
"""


def sa_credentials():
    path = os.environ.get('SKU_POD_SERVICE_ACCOUNT', r'C:\Users\tbd20\Downloads\SKU POD\service_account.json')
    return service_account.Credentials.from_service_account_file(
        path, scopes=['https://www.googleapis.com/auth/spreadsheets.readonly', 'https://www.googleapis.com/auth/drive.readonly'])


def discover(drive):
    months, token = {}, None
    while True:
        res = drive.files().list(q="mimeType='application/vnd.google-apps.spreadsheet' and trashed=false and name contains 'SKU POD MASTER TRACKER'",
                                 fields='nextPageToken, files(id,name,modifiedTime)', pageSize=100, pageToken=token,
                                 includeItemsFromAllDrives=True, supportsAllDrives=True).execute()
        for f in res['files']:
            m = NAME_RE.match(f['name'])
            if not m: continue
            label = m.group(1).title()
            if label in months:
                raise SystemExit(f"two sheets are named for {label}: {months[label]['id']} and {f['id']} -- rename or unshare one")
            months[label] = f
        token = res.get('nextPageToken')
        if not token: return months


def serial_date(v):
    if v in (None, ''): return None
    if isinstance(v, (int, float)):
        if 20000 < v < 80000: return dt.date(1899, 12, 30) + dt.timedelta(days=int(v))
        return None
    s = str(v).strip()
    for f in ('%d-%b-%y', '%d-%m-%Y', '%d-%b-%Y', '%d/%m/%Y', '%Y-%m-%d', '%d-%m-%y', '%d/%m/%y'):
        try: return dt.datetime.strptime(s, f).date()
        except ValueError: pass
    return None


def number(v):
    if v in (None, ''): return None
    if isinstance(v, (int, float)): return v
    s = str(v).replace(',', '').strip()
    try: return float(s)
    except ValueError: return None


def text(v):
    if v in (None, ''): return None
    if isinstance(v, float) and v.is_integer(): v = int(v)      # PO numbers typed as numbers must not become 2.87E+12
    s = str(v).strip()
    return s or None


CONVERT = {'text': text, 'numeric': number, 'date': serial_date}


def read_month(sheets, sheet_id):
    for attempt in range(4):
        try:
            return sheets.spreadsheets().values().get(spreadsheetId=sheet_id, range=f"'{TAB}'",
                                                      valueRenderOption='UNFORMATTED_VALUE').execute().get('values', [])
        except Exception as e:                                       # noqa: BLE001
            if attempt == 3: raise
            log.warning('read failed (%s), retrying', str(e)[:80]); time.sleep(5 * (attempt + 1))


def rows_for(label, sheet_id, values):
    headers = [str(h).strip() for h in values[HEADER_ROW - 1]]
    missing = [h for h in ('Invoice No.', 'Item/Account', 'POD Qty', 'POD Status') if h not in headers]
    if missing: raise SystemExit(f'{label}: header row {HEADER_ROW} lacks {missing}')
    idx = {h: i for i, h in reversed(list(enumerate(headers))) if h}   # first occurrence wins on duplicate headers
    month_start = dt.datetime.strptime(label, '%B %Y').date()
    out = []
    for n, r in enumerate(values[HEADER_ROW:], start=HEADER_ROW + 1):
        get = lambda h: r[idx[h]] if h in idx and idx[h] < len(r) else None
        inv = text(get('Invoice No.'))
        if not inv: continue                                          # blank / formula-only rows
        raw = {h: r[i] for h, i in idx.items() if i < len(r) and r[i] not in (None, '')}
        vals = [CONVERT[t](get(h)) for h, (_, t) in COLUMNS.items()]
        out.append([label, month_start, sheet_id, n] + vals + [json.dumps(raw, default=str, ensure_ascii=False)])
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--month', help='load only this month, e.g. "August 2026"')
    ap.add_argument('--dry-run', action='store_true')
    ap.add_argument('--force', action='store_true', help='reload months whose sheet has not changed since the last load')
    a = ap.parse_args()
    logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
    cr = sa_credentials()
    drive = build('drive', 'v3', credentials=cr, cache_discovery=False)
    sheets = build('sheets', 'v4', credentials=cr, cache_discovery=False)
    months = discover(drive)
    if a.month: months = {k: v for k, v in months.items() if k.lower() == a.month.lower()}
    if not months: raise SystemExit('no SKU POD MASTER TRACKER sheet visible to the service account')
    log.info('months: %s', ', '.join(sorted(months, key=lambda m: dt.datetime.strptime(m, '%B %Y'))))

    conn = None if a.dry_run else psycopg2.connect(os.environ['SUPABASE_DB_URL'])
    if conn:
        with conn.cursor() as cur:
            cur.execute("set statement_timeout = '1800000'")
            cur.execute(DDL)
        conn.commit()
    loaded = {}
    if conn:
        with conn.cursor() as cur:
            cur.execute(f'select month, sheet_id, max(sheet_modified) from {TABLE} group by 1, 2')
            loaded = {(m, sid): ts for m, sid, ts in cur.fetchall()}
    total = 0
    for label, f in sorted(months.items(), key=lambda kv: dt.datetime.strptime(kv[0], '%B %Y')):
        modified = dt.datetime.fromisoformat(f['modifiedTime'].replace('Z', '+00:00'))
        prev = loaded.get((label, f['id']))
        if prev is not None and prev >= modified and not a.force:
            log.info('%s: unchanged since %s, skipped', label, prev.isoformat())
            continue
        rows = rows_for(label, f['id'], read_month(sheets, f['id']))
        log.info('%s: %d rows (sheet modified %s)', label, len(rows), f['modifiedTime'])
        if conn is None or not rows: continue
        buf = io.StringIO()
        wr = csv.writer(buf)
        for r in rows:
            wr.writerow(['' if v is None else v for v in r] + [modified.isoformat()])
        buf.seek(0)
        with conn.cursor() as cur:                                    # replace the month atomically
            cur.execute(f'delete from {TABLE} where month = %s', (label,))
            cur.copy_expert(f"copy {TABLE} (month, month_start, sheet_id, sheet_row, {', '.join(COLS)}, raw_data, sheet_modified) "
                            f"from stdin with (format csv, null '')", buf)
        conn.commit()
        total += len(rows)
    if conn:
        log.info('%s: %d rows loaded', TABLE, total)
        conn.close()


if __name__ == '__main__':
    sys.exit(main())
