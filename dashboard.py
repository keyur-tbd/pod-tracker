"""
Builds a standalone HTML dashboard from the tables produced by
sku_pod_tracker.py. All data is embedded, so the file works offline and can be
mailed or dropped on a shared drive.

Two tabs run live off row-level data and respond to the filter bar:
Fill rate and POD summary. The rest are pre-aggregated tables.
"""

import json
import datetime as dt
import numpy as np
import pandas as pd

MAX_ROWS_EMBEDDED = 5000

PCT_HINTS = ("_pct", "pct_")
MONEY_HINTS = ("value", "amt", "amount", "price")
QTY_HINTS = ("qty", "invoices", "lines", "rows", "days", "count")


def col_kind(name: str) -> str:
    n = str(name).lower()
    if any(h in n for h in PCT_HINTS):
        return "pct"
    if any(h in n for h in MONEY_HINTS):
        return "money"
    if any(h in n for h in QTY_HINTS):
        return "num"
    return "text"


def frame_to_payload(df: pd.DataFrame, label: str, note: str = "") -> dict:
    if df is None:
        df = pd.DataFrame()
    d = df.copy()
    if len(d) > MAX_ROWS_EMBEDDED:
        note = (note + " " if note else "") + (
            f"Showing the first {MAX_ROWS_EMBEDDED:,} of {len(d):,} rows to keep this "
            f"file openable. The complete table is in the Excel workbook.")
        d = d.head(MAX_ROWS_EMBEDDED)
    for c in d.columns:
        if pd.api.types.is_datetime64_any_dtype(d[c]):
            d[c] = d[c].dt.strftime("%d-%b-%y")
        elif isinstance(d[c].dtype, pd.CategoricalDtype):
            d[c] = d[c].astype(str)
        elif d[c].dtype == bool:
            d[c] = d[c].map({True: "Yes", False: "No"})
    d = d.replace({np.nan: None, pd.NaT: None})
    rows = []
    for rec in d.itertuples(index=False, name=None):
        rows.append([None if (isinstance(v, float) and pd.isna(v)) else
                     (float(v) if isinstance(v, np.floating) else
                      int(v) if isinstance(v, np.integer) else v)
                     for v in rec])
    cols = [str(c) for c in d.columns]
    return {"label": label, "note": note, "columns": cols,
            "kinds": [col_kind(c) for c in cols], "rows": rows}


DIM_COLUMNS = {
    "warehouse": "Warehouse",
    "platform": "Display Name",
    "category": "CATAGORY",
    "city": "Ship-to City",
    "tier": "T1/T2",
    "month": "month",
    "location": "Customer location name",
    "sku": "Description",
    "item": "Item/Account",
    "invoice": "Invoice No.",
    "po": "PO. No.",
}

# Bit positions packed into one integer per row, to keep the payload small.
FLAG_BITS = [
    ("cfg", "grn_applicable"),
    ("pend", "grn_pending"),
    ("openpo", "open_po"),
    ("haspod", None),
    ("hasgrn", None),
    ("c24", "cancel_24"),
    ("cpost", "cancel_post"),
    ("podcanc", "pod_cancel_unflagged"),
]


def build_row_dataset(df: pd.DataFrame) -> dict:
    """Row-level data, dictionary encoded, so the browser can re-aggregate
    under any combination of filters instead of showing pre-baked totals."""
    d = df
    dims, codes = {}, {}
    for key, src in DIM_COLUMNS.items():
        col = d[src].astype(str) if src in d.columns else pd.Series([""] * len(d), index=d.index)
        cat = pd.Categorical(col.fillna(""))
        dims[key] = [str(x) for x in cat.categories]
        codes[key] = [int(c) for c in cat.codes]

    for key, src in [("date", "Posting Date"), ("podate", "PO Date")]:
        txt = d[src].dt.strftime("%Y-%m-%d") if src in d.columns else pd.Series([""] * len(d))
        txt = txt.where(txt.notna(), "")
        cat = pd.Categorical(txt)
        dims[key] = [str(x) for x in cat.categories]
        codes[key] = [int(c) for c in cat.codes]

    flags = np.zeros(len(d), dtype=np.int32)
    haspod = d["POD Qty"].notna().to_numpy()
    hasgrn = (d["grn_qty"].fillna(0) > 0).to_numpy()
    for i, (name, col) in enumerate(FLAG_BITS):
        if name == "haspod":
            v = haspod
        elif name == "hasgrn":
            v = hasgrn
        else:
            v = d[col].fillna(False).to_numpy().astype(bool)
        flags |= (v.astype(np.int32) << i)

    def nums(series, dp):
        return [round(float(v), dp) for v in series.fillna(0).to_numpy()]

    return {
        "dims": dims,
        "codes": codes,
        "po": nums(d["PO Qty"], 2),
        "inv": nums(d["Invoice Qty."], 2),
        "grn": nums(d["received_qty"], 2),
        "price": nums(d["Unit Price"], 2),
        "amt": nums(d["Invoice Amt"], 2),
        "shortq": nums(d["short_qty"], 2),
        "flags": [int(f) for f in flags],
        "n": int(len(d)),
        "bits": [n for n, _ in FLAG_BITS],
    }


def build_payload(results_by_month: dict) -> dict:
    payload = {"generated": dt.datetime.now().strftime("%d %b %Y, %H:%M"), "months": {}}
    for month, r in results_by_month.items():
        tables = {
            "claims": [
                frame_to_payload(r["claims_detail"], "Open claims ledger",
                                 "Short invoices with an SRO or credit memo still to raise."),
                frame_to_payload(r["claims_summary"], "SRO and CN status"),
                frame_to_payload(r["claims_ageing"], "Ageing of uncredited shorts"),
            ],
            "cancel": [
                frame_to_payload(r["cancel_flags"], "Cancellation flags",
                                 "Within 24 hours is Is Cancel = Yes. After 24 hours is "
                                 "Return Reason = INC. The last row needs a human look."),
                frame_to_payload(r["cancel_detail"], "Cancelled invoices"),
            ],
            "exceptions": [frame_to_payload(v, k) for k, v in r["exceptions"].items()],
        }
        payload["months"][month] = {"headline": r["headline"], "tables": tables,
                                    "grnstate": frame_to_payload(
                                        r["exceptions"].get("GRN state by platform"),
                                        "GRN state by platform")}
    return payload


def build_dashboard(results_by_month: dict, path: str, line_df: pd.DataFrame = None):
    payload = build_payload(results_by_month)
    payload["rows"] = build_row_dataset(line_df) if line_df is not None else None
    html = HTML.replace("__PAYLOAD__", json.dumps(payload, default=str))
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
    return path


HTML = """<!doctype html>
<html lang="en"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex, nofollow, noarchive">
<title>SKU POD Tracker</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;600&family=IBM+Plex+Sans+Condensed:wght@600;700&family=IBM+Plex+Sans:wght@400;500;600&display=swap" rel="stylesheet">
<script src="https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js"></script>
<style>
:root{
  --ink:#0E1417; --paper:#ECEFEC; --surface:#FFFFFF; --rule:#D2D7D1; --soft:#F5F7F4;
  --muted:#66716B; --risk:#A32218; --wait:#8A5A00; --clear:#0E6455;
  --sans:'IBM Plex Sans',system-ui,sans-serif;
  --cond:'IBM Plex Sans Condensed','IBM Plex Sans',system-ui,sans-serif;
  --mono:'IBM Plex Mono',ui-monospace,monospace;
}
*{box-sizing:border-box}
body{margin:0;background:var(--paper);color:var(--ink);font-family:var(--sans);font-size:14px;line-height:1.45}
.wrap{max-width:1560px;margin:0 auto;padding:0 22px 72px}
h1{font-family:var(--cond);font-weight:700;font-size:25px;letter-spacing:-0.01em;margin:0;text-transform:uppercase}
.sub{font-family:var(--mono);font-size:10.5px;color:var(--muted);margin-top:3px;letter-spacing:0.03em}

header{border-bottom:1px solid var(--ink);background:var(--paper);position:sticky;top:0;z-index:40}
.topbar{max-width:1560px;margin:0 auto;padding:16px 22px 12px;display:flex;flex-wrap:wrap;gap:14px;align-items:flex-end;justify-content:space-between}
select,button,input{font-family:var(--mono);font-size:12px;border:1px solid var(--ink);background:var(--surface);color:var(--ink);border-radius:0}
select,button{padding:7px 11px;cursor:pointer}
button:hover,select:hover{background:var(--ink);color:var(--paper)}
button:focus-visible,select:focus-visible,input:focus-visible{outline:2px solid var(--ink);outline-offset:2px}
.ghost{border-color:var(--rule);color:var(--muted)}

/* --- filter bar --- */
.filterbar{border-bottom:1px solid var(--rule);background:var(--soft)}
.fbin{max-width:1560px;margin:0 auto;padding:11px 22px 13px}
.frow{display:flex;flex-wrap:wrap;gap:9px;align-items:center}
.search{position:relative;flex:1 1 280px;min-width:220px}
.search input{width:100%;padding:8px 11px 8px 30px;border-color:var(--rule);background:var(--surface);font-size:12.5px}
.search .ic{position:absolute;left:10px;top:8px;color:var(--muted);font-size:12px}
.drop{position:relative}
.drop>button{border-color:var(--rule);background:var(--surface);display:inline-flex;gap:7px;align-items:center;white-space:nowrap}
.drop>button.active{border-color:var(--ink);background:var(--ink);color:var(--paper)}
.drop .cnt{font-size:10px;padding:0 5px;border:1px solid currentColor;line-height:15px}
.menu{display:none;position:absolute;top:calc(100% + 4px);left:0;z-index:60;background:var(--surface);border:1px solid var(--ink);min-width:250px;max-width:340px;padding:9px;box-shadow:0 6px 22px rgba(14,20,23,.13)}
.menu.open{display:block}
.menu input{width:100%;padding:6px 8px;border-color:var(--rule);margin-bottom:7px;font-size:11.5px}
.opts{max-height:230px;overflow:auto}
.opt{display:flex;gap:8px;align-items:center;padding:4px 5px;cursor:pointer;font-size:12px}
.opt:hover{background:var(--soft)}
.opt .box{width:12px;height:12px;border:1px solid var(--muted);flex:none;position:relative}
.opt.on .box{background:var(--ink);border-color:var(--ink)}
.opt.on .box:after{content:"";position:absolute;left:3px;top:0;width:4px;height:8px;border:solid var(--paper);border-width:0 1.5px 1.5px 0;transform:rotate(42deg)}
.opt .nm{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.menu .foot{display:flex;justify-content:space-between;margin-top:8px;padding-top:7px;border-top:1px solid var(--rule);font-size:11px}
.menu .foot a{cursor:pointer;text-decoration:underline;color:var(--muted)}
.seg{display:inline-flex;border:1px solid var(--rule);background:var(--surface)}
.seg span{padding:7px 10px;cursor:pointer;font-family:var(--mono);font-size:11px;border-right:1px solid var(--rule)}
.seg span:last-child{border-right:none}
.seg span.on{background:var(--ink);color:var(--paper)}
.seg em{font-style:normal;padding:7px 9px;color:var(--muted);font-family:var(--mono);font-size:10px;letter-spacing:.06em;text-transform:uppercase;background:var(--soft);border-right:1px solid var(--rule)}
.active-line{margin-top:9px;font-size:11.5px;color:var(--muted);display:flex;flex-wrap:wrap;gap:7px;align-items:center}
.tag{background:var(--surface);border:1px solid var(--rule);padding:2px 7px;font-family:var(--mono);font-size:10.5px}
.tag b{font-weight:500;color:var(--ink)}

/* --- kpi --- */
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(172px,1fr));gap:1px;background:var(--rule);border:1px solid var(--rule);margin-top:22px}
.kpi{background:var(--surface);padding:13px 15px 14px}
.kpi .k{font-family:var(--mono);font-size:9.5px;letter-spacing:0.09em;text-transform:uppercase;color:var(--muted)}
.kpi .v{font-family:var(--mono);font-size:22px;font-weight:600;font-variant-numeric:tabular-nums;margin-top:4px;letter-spacing:-0.02em}
.kpi .d{font-size:11px;color:var(--muted);margin-top:2px}
.kpi.risk .v{color:var(--risk)} .kpi.wait .v{color:var(--wait)} .kpi.clear .v{color:var(--clear)}

.fgrid{display:grid;grid-template-columns:repeat(auto-fit,minmax(212px,1fr));gap:1px;background:var(--rule);border:1px solid var(--rule);margin-bottom:20px}
.fk{background:var(--surface);padding:15px 17px 16px}
.fk .k{font-family:var(--mono);font-size:9.5px;letter-spacing:0.09em;text-transform:uppercase;color:var(--muted)}
.fk .v{font-family:var(--mono);font-size:29px;font-weight:600;font-variant-numeric:tabular-nums;letter-spacing:-0.03em;margin:5px 0 2px}
.fk .m{height:5px;background:#E5E9E4;margin:7px 0 8px}
.fk .m i{display:block;height:100%;transition:width .6s cubic-bezier(.22,.7,.28,1)}
.fk .d{font-size:11px;color:var(--muted)}
.fk .loss{font-family:var(--mono);font-size:13px;color:var(--risk);font-variant-numeric:tabular-nums}

/* --- tabs --- */
nav.tabs{display:flex;margin:26px 0 0;border-bottom:1px solid var(--ink);flex-wrap:wrap}
.tab{font-family:var(--cond);font-weight:600;font-size:13.5px;letter-spacing:0.05em;text-transform:uppercase;padding:9px 15px;border:1px solid transparent;border-bottom:none;background:none;cursor:pointer;color:var(--muted)}
.tab[aria-selected="true"]{background:var(--surface);border-color:var(--ink);color:var(--ink);margin-bottom:-1px}

/* --- cards & tables --- */
.panel{display:none;padding-top:20px}
.panel.on{display:block}
.card{background:var(--surface);border:1px solid var(--rule);margin-bottom:20px}
.card-h{display:flex;flex-wrap:wrap;gap:10px;align-items:center;justify-content:space-between;padding:12px 15px;border-bottom:1px solid var(--rule)}
.card-h h3{font-family:var(--cond);font-size:14.5px;letter-spacing:0.04em;text-transform:uppercase;margin:0}
.card-h .note{font-size:11px;color:var(--muted);flex-basis:100%;margin-top:-3px}
.tools{display:flex;gap:7px;align-items:center}
input[type=search]{padding:6px 9px;border-color:var(--rule);background:var(--soft);width:168px}
.scroll{overflow:auto;max-height:540px}
table{border-collapse:collapse;width:100%;font-size:12.5px}
th{position:sticky;top:0;background:var(--surface);font-family:var(--mono);font-size:9.5px;letter-spacing:0.05em;text-transform:uppercase;color:var(--muted);font-weight:500;text-align:left;padding:9px 11px;border-bottom:1px solid var(--ink);cursor:pointer;white-space:nowrap;z-index:2}
th:hover{color:var(--ink)}
td{padding:7px 11px;border-bottom:1px solid var(--rule);white-space:nowrap}
tbody tr:hover{background:var(--soft)}
td.n{text-align:right;font-family:var(--mono);font-variant-numeric:tabular-nums}
td.doc{font-family:var(--mono);font-size:11.5px}
td.neg{color:var(--risk)}
.pill{font-family:var(--mono);font-size:10px;letter-spacing:0.05em;text-transform:uppercase;padding:2px 7px;border:1px solid currentColor}
.pill.p{color:var(--risk)} .pill.r{color:var(--clear)}
.empty{padding:24px 15px;color:var(--muted);font-size:13px}
.notice{font-size:11.5px;color:var(--muted);margin:-6px 0 16px}

/* --- charts --- */
.charts{display:grid;grid-template-columns:repeat(auto-fit,minmax(330px,1fr));gap:1px;background:var(--rule);border:1px solid var(--rule);margin-bottom:20px}
.chart{background:var(--surface);padding:15px 17px 17px;min-width:0}
.chart h4{font-family:var(--cond);font-size:12.5px;letter-spacing:0.07em;text-transform:uppercase;margin:0 0 3px;font-weight:600}
.chart .sub2{font-size:11px;color:var(--muted);margin-bottom:13px}
.brow{display:grid;grid-template-columns:minmax(70px,26%) 1fr auto;gap:10px;align-items:center;margin-bottom:8px;font-size:12px}
.brow .lbl{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.brow .track2{height:15px;background:#E5E9E4;position:relative;min-width:20px}
.brow .track2 i{position:absolute;left:0;top:0;height:100%;transition:width .6s cubic-bezier(.22,.7,.28,1)}
.brow .amt{font-family:var(--mono);font-variant-numeric:tabular-nums;font-size:12px;white-space:nowrap}
.f-ink{background:var(--ink)} .f-risk{background:var(--risk)} .f-wait{background:var(--wait)}
.f-clear{background:var(--clear)} .f-mute{background:#98A39E}
.stack{display:flex;height:32px;border:1px solid var(--ink);margin:2px 0 11px}
.stack span{height:100%}
.legend{display:flex;flex-wrap:wrap;gap:12px;font-size:11px;color:var(--muted)}
.legend b{display:inline-block;width:9px;height:9px;margin-right:5px}
.chart .none{font-size:12px;color:var(--muted);padding:8px 0}

/* --- matrix / heatmap --- */
.hm{overflow:auto;max-height:600px}
.hm table{font-size:11.5px}
.hm th{text-align:center;padding:7px 6px;font-size:9.5px;z-index:3}
.hm td{padding:0;border:none}
.hm td.cell{font-family:var(--mono);font-variant-numeric:tabular-nums;text-align:right;padding:6px 9px;border-right:1px solid rgba(255,255,255,.5);border-bottom:1px solid rgba(255,255,255,.5);white-space:nowrap;font-size:11px}
.hm td.rowh,.hm td.grp{position:sticky;background:var(--surface);z-index:2;padding:6px 11px;border-bottom:1px solid var(--rule);white-space:nowrap}
.hm td.grp{left:0;min-width:118px;font-family:var(--cond);font-weight:600;text-transform:uppercase;letter-spacing:0.04em;border-top:1px solid var(--ink)}
.hm td.rowh{left:0;min-width:118px}
.hm td.met{position:sticky;left:118px;background:var(--surface);z-index:2;padding:6px 11px;border-bottom:1px solid var(--rule);white-space:nowrap;font-size:11.5px}
.hm td.tot{font-weight:600;border-left:1px solid var(--ink)}

/* --- POD summary --- */
.pods{overflow:auto;max-height:640px}
.pods table{font-size:12px}
.pods th.grp1{text-align:center;background:var(--soft);border-bottom:1px solid var(--rule);border-left:1px solid var(--ink);font-family:var(--cond);font-size:10.5px;letter-spacing:.06em;color:var(--ink);text-transform:uppercase;padding:7px 9px}
.pods th.sub{text-align:right;font-size:9.5px;padding:6px 9px;white-space:nowrap}
.pods th.sub.first{border-left:1px solid var(--ink)}
.pods td.wh{position:sticky;left:0;background:var(--surface);z-index:2;font-family:var(--mono);font-size:11.5px;border-bottom:1px solid var(--rule)}
.pods td.v{text-align:right;font-family:var(--mono);font-variant-numeric:tabular-nums;font-size:11.5px}
.pods td.v.first{border-left:1px solid var(--rule)}
.pods tr.total td{border-top:1px solid var(--ink);font-weight:600;background:var(--soft)}
.pods td.flag{color:var(--risk);font-weight:600}
footer{margin-top:38px;padding-top:15px;border-top:1px solid var(--rule);font-size:11.5px;color:var(--muted);line-height:1.75}
footer code{font-family:var(--mono);font-size:11px}
@media (max-width:760px){.wrap{padding:0 13px 48px}.topbar{padding:13px}h1{font-size:19px}}
@media (prefers-reduced-motion:reduce){*{transition:none!important}}
</style></head><body>

<header>
  <div class="topbar">
    <div><h1>SKU POD Tracker</h1><div class="sub" id="stamp"></div></div>
    <div style="display:flex;gap:8px;align-items:center">
      <select id="month" aria-label="Month"></select>
      <button id="dl-xlsx">Download workbook</button>
    </div>
  </div>
  <div class="filterbar"><div class="fbin">
    <div class="frow" id="frow"></div>
    <div class="active-line" id="activeline"></div>
  </div></div>
</header>

<div class="wrap">
  <nav class="tabs" id="tabs" role="tablist"></nav>
  <div id="panels"></div>
  <footer id="notes"></footer>
</div>

<script>
const DATA = __PAYLOAD__;
const R = DATA.rows;
const NL = String.fromCharCode(10);
const SECTIONS = [["fill","Fill rate"],["pod","POD summary"],["claims","Claims"],
                  ["cancel","Cancellations"],["exceptions","Exceptions"]];
const LIVE = {fill:1, pod:1};
let month = Object.keys(DATA.months)[0];
let tab = "fill";

const inr = v => v==null ? "" : "\\u20B9" + Math.round(v).toLocaleString("en-IN");
const num = v => v==null ? "" : Number(v).toLocaleString("en-IN",{maximumFractionDigits:2});
const pct = v => (v==null||isNaN(v)) ? "\\u2014" : Number(v).toFixed(2) + "%";
function fmt(v,kind){
  if(v===null||v===undefined||v==="") return kind==="text" ? "" : "\\u2014";
  if(kind==="money") return inr(v);
  if(kind==="pct")   return pct(v);
  if(kind==="num")   return num(v);
  return String(v);
}
function fmtDay(iso){
  if(!iso) return "";
  const p=iso.split("-"), M=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  return p[2]+" "+M[(+p[1])-1];
}
function saveBlob(blob,name){
  const a=document.createElement("a"); a.href=URL.createObjectURL(blob); a.download=name;
  document.body.appendChild(a); a.click(); a.remove(); setTimeout(()=>URL.revokeObjectURL(a.href),1000);
}
function csvEsc(v){
  if(v==null) return "";
  const s=String(v);
  return (s.indexOf(",")>=0||s.indexOf(String.fromCharCode(34))>=0||s.indexOf(NL)>=0)
    ? String.fromCharCode(34)+s.split(String.fromCharCode(34)).join(String.fromCharCode(34,34))+String.fromCharCode(34) : s;
}
function toCsv(head,rows){ return [head.map(csvEsc).join(",")].concat(
  rows.map(r=>r.map(csvEsc).join(","))).join(NL); }

/* ---------- colour ---------- */
function mix(a,b,t){
  const h=x=>[parseInt(x.slice(1,3),16),parseInt(x.slice(3,5),16),parseInt(x.slice(5,7),16)];
  const A=h(a),B=h(b);
  return "rgb("+A.map((v,i)=>Math.round(v+(B[i]-v)*t)).join(",")+")";
}
function heat(v){
  if(v==null||isNaN(v)) return "#F1F3F0";
  if(v<=70) return mix("#8E1B12","#C0392B",Math.max(0,v)/70);
  if(v<=88) return mix("#C0392B","#D9A404",(v-70)/18);
  if(v<=96) return mix("#D9A404","#5E9B3F",(v-88)/8);
  return mix("#5E9B3F","#12734F",Math.min(1,(v-96)/4));
}
const heatInk = v => (v==null||isNaN(v)||v<=88||v>96) ? "#FFFFFF" : "#1A1A1A";

/* ---------- filters ---------- */
const BIT = {};
if(R) R.bits.forEach((b,i)=>BIT[b]=1<<i);
const has = (i,b) => (R.flags[i] & BIT[b]) !== 0;

const FILTERS = [["warehouse","Warehouse"],["platform","Display Name"],["category","Category"],
                 ["city","City"],["tier","T1/T2"],["month","Month"]];
const sel = {}; FILTERS.forEach(([k])=>sel[k]=new Set());
let cfgMode="yes", pendMode="exclude", q="", openMenu=null;

function matchQ(i){
  if(!q) return true;
  const t=q.toLowerCase();
  const fields=["invoice","po","date","podate"];
  for(const f of fields){
    const v=R.dims[f][R.codes[f][i]];
    if(v && v.toLowerCase().indexOf(t)>=0) return true;
  }
  return false;
}
function rowIdx(opts){
  opts = opts || {};
  const out=[], n=R.n;
  for(let i=0;i<n;i++){
    if(!opts.keepCancelled && (has(i,"c24")||has(i,"cpost"))) continue;
    if(opts.grnScope){
      if(cfgMode==="yes" && !has(i,"cfg")) continue;
      if(cfgMode==="no"  &&  has(i,"cfg")) continue;
      if(pendMode==="exclude" && has(i,"pend")) continue;
    }
    let ok=true;
    for(const [k] of FILTERS){ if(sel[k].size && !sel[k].has(R.codes[k][i])){ ok=false; break; } }
    if(ok && matchQ(i)) out.push(i);
  }
  return out;
}
function activeCount(){
  let n = q?1:0;
  FILTERS.forEach(([k])=>{ if(sel[k].size) n++; });
  return n;
}

function renderFilters(){
  const menus = FILTERS.map(([k,label])=>{
    const used=new Set(); const c=R.codes[k];
    for(let i=0;i<R.n;i++) used.add(c[i]);
    const opts=[...used].sort((a,b)=>String(R.dims[k][a]).localeCompare(String(R.dims[k][b])));
    const body=opts.map(code=>
      `<div class="opt${sel[k].has(code)?" on":""}" data-k="${k}" data-c="${code}">
         <span class="box"></span><span class="nm">${R.dims[k][code]||"(blank)"}</span></div>`).join("");
    const n=sel[k].size;
    return `<div class="drop" data-menu="${k}">
      <button class="${n?"active":""}" data-toggle="${k}">${label}${n?` <span class="cnt">${n}</span>`:""} \\u25BE</button>
      <div class="menu" id="menu-${k}">
        ${opts.length>10?`<input type="search" placeholder="Search ${label}" data-search="${k}">`:""}
        <div class="opts" id="opts-${k}">${body}</div>
        <div class="foot"><a data-all="${k}">Select all</a><a data-none="${k}">Clear</a></div>
      </div></div>`;
  }).join("");
  document.getElementById("frow").innerHTML =
    `<div class="search"><span class="ic">\\u2315</span>
       <input type="search" id="gq" placeholder="Search PO number, invoice number, PO date or posting date" value="${q}"></div>`+
    menus +
    `<div class="seg"><em>GRN cfg</em>
       <span data-cfg="yes" class="${cfgMode==="yes"?"on":""}">Yes</span>
       <span data-cfg="no" class="${cfgMode==="no"?"on":""}">No</span>
       <span data-cfg="all" class="${cfgMode==="all"?"on":""}">All</span></div>
     <div class="seg"><em>Pending GRN</em>
       <span data-pend="exclude" class="${pendMode==="exclude"?"on":""}">Exclude</span>
       <span data-pend="include" class="${pendMode==="include"?"on":""}">Include</span></div>
     <button class="ghost" id="clearall">Clear all</button>`;

  const tags=[];
  if(q) tags.push(`<span class="tag">search <b>${q}</b></span>`);
  FILTERS.forEach(([k,label])=>{
    if(!sel[k].size) return;
    const names=[...sel[k]].map(c=>R.dims[k][c]);
    tags.push(`<span class="tag">${label} <b>${names.length>2?names.length+" selected":names.join(", ")}</b></span>`);
  });
  document.getElementById("activeline").innerHTML =
    (tags.length?tags.join(""):`<span class="tag">no filters \\u2014 showing everything</span>`)+
    `<span>${LIVE[tab]?"":"filters do not apply to this tab"}</span>`;
  bindFilters();
}

function bindFilters(){
  const gq=document.getElementById("gq");
  gq.oninput=()=>{ q=gq.value; refresh(true); };
  document.querySelectorAll("[data-toggle]").forEach(b=>b.onclick=e=>{
    e.stopPropagation();
    const k=b.dataset.toggle, el=document.getElementById("menu-"+k);
    const wasOpen=el.classList.contains("open");
    document.querySelectorAll(".menu").forEach(m=>m.classList.remove("open"));
    if(!wasOpen){ el.classList.add("open"); openMenu=k; } else openMenu=null;
  });
  document.querySelectorAll(".menu").forEach(m=>m.onclick=e=>e.stopPropagation());
  document.querySelectorAll("[data-k]").forEach(el=>el.onclick=()=>{
    const k=el.dataset.k, c=+el.dataset.c;
    sel[k].has(c) ? sel[k].delete(c) : sel[k].add(c);
    el.classList.toggle("on");
    refresh(true);
  });
  document.querySelectorAll("[data-search]").forEach(inp=>inp.oninput=()=>{
    const k=inp.dataset.search, t=inp.value.toLowerCase();
    document.getElementById("opts-"+k).querySelectorAll(".opt").forEach(o=>{
      o.style.display = o.textContent.toLowerCase().indexOf(t)>=0 ? "" : "none"; });
  });
  document.querySelectorAll("[data-all]").forEach(a=>a.onclick=()=>{
    const k=a.dataset.all; const c=R.codes[k];
    for(let i=0;i<R.n;i++) sel[k].add(c[i]);
    refresh();
  });
  document.querySelectorAll("[data-none]").forEach(a=>a.onclick=()=>{
    sel[a.dataset.none].clear(); refresh(); });
  document.querySelectorAll("[data-cfg]").forEach(el=>el.onclick=()=>{ cfgMode=el.dataset.cfg; refresh(); });
  document.querySelectorAll("[data-pend]").forEach(el=>el.onclick=()=>{ pendMode=el.dataset.pend; refresh(); });
  document.getElementById("clearall").onclick=()=>{
    FILTERS.forEach(([k])=>sel[k].clear()); q=""; refresh(); };
}
document.addEventListener("click",()=>{
  document.querySelectorAll(".menu").forEach(m=>m.classList.remove("open")); openMenu=null; });

/* ---------- aggregation ---------- */
function blank(){ return {po:0,inv:0,grn:0,lpi:0,lig:0}; }
function add(a,i){
  const po=R.po[i], inv=R.inv[i], grn=R.grn[i], pr=R.price[i];
  a.po+=po; a.inv+=inv; a.grn+=grn; a.lpi+=(po-inv)*pr; a.lig+=(inv-grn)*pr;
  return a;
}
function ratios(a){
  return Object.assign({}, a, {
    poInv: a.po? a.inv/a.po*100 : null,
    poGrn: a.po? a.grn/a.po*100 : null,
    invGrn: a.inv? a.grn/a.inv*100 : null,
    lpg: a.lpi + a.lig});
}
function groupBy(idx,key){
  const m=new Map(), c=R.codes[key];
  idx.forEach(i=>{ const k=c[i]; if(!m.has(k)) m.set(k,blank()); add(m.get(k),i); });
  return [...m.entries()].map(([k,v])=>Object.assign({label:R.dims[key][k]||"(blank)"},ratios(v)));
}

/* invoice-level roll-up, used by the POD summary */
function invoices(idx){
  const m=new Map();
  idx.forEach(i=>{
    const key=R.codes.invoice[i];
    let o=m.get(key);
    if(!o){ o={inv:R.dims.invoice[key], wh:R.codes.warehouse[i], tier:R.codes.tier[i],
               date:R.codes.date[i], amt:0, shortq:0, shortv:0,
               pod:false, grn:false, c24:false, cpost:false, podcanc:false};
            m.set(key,o); }
    o.amt+=R.amt[i]; o.shortq+=R.shortq[i]; o.shortv+=R.shortq[i]*R.price[i];
    if(has(i,"haspod")) o.pod=true;
    if(has(i,"hasgrn")) o.grn=true;
    if(has(i,"c24")) o.c24=true;
    if(has(i,"cpost")) o.cpost=true;
    if(has(i,"podcanc")) o.podcanc=true;
  });
  return [...m.values()];
}

/* ---------- charts ---------- */
function bars(rows,opts){
  opts=opts||{};
  const max = opts.max!=null ? opts.max : Math.max(1,...rows.map(r=>r.v||0));
  return rows.map(r=>{
    const w=Math.max(0,Math.min(100,((r.v||0)/max)*100));
    const mark = opts.target!=null ? `<div style="position:absolute;top:-3px;bottom:-3px;width:1px;background:var(--ink);opacity:.5;left:${(opts.target/max)*100}%"></div>`:"";
    return `<div class="brow"><div class="lbl" title="${r.label}">${r.label}</div>`+
      `<div class="track2"><i class="${r.cls||"f-ink"}" style="width:${w}%"></i>${mark}</div>`+
      `<div class="amt">${r.text}</div></div>`;
  }).join("");
}
function stacked(segs,total){
  if(!total) return `<div class="none">Nothing to show.</div>`;
  return `<div class="stack">`+segs.map(s=>
    `<span class="${s.cls}" style="width:${(s.v/total)*100}%" title="${s.label}"></span>`).join("")+`</div>`+
    `<div class="legend">`+segs.filter(s=>s.v>0).map(s=>
      `<span><b class="${s.cls}"></b>${s.label} ${num(s.v)} (${((s.v/total)*100).toFixed(1)}%)</span>`).join("")+`</div>`;
}
const panelChart = (t,s,b) =>
  `<div class="chart"><h4>${t}</h4><div class="sub2">${s}</div>${b||`<div class="none">Nothing to show.</div>`}</div>`;

/* ---------- fill rate tab ---------- */
const METRICS=[["poGrn","PO to GRN"],["invGrn","INV to GRN"],["poInv","PO to INV"]];
const BREAKDOWNS=[["warehouse","Warehouse"],["platform","Customer"],["category","Category"],
                  ["city","City"],["location","Delivery location"],["sku","SKU"],
                  ["invoice","Invoice"],["po","PO number"]];
let fdim="warehouse", fsort={col:null,dir:-1}, fq="", fbreak={cols:[],rows:[]};

function fillShell(){
  return `<div class="fgrid" id="fkpi"></div>
    <div class="card"><div class="card-h"><h3>Fill rate by day</h3>
      <div class="tools"><button data-csv-hm="1">CSV</button></div>
      <div class="note">Every customer against each posting date in the current selection. Colour runs red below 70% to green above 96%.</div>
      </div><div class="hm" id="fheat"></div></div>
    <div class="card"><div class="card-h"><h3>Breakdown</h3>
      <div class="tools">
        <select id="fdim">${BREAKDOWNS.map(([k,l])=>`<option value="${k}">${l}</option>`).join("")}</select>
        <input type="search" id="fq2" placeholder="Filter rows">
        <button data-csv-bd="1">CSV</button>
      </div></div><div class="scroll" id="fbreak"></div></div>
    <div class="charts" id="fcharts"></div>`;
}
function renderFillKpis(t){
  const items=[["PO to INV",t.poInv,t.lpi,"invoiced against ordered"],
               ["PO to GRN",t.poGrn,t.lpg,"received against ordered"],
               ["INV to GRN",t.invGrn,t.lig,"received against invoiced"]];
  const qty=`<div class="fk"><div class="k">Quantities</div>
    <div class="d" style="margin-top:9px;line-height:2">
      PO <b style="font-family:var(--mono)">${num(t.po)}</b><br>
      Invoiced <b style="font-family:var(--mono)">${num(t.inv)}</b><br>
      Received <b style="font-family:var(--mono)">${num(t.grn)}</b></div></div>`;
  document.getElementById("fkpi").innerHTML=items.map(([k,v,loss,d])=>
    `<div class="fk"><div class="k">${k}</div><div class="v">${pct(v)}</div>
     <div class="m"><i style="width:${Math.max(0,Math.min(100,v||0))}%;background:${heat(v)}"></i></div>
     <div class="d">${d}</div><div class="loss">${inr(loss)} lost</div></div>`).join("")+qty;
}
function heatData(idx){
  const dates=[...new Set(idx.map(i=>R.codes.date[i]))].filter(c=>R.dims.date[c])
    .sort((a,b)=>String(R.dims.date[a]).localeCompare(String(R.dims.date[b])));
  const plats=[...new Set(idx.map(i=>R.codes.platform[i]))]
    .sort((a,b)=>String(R.dims.platform[a]).localeCompare(String(R.dims.platform[b])));
  const cell={},rowTot={};
  plats.forEach(p=>rowTot[p]=blank());
  idx.forEach(i=>{
    const k=R.codes.platform[i]+"|"+R.codes.date[i];
    if(!cell[k]) cell[k]=blank();
    add(cell[k],i); add(rowTot[R.codes.platform[i]],i);
  });
  return {dates,plats,cell,rowTot};
}
function renderHeat(idx){
  const h=heatData(idx), el=document.getElementById("fheat");
  if(!h.plats.length||!h.dates.length){ el.innerHTML=`<div class="empty">No rows match the current filters.</div>`; return; }
  const head=`<tr><th style="text-align:left;position:sticky;left:0;z-index:4;background:var(--surface);min-width:118px">Customer</th>`+
    `<th style="text-align:left;position:sticky;left:118px;z-index:4;background:var(--surface)">Metric</th>`+
    h.dates.map(d=>`<th>${fmtDay(R.dims.date[d])}</th>`).join("")+`<th>Total</th></tr>`;
  const body=h.plats.map(p=>METRICS.map(([m,label],j)=>{
    const cells=h.dates.map(d=>{
      const a=h.cell[p+"|"+d], v=a?ratios(a)[m]:null;
      return `<td class="cell" style="background:${heat(v)};color:${heatInk(v)}">${v==null?"":v.toFixed(1)+"%"}</td>`;
    }).join("");
    const tv=ratios(h.rowTot[p])[m];
    return `<tr>${j===0?`<td class="grp">${R.dims.platform[p]}</td>`:`<td class="rowh"></td>`}`+
      `<td class="met">${label}</td>${cells}`+
      `<td class="cell tot" style="background:${heat(tv)};color:${heatInk(tv)}">${tv==null?"":tv.toFixed(1)+"%"}</td></tr>`;
  }).join("")).join("");
  el.innerHTML=`<table><thead>${head}</thead><tbody>${body}</tbody></table>`;
}
function renderBreak(idx){
  let rows=groupBy(idx,fdim);
  if(fq){ const t=fq.toLowerCase(); rows=rows.filter(r=>r.label.toLowerCase().indexOf(t)>=0); }
  const cols=[["label",BREAKDOWNS.find(b=>b[0]===fdim)[1],"text"],
    ["po","PO qty","num"],["inv","INV qty","num"],["grn","GRN qty","num"],
    ["poInv","PO to INV","pct"],["poGrn","PO to GRN","pct"],["invGrn","INV to GRN","pct"],
    ["lpi","PO to INV loss","money"],["lpg","PO to GRN loss","money"],["lig","INV to GRN loss","money"]];
  if(fsort.col) rows.sort((a,b)=>{
    const x=a[fsort.col],y=b[fsort.col];
    if(typeof x==="string") return x.localeCompare(y)*fsort.dir;
    return (((x==null)?-1e18:x)-((y==null)?-1e18:y))*fsort.dir; });
  else rows.sort((a,b)=>b.lpg-a.lpg);
  fbreak={cols,rows};
  const head=cols.map(([k,l])=>`<th data-fs="${k}">${l}${fsort.col===k?(fsort.dir>0?" \\u2191":" \\u2193"):""}</th>`).join("");
  const body=rows.map(r=>"<tr>"+cols.map(([k,,kind])=>{
    if(kind==="text") return `<td class="doc">${r[k]}</td>`;
    if(kind==="pct") return `<td class="n" style="background:${r[k]==null?"transparent":heat(r[k])};color:${r[k]==null?"inherit":heatInk(r[k])}">${pct(r[k])}</td>`;
    if(kind==="money") return `<td class="n${r[k]>0?" neg":""}">${inr(r[k])}</td>`;
    return `<td class="n">${num(r[k])}</td>`;
  }).join("")+"</tr>").join("");
  const el=document.getElementById("fbreak");
  el.innerHTML = rows.length ? `<table><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table>`
                             : `<div class="empty">No rows match the current filters.</div>`;
  el.querySelectorAll("th").forEach(th=>th.onclick=()=>{
    const k=th.dataset.fs;
    if(fsort.col===k) fsort.dir*=-1; else fsort={col:k,dir:k==="label"?1:-1};
    renderBreak(rowIdx({grnScope:true}));
  });
}
function renderFillCharts(idx,m){
  const out=[];
  const byW=groupBy(idx,"warehouse").sort((a,b)=>(b.poGrn||0)-(a.poGrn||0));
  if(byW.length) out.push(panelChart("PO to GRN by warehouse","Against a 95% reference line",
    bars(byW.map(r=>({label:r.label,v:r.poGrn,text:pct(r.poGrn),
      cls:r.poGrn>=98?"f-clear":r.poGrn>=95?"f-wait":"f-risk"})),{max:100,target:95})));
  const byP=groupBy(idx,"platform").sort((a,b)=>(b.poGrn||0)-(a.poGrn||0));
  if(byP.length) out.push(panelChart("PO to GRN by customer","Against a 95% reference line",
    bars(byP.map(r=>({label:r.label,v:r.poGrn,text:pct(r.poGrn),
      cls:r.poGrn>=98?"f-clear":r.poGrn>=95?"f-wait":"f-risk"})),{max:100,target:95})));
  const lost=byP.filter(r=>r.lpg>0).sort((a,b)=>b.lpg-a.lpg);
  if(lost.length) out.push(panelChart("Value lost, PO to GRN","Ordered but never received",
    bars(lost.map(r=>({label:r.label,v:r.lpg,text:inr(r.lpg),cls:"f-risk"})))));
  const bySku=groupBy(idx,"sku").filter(r=>r.lpg>0).sort((a,b)=>b.lpg-a.lpg).slice(0,10);
  if(bySku.length) out.push(panelChart("Worst SKUs by value lost","Top ten in the current selection",
    bars(bySku.map(r=>({label:r.label,v:r.lpg,text:inr(r.lpg),cls:"f-risk"})))));
  document.getElementById("fcharts").innerHTML=out.join("");
}
function renderFill(){
  const idx=rowIdx({grnScope:true});
  renderFillKpis(ratios(idx.reduce((a,i)=>add(a,i),blank())));
  renderHeat(idx);
  renderBreak(idx);
  renderFillCharts(idx,DATA.months[month]);
  const d=document.getElementById("fdim");
  d.value=fdim; d.onchange=()=>{ fdim=d.value; fsort={col:null,dir:-1}; renderBreak(rowIdx({grnScope:true})); };
  const s=document.getElementById("fq2");
  s.value=fq; s.oninput=()=>{ fq=s.value; renderBreak(rowIdx({grnScope:true})); };
  document.querySelector("[data-csv-bd]").onclick=()=>{
    saveBlob(new Blob([toCsv(fbreak.cols.map(c=>c[1]),
      fbreak.rows.map(r=>fbreak.cols.map(c=>{
        const v=r[c[0]]; return typeof v==="string"?v:(v==null?"":Math.round(v*100)/100); })))],
      {type:"text/csv;charset=utf-8"}), "fill_rate_by_"+fdim+".csv"); };
  document.querySelector("[data-csv-hm]").onclick=()=>{
    const h=heatData(rowIdx({grnScope:true}));
    const head=["Customer","Metric"].concat(h.dates.map(d=>R.dims.date[d])).concat(["Total"]);
    const rows=[];
    h.plats.forEach(p=>METRICS.forEach(([m,label])=>{
      rows.push([R.dims.platform[p],label].concat(h.dates.map(d=>{
        const a=h.cell[p+"|"+d], v=a?ratios(a)[m]:null; return v==null?"":v.toFixed(2); }))
        .concat([ratios(h.rowTot[p])[m].toFixed(2)])); }));
    saveBlob(new Blob([toCsv(head,rows)],{type:"text/csv;charset=utf-8"}),"fill_rate_by_day.csv"); };
}

/* ---------- POD summary tab ---------- */
const POD_GROUPS=[
  ["Total Invoice",["count","value"]],
  ["POD Received",["count","value"]],
  ["POD Not Received",["count","value"]],
  ["POD/GRN Not Received",["count","value"]],
  ["POD Received T1",["count","value"]],
  ["POD Received T2",["count","value"]],
  ["POD Not Received But GRN Done",["count","value"]],
  ["Short GRN",["count","qty","value"]],
  ["Cancelled Within 24 hours",["count","value"]],
  ["Cancelled After 24 hours",["count","value"]],
  ["POD cancelled, no INC",["count","value"]],
];
function podShell(){
  return `<div class="kpis" id="pkpi"></div>
   <div class="card"><div class="card-h"><h3>Invoice to payment tracker</h3>
     <div class="tools"><button data-csv-pod="1">CSV</button></div>
     <div class="note">Unique invoices by warehouse. POD Received, POD Not Received, Cancelled Within 24 hours and Cancelled After 24 hours are mutually exclusive and add up to Total Invoice. Value is the sum of Invoice Amt.</div>
     </div><div class="pods" id="podtable"></div></div>
   <div class="card"><div class="card-h"><h3>Pending PODs by posting date</h3>
     <div class="tools"><button data-csv-pend="1">CSV</button></div>
     <div class="note">Unique invoice count where no POD has been received, grouped by T1/T2 then warehouse.</div>
     </div><div class="pods" id="podpivot"></div></div>
   <div class="charts" id="pcharts"></div>`;
}
function podBuckets(list){
  const b={total:[0,0],pod:[0,0],nopod:[0,0],nopodnogrn:[0,0],t1:[0,0],t2:[0,0],
           nopodgrn:[0,0],short:[0,0,0],c24:[0,0],cpost:[0,0],flag:[0,0]};
  list.forEach(o=>{
    b.total[0]++; b.total[1]+=o.amt;
    if(o.c24){ b.c24[0]++; b.c24[1]+=o.amt; }
    else if(o.cpost){ b.cpost[0]++; b.cpost[1]+=o.amt; }
    else if(o.pod){
      b.pod[0]++; b.pod[1]+=o.amt;
      const tier=(R.dims.tier[o.tier]||"").toUpperCase();
      if(tier==="T1"){ b.t1[0]++; b.t1[1]+=o.amt; }
      else if(tier==="T2"){ b.t2[0]++; b.t2[1]+=o.amt; }
    } else {
      b.nopod[0]++; b.nopod[1]+=o.amt;
      if(o.grn){ b.nopodgrn[0]++; b.nopodgrn[1]+=o.amt; }
      else { b.nopodnogrn[0]++; b.nopodnogrn[1]+=o.amt; }
    }
    if(o.shortq>0){ b.short[0]++; b.short[1]+=o.shortq; b.short[2]+=o.shortv; }
    if(o.podcanc){ b.flag[0]++; b.flag[1]+=o.amt; }
  });
  return b;
}
const PODKEYS=["total","pod","nopod","nopodnogrn","t1","t2","nopodgrn","short","c24","cpost","flag"];
function podCells(b){
  const out=[];
  PODKEYS.forEach(k=>{
    const v=b[k];
    if(k==="short"){ out.push([v[0],"num"],[v[1],"num"],[v[2],"money"]); }
    else out.push([v[0],"num"],[v[1],"money"]);
  });
  return out;
}
function renderPod(){
  const list=invoices(rowIdx({keepCancelled:true}));
  const whs=[...new Set(list.map(o=>o.wh))].sort((a,b)=>
    String(R.dims.warehouse[a]).localeCompare(String(R.dims.warehouse[b])));
  const head1=`<tr><th style="position:sticky;left:0;z-index:4;background:var(--surface)"></th>`+
    POD_GROUPS.map(([g,parts])=>`<th class="grp1" colspan="${parts.length}">${g}</th>`).join("")+`</tr>`;
  const head2=`<tr><th style="text-align:left;position:sticky;left:0;z-index:4;background:var(--surface);min-width:110px">Warehouse</th>`+
    POD_GROUPS.map(([,parts])=>parts.map((p,i)=>
      `<th class="sub${i===0?" first":""}">${p==="count"?"Count":p==="qty"?"Short qty":"Value"}</th>`).join("")).join("")+`</tr>`;
  const rowFor=(label,b,cls)=>{
    const cells=podCells(b).map(([v,kind],i)=>
      `<td class="v${i%2===0?" first":""}${label==="Total"?"":""}">${kind==="money"?inr(v):num(v)}</td>`).join("");
    return `<tr class="${cls||""}"><td class="wh">${label}</td>${cells}</tr>`;
  };
  const body=whs.map(w=>rowFor(R.dims.warehouse[w]||"(blank)",
    podBuckets(list.filter(o=>o.wh===w)))).join("")+
    rowFor("Total",podBuckets(list),"total");
  document.getElementById("podtable").innerHTML =
    list.length ? `<table><thead>${head1}${head2}</thead><tbody>${body}</tbody></table>`
                : `<div class="empty">No invoices match the current filters.</div>`;

  const tot=podBuckets(list);
  const kp=[["Total invoices",num(tot.total[0]),inr(tot.total[1]),""],
    ["POD received",num(tot.pod[0]),pctOf(tot.pod[0],tot.total[0]),"clear"],
    ["POD not received",num(tot.nopod[0]),inr(tot.nopod[1]),tot.nopod[0]?"risk":""],
    ["POD/GRN not received",num(tot.nopodnogrn[0]),inr(tot.nopodnogrn[1]),tot.nopodnogrn[0]?"risk":""],
    ["No POD, GRN done",num(tot.nopodgrn[0]),inr(tot.nopodgrn[1]),tot.nopodgrn[0]?"wait":""],
    ["Short GRN",num(tot.short[0]),inr(tot.short[2]),tot.short[0]?"risk":""],
    ["Cancelled within 24h",num(tot.c24[0]),inr(tot.c24[1]),tot.c24[0]?"wait":""],
    ["Cancelled after 24h",num(tot.cpost[0]),inr(tot.cpost[1]),tot.cpost[0]?"wait":""],
    ["POD cancelled, no INC",num(tot.flag[0]),"needs checking",tot.flag[0]?"risk":""]];
  document.getElementById("pkpi").innerHTML=kp.map(([k,v,d,c])=>
    `<div class="kpi ${c}"><div class="k">${k}</div><div class="v">${v}</div><div class="d">${d}</div></div>`).join("");

  renderPendingPivot(list);

  const out=[];
  const rows=whs.map(w=>{ const b=podBuckets(list.filter(o=>o.wh===w));
    return {label:R.dims.warehouse[w]||"(blank)", v:b.total[0]?b.pod[0]/b.total[0]*100:0,
            nopod:b.nopod[0], flagged:b.flag[0]}; });
  out.push(panelChart("POD coverage by warehouse","Share of invoices with a POD received",
    bars(rows.slice().sort((a,b)=>b.v-a.v).map(r=>({label:r.label,v:r.v,text:pct(r.v),
      cls:r.v>=98?"f-clear":r.v>=95?"f-wait":"f-risk"})),{max:100,target:98})));
  const miss=rows.filter(r=>r.nopod>0).sort((a,b)=>b.nopod-a.nopod);
  out.push(panelChart("Invoices with no POD","Count by warehouse",
    miss.length?bars(miss.map(r=>({label:r.label,v:r.nopod,text:num(r.nopod),cls:"f-risk"}))):null));
  out.push(panelChart("Invoice split","Every invoice in the current selection",
    stacked([{label:"POD received",v:tot.pod[0],cls:"f-clear"},
             {label:"POD not received",v:tot.nopod[0],cls:"f-risk"},
             {label:"Cancelled within 24h",v:tot.c24[0],cls:"f-wait"},
             {label:"Cancelled after 24h",v:tot.cpost[0],cls:"f-mute"}], tot.total[0])));
  document.getElementById("pcharts").innerHTML=out.join("");

  document.querySelector("[data-csv-pod]").onclick=()=>{
    const head=["Warehouse"];
    POD_GROUPS.forEach(([g,parts])=>parts.forEach(p=>head.push(g+" "+p)));
    const rows=whs.map(w=>[R.dims.warehouse[w]].concat(
      podCells(podBuckets(list.filter(o=>o.wh===w))).map(c=>Math.round(c[0]*100)/100)));
    rows.push(["Total"].concat(podCells(podBuckets(list)).map(c=>Math.round(c[0]*100)/100)));
    saveBlob(new Blob([toCsv(head,rows)],{type:"text/csv;charset=utf-8"}),"pod_summary.csv"); };
}
function pctOf(a,b){ return b? (a/b*100).toFixed(1)+"% of invoices" : ""; }

function renderPendingPivot(list){
  const pend=list.filter(o=>!o.pod && !o.c24 && !o.cpost);
  const el=document.getElementById("podpivot");
  if(!pend.length){ el.innerHTML=`<div class="empty">No pending PODs in the current selection.</div>`; return; }
  const dates=[...new Set(pend.map(o=>o.date))].filter(d=>R.dims.date[d])
    .sort((a,b)=>String(R.dims.date[a]).localeCompare(String(R.dims.date[b])));
  const tiers=[...new Set(pend.map(o=>o.tier))]
    .sort((a,b)=>String(R.dims.tier[a]).localeCompare(String(R.dims.tier[b])));
  const key=(t,w,d)=>t+"|"+w+"|"+d;
  const cnt={};
  pend.forEach(o=>{ [key(o.tier,o.wh,o.date),key(o.tier,"ALL",o.date),key("ALL","ALL",o.date)]
    .forEach(k=>cnt[k]=(cnt[k]||0)+1); });
  const head=`<tr><th style="text-align:left;position:sticky;left:0;z-index:4;background:var(--surface);min-width:80px">T1/T2</th>`+
    `<th style="text-align:left">Warehouse</th>`+
    dates.map(d=>`<th class="sub">${fmtDay(R.dims.date[d])}</th>`).join("")+
    `<th class="sub">Grand total</th></tr>`;
  let body="";
  tiers.forEach(t=>{
    const whs=[...new Set(pend.filter(o=>o.tier===t).map(o=>o.wh))]
      .sort((a,b)=>String(R.dims.warehouse[a]).localeCompare(String(R.dims.warehouse[b])));
    whs.forEach((w,j)=>{
      const cells=dates.map(d=>`<td class="v">${cnt[key(t,w,d)]||""}</td>`).join("");
      const tot=dates.reduce((s,d)=>s+(cnt[key(t,w,d)]||0),0);
      body+=`<tr><td class="wh">${j===0?(R.dims.tier[t]||"(blank)"):""}</td>`+
        `<td style="font-family:var(--mono);font-size:11.5px">${R.dims.warehouse[w]||"(blank)"}</td>`+
        cells+`<td class="v" style="font-weight:600">${tot}</td></tr>`;
    });
    const subs=dates.map(d=>`<td class="v">${cnt[key(t,"ALL",d)]||""}</td>`).join("");
    const st=dates.reduce((s,d)=>s+(cnt[key(t,"ALL",d)]||0),0);
    body+=`<tr class="total"><td class="wh">${R.dims.tier[t]||"(blank)"} total</td><td></td>${subs}<td class="v">${st}</td></tr>`;
  });
  const gs=dates.map(d=>`<td class="v">${cnt[key("ALL","ALL",d)]||""}</td>`).join("");
  body+=`<tr class="total"><td class="wh">Grand total</td><td></td>${gs}<td class="v">${pend.length}</td></tr>`;
  el.innerHTML=`<table><thead>${head}</thead><tbody>${body}</tbody></table>`;
  document.querySelector("[data-csv-pend]").onclick=()=>{
    const head=["T1/T2","Warehouse"].concat(dates.map(d=>R.dims.date[d])).concat(["Grand total"]);
    const rows=[];
    tiers.forEach(t=>{
      [...new Set(pend.filter(o=>o.tier===t).map(o=>o.wh))].forEach(w=>{
        rows.push([R.dims.tier[t],R.dims.warehouse[w]]
          .concat(dates.map(d=>cnt[key(t,w,d)]||0))
          .concat([dates.reduce((s,d)=>s+(cnt[key(t,w,d)]||0),0)])); }); });
    saveBlob(new Blob([toCsv(head,rows)],{type:"text/csv;charset=utf-8"}),"pending_pods.csv"); };
}

/* ---------- pre-aggregated tables ---------- */
const state={};
function tableCard(t,key,idx){
  const id=key+"-"+idx;
  state[id]={t,sort:null,dir:1,q:""};
  return `<div class="card"><div class="card-h"><h3>${t.label}</h3>
    <div class="tools"><input type="search" placeholder="Filter rows" data-f="${id}">
    <button data-csv="${id}">CSV</button></div>
    ${t.note?`<div class="note">${t.note}</div>`:""}
    </div><div class="scroll" id="body-${id}"></div></div>`;
}
function renderBody(id){
  const s=state[id],t=s.t,el=document.getElementById("body-"+id);
  if(!el) return;
  if(!t.rows.length){ el.innerHTML=`<div class="empty">Nothing in this bucket.</div>`; return; }
  let rows=t.rows;
  if(s.q){ const x=s.q.toLowerCase();
    rows=rows.filter(r=>r.some(v=>v!=null&&String(v).toLowerCase().indexOf(x)>=0)); }
  if(s.sort!==null){ const k=s.sort,kind=t.kinds[k];
    rows=rows.slice().sort((a,b)=>{
      let x=a[k],y=b[k];
      if(x==null) return 1; if(y==null) return -1;
      return kind!=="text" ? (Number(x)-Number(y))*s.dir : String(x).localeCompare(String(y))*s.dir; }); }
  const head=t.columns.map((c,i)=>`<th data-s="${id}" data-i="${i}">${c.split("_").join(" ")}${s.sort===i?(s.dir>0?" \\u2191":" \\u2193"):""}</th>`).join("");
  const body=rows.map(r=>"<tr>"+r.map((v,i)=>{
    const kind=t.kinds[i];
    if(kind==="text"){
      const sv=String(v==null?"":v);
      if(sv.toLowerCase().indexOf("pending")>=0) return `<td><span class="pill p">${sv}</span></td>`;
      if(sv.toLowerCase().indexOf("raised")>=0) return `<td><span class="pill r">${sv}</span></td>`;
      return `<td>${fmt(v,kind)}</td>`;
    }
    return `<td class="n${(typeof v==="number"&&v<0)?" neg":""}">${fmt(v,kind)}</td>`;
  }).join("")+"</tr>").join("");
  el.innerHTML=`<table><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table>`;
  el.querySelectorAll("th").forEach(th=>th.onclick=()=>{
    const i=+th.dataset.i;
    if(s.sort===i) s.dir*=-1; else { s.sort=i; s.dir=(t.kinds[i]==="text")?1:-1; }
    renderBody(id); });
}

/* ---------- shell ---------- */
function refresh(keepMenus){
  if(!keepMenus) renderFilters();
  else {
    const tags=[]; if(q) tags.push(`<span class="tag">search <b>${q}</b></span>`);
    FILTERS.forEach(([k,label])=>{ if(sel[k].size){
      const names=[...sel[k]].map(c=>R.dims[k][c]);
      tags.push(`<span class="tag">${label} <b>${names.length>2?names.length+" selected":names.join(", ")}</b></span>`); }});
    document.getElementById("activeline").innerHTML=
      (tags.length?tags.join(""):`<span class="tag">no filters \\u2014 showing everything</span>`);
  }
  if(tab==="fill") renderFill();
  if(tab==="pod") renderPod();
}
function draw(){
  const m=DATA.months[month];
  document.getElementById("tabs").innerHTML=SECTIONS.map(([k,l])=>
    `<button class="tab" role="tab" data-t="${k}" aria-selected="${k===tab}">${l}</button>`).join("");
  document.getElementById("panels").innerHTML=SECTIONS.map(([k])=>{
    let inner;
    if(k==="fill") inner = R?fillShell():`<div class="empty">Row-level data was not embedded.</div>`;
    else if(k==="pod") inner = R?podShell():`<div class="empty">Row-level data was not embedded.</div>`;
    else inner = `<div class="notice">The filters above do not apply to this tab.</div>`+
                 (m.tables[k]||[]).map((t,j)=>tableCard(t,k,j)).join("");
    return `<div class="panel${k===tab?" on":""}" id="p-${k}">${inner}</div>`;
  }).join("");
  SECTIONS.forEach(([k])=>{ if(LIVE[k]) return;
    (m.tables[k]||[]).forEach((t,j)=>renderBody(k+"-"+j)); });
  document.querySelectorAll(".tab").forEach(b=>b.onclick=()=>{
    tab=b.dataset.t;
    document.querySelectorAll(".tab").forEach(x=>x.setAttribute("aria-selected",x===b));
    document.querySelectorAll(".panel").forEach(p=>p.classList.toggle("on",p.id==="p-"+tab));
    refresh();
  });
  document.querySelectorAll("[data-f]").forEach(i=>i.oninput=()=>{
    state[i.dataset.f].q=i.value; renderBody(i.dataset.f); });
  document.querySelectorAll("[data-csv]").forEach(b=>b.onclick=()=>{
    const t=state[b.dataset.csv].t;
    saveBlob(new Blob([toCsv(t.columns,t.rows)],{type:"text/csv;charset=utf-8"}),
      month.split(" ").join("_")+"_"+t.label.split(" ").join("_")+".csv"); });
  refresh();
}

document.getElementById("dl-xlsx").onclick=()=>{
  const m=DATA.months[month];
  if(typeof XLSX==="undefined"){
    alert("The spreadsheet library did not load, so the workbook is unavailable offline. Use the CSV button on each table instead.");
    return; }
  const wb=XLSX.utils.book_new(), used={};
  Object.keys(m.tables).forEach(k=>(m.tables[k]||[]).forEach(t=>{
    let name=(k+" "+t.label).slice(0,31); while(used[name]) name=name.slice(0,29)+"_"+Object.keys(used).length;
    used[name]=1;
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([t.columns].concat(t.rows)),name); }));
  XLSX.writeFile(wb,"SKU_POD_Tracker_"+month.split(" ").join("_")+".xlsx"); };

document.getElementById("stamp").textContent =
  "Generated "+DATA.generated+" \\u00B7 GRN over POD except open-PO parties \\u00B7 fill rate on GRN Config = Yes only \\u00B7 counts on unique invoice numbers";
const msel=document.getElementById("month");
msel.innerHTML=Object.keys(DATA.months).map(m=>`<option>${m}</option>`).join("");
msel.onchange=()=>{ month=msel.value; draw(); };
document.getElementById("notes").innerHTML=`
  <strong>How the numbers are built.</strong>
  Received quantity takes GRN over POD for every party except the open-PO accounts
  (NB, Metro C&amp;C, Reliance Signature, Reliance Smart), where POD wins.
  Fill rate is measured only on GRN Config = Yes accounts, since those are the only ones
  that send a GRN; an invoice with no GRN yet is Pending GRN, never Short GRN.
  On the POD summary, cancelled within 24 hours is <code>Is Cancel = Yes</code> and cancelled
  after 24 hours is <code>Return Reason = INC</code>; where both are set the invoice counts as
  within 24 hours, so the four buckets stay mutually exclusive and add to Total Invoice.
  Invoices whose POD Status is CANCELLED without an INC reason are counted separately for checking.
  A short on any single line marks the whole invoice short, so every count is on unique invoice numbers.`;
if(R) renderFilters();
draw();
</script></body></html>"""
