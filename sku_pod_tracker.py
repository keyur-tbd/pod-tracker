"""
SKU POD MASTER TRACKER -- metrics engine
========================================
Reads one Google Sheet per month via a service account and computes:
  1. PO -> GRN fill rate (denominator PO Qty) and INV -> GRN fill rate (denominator Invoice Qty)
  2. Pending SRO / Credit Memo on Short GRN, counted on unique invoice numbers
  3. POD tracker -- invoice universe split by GRN / POD receipt
  4. Cancelled invoices (Is Cancel = Yes, Return Reason = INC)
  5. Exception buckets, including the POD-vs-GRN quantity gap

Run:
    python sku_pod_tracker.py                      # all months in MONTH_SOURCES
    python sku_pod_tracker.py --csv sample.csv --month "August 2026"

Dependencies: pip install gspread pandas openpyxl
"""

import argparse
import faulthandler
import re
import warnings

faulthandler.enable()   # a hard crash prints a stack instead of exiting silently
import numpy as np
import pandas as pd
from pandas.errors import OutOfBoundsDatetime

# ============================================================================
# CONFIG
# ============================================================================
# Settings live in config.py so that replacing this file never wipes them.
# If config.py is missing, the defaults below are used.
SERVICE_ACCOUNT_FILE = "service_account.json"
MONTH_SOURCES = {}
WORKSHEET_NAME = "SKU POD ENTRY"
HEADER_ROW = 2
OUTPUT_XLSX = "sku_pod_tracker.xlsx"
OUTPUT_HTML = "sku_pod_dashboard.html"
OPEN_PO_PARTIES = ["NB", "METRO C&C", "RELIANCE SIGNATURE", "RELIANCE SMART"]
GRN_PRIORITY = ["NET GRN QTY", "GRN Qty", "Auto GRN Qty"]
GRN_CONFIG_NO_FALLS_BACK_TO_POD = True
EXCLUDE_PENDING_GRN_FROM_FILL_RATE = False

try:
    import config as _cfg
    for _k in ["SERVICE_ACCOUNT_FILE", "MONTH_SOURCES", "WORKSHEET_NAME", "HEADER_ROW",
               "OUTPUT_XLSX", "OUTPUT_HTML", "OPEN_PO_PARTIES", "GRN_PRIORITY",
               "GRN_CONFIG_NO_FALLS_BACK_TO_POD", "EXCLUDE_PENDING_GRN_FROM_FILL_RATE"]:
        if hasattr(_cfg, _k):
            globals()[_k] = getattr(_cfg, _k)
except ImportError:
    pass

# ---------------------------------------------------------------------------
NUMERIC_COLS = [
    "PO Qty", "Invoice Qty.", "Auto GRN Qty", "Auto GDN QTY", "GRN Qty",
    "GDN QTY", "NET GRN QTY", "Invoice Amt", "PO Amt", "GRN Amt", "GDN amt",
    "POD Qty", "Inv - POD Diff", "POD & GRN Qty Diff.", "Return QTY",
    "CN AMOUNT", "Unit Price", "count",
]
DATE_COLS = ["PO Date", "Posting Date", "Delivery/GRN Date"]

# Tried in order. The sheet mixes "30-Jul-26" and "04-08-2026"; anything that
# matches none of these falls back to the slow per-element parser.
DATE_FORMATS = ["%d-%b-%y", "%d-%m-%Y", "%d-%b-%Y", "%d-%m-%y",
                "%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d", "%m/%d/%Y"]
INVOICE_KEYS = ["month", "Invoice No."]

TEXT_COLS = ["Display Name", "Is Cancel", "Return Reason", "GRN Config",
             "SRO No", "RTV DOC", "GRN Status", "POD Status", "Concat",
             "Customer location name", "Warehouse", "Ship-to City",
             "CATAGORY", "T1/T2", "Description", "Item/Account", "PO. No."]


def squash(s):
    """Uppercase, letters and digits only -- tolerant party-name matching."""
    return re.sub(r"[^A-Z0-9]", "", str(s).upper())


OPEN_PO_KEYS = {squash(p) for p in OPEN_PO_PARTIES}


# ============================================================================
# LOAD
# ============================================================================
def load_from_sheets(month_sources=None) -> pd.DataFrame:
    import gspread
    from gspread.exceptions import SpreadsheetNotFound, WorksheetNotFound, APIError

    sources = month_sources or MONTH_SOURCES
    if not sources:
        raise SystemExit(
            "No sheets configured. Add your month -> sheet ID pairs to MONTH_SOURCES "
            "in config.py.")
    bad = [m for m, sid in sources.items() if not sid or "PUT_" in str(sid)]
    if bad:
        raise SystemExit(
            "These months still hold the placeholder sheet ID: " + ", ".join(bad) +
            "\nPut the real ID (the /d/<ID>/edit part of the sheet URL) in config.py.")

    import json as _json, os
    if not os.path.exists(SERVICE_ACCOUNT_FILE):
        raise SystemExit(
            f"Service account key not found at {os.path.abspath(SERVICE_ACCOUNT_FILE)}.\n"
            "Put the downloaded JSON key there, or point SERVICE_ACCOUNT_FILE in "
            "config.py at wherever it lives.")
    with open(SERVICE_ACCOUNT_FILE, encoding="utf-8") as fh:
        sa_email = _json.load(fh).get("client_email", "the service account")
    gc = gspread.service_account(filename=SERVICE_ACCOUNT_FILE)

    frames = []
    for month, sheet_id in sources.items():
        try:
            book = gc.open_by_key(sheet_id)
        except SpreadsheetNotFound:
            raise SystemExit(
                f"{month}: no sheet found with ID {sheet_id!r}.\n"
                f"Either the ID is wrong, or the sheet has not been shared with {sa_email}.\n"
                "The ID is the part of the URL between /d/ and /edit -- not the whole URL "
                "and not the #gid number.")
        except APIError as e:
            raise SystemExit(f"{month}: Google refused the request -- {e}")
        try:
            ws = book.worksheet(WORKSHEET_NAME)
        except WorksheetNotFound:
            tabs = ", ".join(w.title for w in book.worksheets())
            raise SystemExit(
                f"{month}: no tab named {WORKSHEET_NAME!r} in {book.title!r}.\n"
                f"Tabs available: {tabs}")
        values = ws.get_all_values()          # formulas arrive already evaluated
        if len(values) <= HEADER_ROW:
            print(f"  {month}: no data rows, skipped")
            continue
        frame = pd.DataFrame(values[HEADER_ROW:], columns=values[HEADER_ROW - 1])
        frame["month"] = month
        frames.append(frame)
        print(f"  {month}: {len(frame)} rows")
    if not frames:
        raise RuntimeError("No data loaded from any month.")
    return pd.concat(frames, ignore_index=True)


def load_from_csv(path: str, month: str = "Sample") -> pd.DataFrame:
    df = pd.read_csv(path, skiprows=HEADER_ROW - 1, dtype=str)
    df["month"] = month
    return df


# ============================================================================
# NORMALISE
# ============================================================================
def to_num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(
        s.astype(str)
         .str.replace(",", "", regex=False)
         .str.replace("\u20b9", "", regex=False)
         .str.replace("%", "", regex=False)
         .str.strip()
         .replace({"": None, "nan": None, "None": None, "-": None}),
        errors="coerce",
    )


# Dates outside this window are treated as data entry errors, not dates.
# A typo like "1-08-09" parses to the year 1, which pandas cannot even store,
# and would otherwise abort the whole run.
DATE_MIN = pd.Timestamp("2000-01-01")
DATE_MAX = pd.Timestamp("2100-01-01")

# Filled by normalise(), read by exceptions() for the data quality table.
DATE_ISSUES = {}


def _clean_parsed(parsed) -> pd.Series:
    """Drop anything outside the plausible window and pin the unit to ns."""
    out = pd.Series(pd.to_datetime(parsed, errors="coerce"))
    ok = out.notna() & (out >= DATE_MIN) & (out <= DATE_MAX)
    return out.where(ok).astype("datetime64[ns]")


def to_date(s: pd.Series) -> pd.Series:
    """Vectorised date parsing. Each format is applied to the whole column at
    once; only leftovers reach the slow inferring parser. Values that parse to
    an impossible date are discarded rather than raising."""
    txt = s.astype(str).str.strip()
    txt = txt.where(~txt.isin(["", "nan", "None", "NaT", "-", "0"]))
    out = pd.Series(pd.NaT, index=txt.index, dtype="datetime64[ns]")

    for fmt in DATE_FORMATS:
        todo = out.isna() & txt.notna()
        if not todo.any():
            return out
        parsed = _clean_parsed(pd.to_datetime(txt[todo], format=fmt, errors="coerce"))
        out.loc[todo] = parsed.to_numpy()

    todo = out.isna() & txt.notna()
    if todo.any():
        try:
            parsed = pd.to_datetime(txt[todo], errors="coerce",
                                    dayfirst=True, format="mixed")
        except (ValueError, OutOfBoundsDatetime):
            parsed = pd.Series(pd.NaT, index=txt[todo].index, dtype="datetime64[ns]")
        out.loc[todo] = _clean_parsed(parsed).to_numpy()

    return out


def is_text(s: pd.Series) -> bool:
    """True for text columns on any pandas version. Pandas 3.0 gives strings
    the 'str' dtype rather than 'object', so an == object test silently skips
    every text column and the cleaning below never runs."""
    return s.dtype == object or pd.api.types.is_string_dtype(s)


def coalesce(df, cols):
    out = pd.Series(np.nan, index=df.index, dtype="float64")
    for c in cols:
        if c in df.columns:
            out = out.where(out.notna(), df[c])
    return out


def normalise(df: pd.DataFrame) -> pd.DataFrame:
    print(f"Normalising {len(df):,} rows...", flush=True)
    cols, seen = [], set()
    for c in df.columns:
        c2 = str(c).strip()
        if c2 and not c2.startswith("Unnamed") and c2 not in seen:
            seen.add(c2)
            cols.append((c, c2))

    # Rebuild the frame from raw arrays rather than slicing the original. A
    # sliced frame keeps a reference to its parent, and writing a column into
    # it is what pandas flags as chained assignment. Building fresh arrays
    # leaves nothing to reference, so the conversions below are plain writes.
    print(f"  rebuilding {len(cols)} columns...", flush=True)
    data = {}
    for orig, name in cols:
        col = df[orig]
        if isinstance(col, pd.DataFrame):          # duplicate header in the sheet
            col = col.iloc[:, 0]
        data[name] = col.to_numpy()
    df = pd.DataFrame(data)

    # Conversions are applied in one pass each, not column by column, so the
    # frame is rewritten twice rather than forty times.
    print("  converting numbers...", flush=True)
    converted = {c: to_num(df[c]) for c in NUMERIC_COLS if c in df.columns}
    print("  converting dates...", flush=True)
    DATE_ISSUES.clear()
    for c in DATE_COLS:
        if c in df.columns:
            print(f"    {c}", flush=True)
            converted[c] = to_date(df[c])
            filled = df[c].astype(str).str.strip().replace(
                {"": None, "nan": None, "None": None, "-": None, "0": None}).notna()
            broken = filled & converted[c].isna()
            bad = int(broken.sum())
            if bad:
                DATE_ISSUES[c] = bad
                samples = df.loc[broken, c].astype(str).unique()[:5]
                print(f"      {bad:,} cell(s) could not be read as a date "
                      f"-- left blank. Values seen: "
                      f"{', '.join(repr(v) for v in samples)}", flush=True)
    missing_num = {c: np.nan for c in NUMERIC_COLS if c not in df.columns}
    missing_date = {c: pd.NaT for c in DATE_COLS if c not in df.columns}
    df = df.assign(**converted, **missing_num, **missing_date)

    print("  cleaning text columns...", flush=True)
    text = {c: df[c].astype(str).str.strip().replace({"nan": "", "None": ""})
            for c in df.columns if is_text(df[c])}
    defaults = {c: "" for c in TEXT_COLS if c not in df.columns}
    if "month" not in df.columns:
        defaults["month"] = "Sample"
    df = df.assign(**text, **defaults)

    print("  deriving fields...", flush=True)
    df = df[df["Invoice No."].astype(str).str.len() > 0].copy()
    # Unique invoice identity = month + invoice number, so the same invoice
    # number appearing in two monthly sheets is never merged.
    df["inv_key"] = df["month"].astype(str) + "|" + df["Invoice No."].astype(str)

    # ---- party classification --------------------------------------------
    df["party_key"] = df["Display Name"].map(squash)
    df["open_po"] = df["party_key"].isin(OPEN_PO_KEYS)

    # ---- GRN quantity -----------------------------------------------------
    df["grn_qty_raw"] = coalesce(df, GRN_PRIORITY)
    # Only GRN Config = Yes accounts ever produce a GRN. Everything else is
    # outside the fill rate universe entirely.
    df["grn_applicable"] = df["GRN Config"].str.upper() == "YES"
    df["grn_configured"] = df["grn_applicable"]
    df["grn_qty"] = df["grn_qty_raw"].where(df["grn_applicable"])
    if not GRN_CONFIG_NO_FALLS_BACK_TO_POD:
        df["grn_qty"] = df["grn_qty_raw"]

    # A GRN has landed on this line only if there is a real quantity on it.
    df["grn_received"] = df["grn_applicable"] & df["grn_qty"].notna() & (df["grn_qty"] > 0)
    # Pending is an invoice-level state: if no line on the invoice carries a
    # GRN, the whole invoice is awaiting one and nothing on it is short.
    df["inv_has_grn"] = df.groupby("inv_key")["grn_received"].transform("any")
    df["grn_pending"] = df["grn_applicable"] & ~df["inv_has_grn"]

    # ---- received quantity ------------------------------------------------
    # Default GRN over POD; open-PO parties take POD over GRN.
    grn_first = coalesce(df, ["grn_qty", "POD Qty"])
    pod_first = coalesce(df, ["POD Qty", "grn_qty"])
    df["received_qty"] = np.where(df["open_po"], pod_first, grn_first)
    df["received_src"] = np.where(
        df["open_po"],
        np.where(df["POD Qty"].notna(), "POD", np.where(df["grn_qty"].notna(), "GRN", "none")),
        np.where(df["grn_qty"].notna(), "GRN", np.where(df["POD Qty"].notna(), "POD", "none")),
    )

    # ---- gaps -------------------------------------------------------------
    # Short only counts where a GRN was expected and one arrived. No GRN yet
    # means Pending GRN; no GRN ever means the account is out of scope.
    scoreable = df["grn_applicable"] & df["inv_has_grn"]
    df["scoreable"] = scoreable
    df["short_qty"] = ((df["Invoice Qty."] - df["received_qty"]).clip(lower=0)
                       ).where(scoreable, 0.0)
    df["excess_qty"] = (df["received_qty"] - df["Invoice Qty."]).clip(lower=0)
    df["short_value"] = df["short_qty"] * df["Unit Price"]
    df["po_inv_gap"] = (df["PO Qty"] - df["Invoice Qty."]).clip(lower=0)
    df["po_gap_value"] = df["po_inv_gap"] * df["Unit Price"]

    # POD vs GRN quantity gap -- only meaningful for non open-PO parties
    both = df["POD Qty"].notna() & df["grn_qty"].notna() & df["grn_applicable"]
    df["pod_grn_gap"] = np.where(both & ~df["open_po"], df["POD Qty"] - df["grn_qty"], np.nan)
    df["pod_grn_gap_value"] = df["pod_grn_gap"].abs() * df["Unit Price"]
    df["pod_grn_flag"] = df["pod_grn_gap"].fillna(0) != 0

    df["is_short"] = df["short_qty"] > 0
    # Cancellation is two separate buckets, not one combined test.
    #   Is Cancel = Yes      -> cancelled within 24 hours
    #   Return Reason = INC  -> cancelled after 24 hours
    # Within wins if both are set, so the buckets stay mutually exclusive and
    # still add up to the invoice count.
    df["cancel_24"] = df["Is Cancel"].str.upper() == "YES"
    df["cancel_post"] = (df["Return Reason"].str.upper() == "INC") & ~df["cancel_24"]
    df["cancel_both"] = df["cancel_24"] & (df["Return Reason"].str.upper() == "INC")
    df["is_cancelled"] = df["cancel_24"] | df["cancel_post"]
    # POD says cancelled but nothing marks it as such -- needs a human look.
    df["pod_cancel_unflagged"] = ((df["POD Status"].str.upper() == "CANCELLED")
                                  & (df["Return Reason"].str.upper() != "INC"))

    df["sro_raised"] = df["SRO No"].str.len() > 0
    df["cn_raised"] = (df["RTV DOC"].str.len() > 0) | (df["CN AMOUNT"].fillna(0) != 0)
    df["cn_value"] = df["CN AMOUNT"].fillna(0).abs()

    today = pd.Timestamp("today").normalize()
    df["ageing_days"] = (today - df["Posting Date"]).dt.days
    df["grn_tat_days"] = (df["Delivery/GRN Date"] - df["Posting Date"]).dt.days
    return df


# ============================================================================
# INVOICE ROLL-UP  -- every downstream count is on unique invoice numbers
# ============================================================================
def invoice_view(df: pd.DataFrame) -> pd.DataFrame:
    print("Rolling up to invoice level...", flush=True)
    inv = df.groupby(INVOICE_KEYS, dropna=False).agg(
        platform=("Display Name", "first"),
        location=("Customer location name", "first"),
        warehouse=("Warehouse", "first"),
        city=("Ship-to City", "first"),
        tier=("T1/T2", "first"),
        category=("CATAGORY", "first"),
        po_no=("PO. No.", "first"),
        po_date=("PO Date", "min"),
        posting_date=("Posting Date", "min"),
        grn_date=("Delivery/GRN Date", "max"),
        open_po=("open_po", "first"),
        grn_applicable=("grn_applicable", "any"),
        grn_configured=("grn_configured", "first"),
        grn_pending=("grn_pending", "any"),
        scoreable=("scoreable", "any"),
        lines=("Invoice No.", "size"),
        short_lines=("is_short", "sum"),
        po_qty=("PO Qty", "sum"),
        invoice_qty=("Invoice Qty.", "sum"),
        grn_qty=("grn_qty", "sum"),
        pod_qty=("POD Qty", "sum"),
        received_qty=("received_qty", "sum"),
        short_qty=("short_qty", "sum"),
        short_value=("short_value", "sum"),
        po_inv_gap=("po_inv_gap", "sum"),
        invoice_amt=("Invoice Amt", "sum"),
        pod_grn_gap_lines=("pod_grn_flag", "sum"),
        pod_grn_gap_value=("pod_grn_gap_value", "sum"),
        sro_no=("SRO No", "max"),
        rtv_doc=("RTV DOC", "max"),
        cn_value=("cn_value", "sum"),
        is_cancelled=("is_cancelled", "any"),
        cancel_24=("cancel_24", "any"),
        cancel_post=("cancel_post", "any"),
        pod_cancel_unflagged=("pod_cancel_unflagged", "any"),
        ageing_days=("ageing_days", "max"),
        grn_tat_days=("grn_tat_days", "max"),
        has_pod_line=("POD Qty", lambda s: bool(s.notna().any())),
        has_grn_line=("grn_qty", lambda s: bool((s.fillna(0) > 0).any())),
    ).reset_index()

    inv["is_short"] = (inv["short_lines"] > 0) & inv["scoreable"]
    inv["grn_state"] = np.select(
        [~inv["grn_applicable"], inv["grn_pending"], inv["is_short"]],
        ["No GRN expected", "Pending GRN", "Short GRN"], default="GRN complete")
    inv["sro_raised"] = inv["sro_no"].astype(str).str.len() > 0
    inv["cn_raised"] = (inv["rtv_doc"].astype(str).str.len() > 0) | (inv["cn_value"] > 0)
    inv["unrecovered_value"] = (inv["short_value"] - inv["cn_value"]).clip(lower=0)
    inv["sro_status"] = np.where(inv["sro_raised"], "SRO raised", "SRO pending")
    inv["cn_status"] = np.where(inv["cn_raised"], "CN raised", "CN pending")
    inv["ageing_bucket"] = pd.cut(
        inv["ageing_days"], bins=[-np.inf, 3, 7, 15, 30, np.inf],
        labels=["0-3 d", "4-7 d", "8-15 d", "16-30 d", "30+ d"])

    def bucket(r):
        if r["has_pod_line"] and r["has_grn_line"]:
            return "GRN & POD received"
        if r["has_grn_line"]:
            return "GRN received only"
        if r["has_pod_line"]:
            return "POD received only"
        return "Neither received"
    inv["receipt_bucket"] = inv.apply(bucket, axis=1)
    return inv


# ============================================================================
# METRICS
# ============================================================================
def fill_rate(df: pd.DataFrame, inv: pd.DataFrame, by=None) -> pd.DataFrame:
    # Fill rate lives only where GRNs are received at all.
    d = df[~df["is_cancelled"] & df["grn_applicable"]].copy()
    i = inv[~inv["is_cancelled"] & inv["grn_applicable"]].copy()
    if EXCLUDE_PENDING_GRN_FROM_FILL_RATE:
        d = d[~d["grn_pending"]]
        i = i[~i["grn_pending"]]

    if by:
        g = d.groupby(by, dropna=False)
    else:
        g = d.groupby(lambda _: "OVERALL")

    out = g.agg(
        lines=("Invoice No.", "size"),
        invoices=("inv_key", "nunique"),
        po_qty=("PO Qty", "sum"),
        invoice_qty=("Invoice Qty.", "sum"),
        received_qty=("received_qty", "sum"),
        short_qty=("short_qty", "sum"),
        short_value=("short_value", "sum"),
        open_po=("open_po", "any"),
    )
    # PO -> GRN uses PO Qty. Open-PO parties have no committed PO quantity, so
    # that ratio is left blank for them rather than reported as a fill rate.
    out["po_to_grn_pct"] = np.where(
        out["open_po"], np.nan, (out["received_qty"] / out["po_qty"] * 100).round(2))
    out["inv_to_grn_pct"] = (out["received_qty"] / out["invoice_qty"] * 100).round(2)
    out["po_to_inv_pct"] = np.where(
        out["open_po"], np.nan, (out["invoice_qty"] / out["po_qty"] * 100).round(2))
    cols = ["invoices", "lines", "po_qty", "invoice_qty", "received_qty",
            "po_to_grn_pct", "inv_to_grn_pct", "po_to_inv_pct",
            "short_qty", "short_value", "open_po"]
    return out[cols].reset_index()


def claims(inv: pd.DataFrame):
    short = inv[inv["is_short"] & ~inv["is_cancelled"]].copy()
    summary = short.groupby(["sro_status", "cn_status"], observed=False).agg(
        invoices=("Invoice No.", "size"),
        short_lines=("short_lines", "sum"),
        short_qty=("short_qty", "sum"),
        short_value=("short_value", "sum"),
        cn_value=("cn_value", "sum"),
        unrecovered_value=("unrecovered_value", "sum"),
    ).reset_index()

    ageing = short[~short["cn_raised"]].groupby("ageing_bucket", observed=False).agg(
        invoices=("Invoice No.", "size"),
        short_qty=("short_qty", "sum"),
        unrecovered_value=("unrecovered_value", "sum"),
    ).reset_index()

    detail = short.loc[~short["sro_raised"] | ~short["cn_raised"]].sort_values(
        "unrecovered_value", ascending=False)
    return summary, ageing, detail


def pod_tracker(inv: pd.DataFrame):
    d = inv[~inv["is_cancelled"]]
    summary = d.groupby("receipt_bucket").agg(
        invoices=("Invoice No.", "size"),
        lines=("lines", "sum"),
        invoice_qty=("invoice_qty", "sum"),
        received_qty=("received_qty", "sum"),
    ).reset_index()
    total = summary["invoices"].sum()
    summary["pct_of_invoices"] = (summary["invoices"] / total * 100).round(2) if total else 0.0
    by_platform = pd.crosstab(d["platform"], d["receipt_bucket"]).reset_index()
    return summary, by_platform


def cancelled(df: pd.DataFrame, inv: pd.DataFrame):
    rows = [
        ("Cancelled within 24 hours (Is Cancel = Yes)", df["cancel_24"]),
        ("Cancelled after 24 hours (Return Reason = INC)", df["cancel_post"]),
        ("Both flags set -- counted as within 24 hours", df["cancel_both"]),
        ("POD Status = CANCELLED but Return Reason is not INC", df["pod_cancel_unflagged"]),
    ]
    flags = pd.DataFrame([{
        "condition": label,
        "invoices": int(df.loc[m, "inv_key"].nunique()),
        "lines": int(m.sum()),
        "invoice_value": float(df.loc[m, "Invoice Amt"].sum()),
    } for label, m in rows])
    detail = inv[inv["is_cancelled"] | inv["pod_cancel_unflagged"]].copy()
    detail["cancel_bucket"] = np.select(
        [detail["cancel_24"], detail["cancel_post"]],
        ["Within 24 hours", "After 24 hours"], default="Not flagged as cancelled")
    return flags, detail


def exceptions(df: pd.DataFrame, inv: pd.DataFrame):
    d = df[~df["is_cancelled"]].copy()
    out = {}

    gap = d[d["pod_grn_flag"]].copy()
    gap["direction"] = np.where(gap["pod_grn_gap"] > 0,
                                "POD > GRN (delivered, not booked)",
                                "GRN > POD (booked above delivery)")
    out["POD vs GRN gap"] = gap[[c for c in [
        "month", "Posting Date", "Display Name", "Customer location name",
        "Invoice No.", "Item/Account", "Description", "Invoice Qty.", "POD Qty",
        "grn_qty", "pod_grn_gap", "pod_grn_gap_value", "direction",
        "GRN Status", "POD Status", "SRO No", "RTV DOC"]
        if c in gap.columns]].sort_values("pod_grn_gap_value", ascending=False)

    out["PO not fully invoiced"] = d.loc[d["po_inv_gap"] > 0, [c for c in [
        "month", "PO Date", "Display Name", "Customer location name", "PO. No.",
        "Invoice No.", "Description", "PO Qty", "Invoice Qty.", "po_inv_gap",
        "po_gap_value"] if c in d.columns]].sort_values("po_gap_value", ascending=False)

    out["Over receipt"] = d.loc[d["excess_qty"] > 0, [c for c in [
        "month", "Display Name", "Invoice No.", "Description", "Invoice Qty.",
        "received_qty", "excess_qty", "received_src"] if c in d.columns]]

    out["GRN pending"] = inv.loc[inv["grn_pending"] & ~inv["is_cancelled"], [
        "month", "Invoice No.", "platform", "location", "posting_date",
        "invoice_qty", "pod_qty", "grn_qty", "ageing_days"]].sort_values(
        "ageing_days", ascending=False)

    out["GRN state by platform"] = (inv[~inv["is_cancelled"]]
        .groupby(["platform", "grn_state"], dropna=False)
        .agg(invoices=("Invoice No.", "size"), invoice_qty=("invoice_qty", "sum"))
        .reset_index())

    flagged = df[df["pod_cancel_unflagged"]]
    out["POD cancelled without INC"] = flagged[[c for c in [
        "month", "Posting Date", "Warehouse", "Display Name", "Customer location name",
        "Invoice No.", "Description", "Invoice Qty.", "Invoice Amt", "POD Status",
        "Is Cancel", "Return Reason"] if c in flagged.columns]]

    out["Short by SKU"] = (d[d["is_short"]]
        .groupby(["Item/Account", "Description"], dropna=False)
        .agg(invoices=("inv_key", "nunique"), short_qty=("short_qty", "sum"),
             short_value=("short_value", "sum"), invoice_qty=("Invoice Qty.", "sum"))
        .assign(short_rate_pct=lambda x: (x.short_qty / x.invoice_qty * 100).round(2))
        .sort_values("short_value", ascending=False).reset_index())

    out["Short by location"] = (d[d["is_short"]]
        .groupby(["Display Name", "Customer location name"], dropna=False)
        .agg(invoices=("inv_key", "nunique"), short_qty=("short_qty", "sum"),
             short_value=("short_value", "sum"))
        .sort_values("short_value", ascending=False).reset_index())

    checks = []
    dup = d["Concat"].duplicated(keep=False) & (d["Concat"].str.len() > 0)
    checks.append(("Duplicate Concat key", int(dup.sum())))
    checks += [
        ("Blank or zero Unit Price", int((d["Unit Price"].isna() | (d["Unit Price"] == 0)).sum())),
        ("Blank Invoice Qty", int(d["Invoice Qty."].isna().sum())),
        ("Neither GRN nor POD qty", int(pd.isna(d["received_qty"]).sum())),
        ("Blank Delivery/GRN Date", int(d["Delivery/GRN Date"].isna().sum())),
        ("GRN turnaround over 7 days", int((d["grn_tat_days"] > 7).sum())),
    ]
    odd = d.loc[~d["GRN Config"].str.upper().isin(["YES", "NO"]), "GRN Config"]
    checks.append(("GRN Config neither Yes nor No", int(len(odd))))
    for col, n in DATE_ISSUES.items():
        checks.append((f"Unreadable date in {col}", n))
    # A date that parses cleanly can still be wrong. "1-08-09" reads as a valid
    # 2009 date; against a posting date in 2026 it is obviously a typo.
    for col in ["PO Date", "Delivery/GRN Date"]:
        off = (d[col] - d["Posting Date"]).abs() > pd.Timedelta(days=180)
        checks.append((f"{col} more than 180 days from posting date",
                       int(off.fillna(False).sum())))
    out["Data quality"] = pd.DataFrame(checks, columns=["check", "rows"])
    return out


# ============================================================================
# ORCHESTRATION
# ============================================================================
def compute(df: pd.DataFrame, inv: pd.DataFrame):
    print(f"Computing metrics on {len(inv):,} invoices...", flush=True)
    res = {
        "fill_overall": fill_rate(df, inv),
        "fill_month": fill_rate(df, inv, ["month"]),
        "fill_platform": fill_rate(df, inv, ["Display Name"]),
        "fill_location": fill_rate(df, inv, ["Display Name", "Customer location name"]),
        "fill_sku": fill_rate(df, inv, ["Item/Account", "Description"]),
    }
    res["claims_summary"], res["claims_ageing"], res["claims_detail"] = claims(inv)
    res["pod_summary"], res["pod_platform"] = pod_tracker(inv)
    res["cancel_flags"], res["cancel_detail"] = cancelled(df, inv)
    res["exceptions"] = exceptions(df, inv)
    res["headline"] = headline(df, inv)
    return res


def headline(df: pd.DataFrame, inv: pd.DataFrame) -> dict:
    """Small dict of top-line numbers for the dashboard hero and KPI strip."""
    everything = inv[~inv["is_cancelled"]]
    d = df[~df["is_cancelled"] & df["grn_applicable"] & df["scoreable"]]
    i = inv[~inv["is_cancelled"] & inv["grn_applicable"] & inv["scoreable"]]
    closed = d[~d["open_po"]]          # PO quantity is only meaningful here
    po_qty = float(closed["PO Qty"].sum())
    po_received = float(closed["received_qty"].sum())
    po_invoiced = float(closed["Invoice Qty."].sum())
    short_inv = i[i["is_short"]]
    return {
        "invoices": int(len(i)),
        "all_invoices": int(len(everything)),
        "not_applicable_invoices": int((~everything["grn_applicable"]).sum()),
        "lines": int(len(d)),
        "po_qty": po_qty,
        "po_invoiced_qty": po_invoiced,
        "po_received_qty": po_received,
        "invoice_qty": float(d["Invoice Qty."].sum()),
        "received_qty": float(d["received_qty"].sum()),
        "po_to_grn_pct": round(po_received / po_qty * 100, 2) if po_qty else None,
        "inv_to_grn_pct": round(float(d["received_qty"].sum()) / float(d["Invoice Qty."].sum()) * 100, 2)
                          if d["Invoice Qty."].sum() else None,
        "po_to_inv_pct": round(po_invoiced / po_qty * 100, 2) if po_qty else None,
        "short_invoices": int(len(short_inv)),
        "short_qty": float(i["short_qty"].sum()),
        "short_value": float(i["short_value"].sum()),
        "recovered_value": float(i["cn_value"].sum()),
        "unrecovered_value": float(i["unrecovered_value"].sum()),
        "sro_pending_invoices": int((~short_inv["sro_raised"]).sum()),
        "cn_pending_invoices": int((~short_inv["cn_raised"]).sum()),
        "pod_grn_gap_invoices": int((everything["pod_grn_gap_lines"] > 0).sum()),
        "pod_grn_gap_value": float(everything["pod_grn_gap_value"].sum()),
        "grn_pending_invoices": int(everything["grn_pending"].sum()),
        "cancelled_invoices": int(inv["is_cancelled"].sum()),
        "cancelled_within_24": int(inv["cancel_24"].sum()),
        "cancelled_after_24": int(inv["cancel_post"].sum()),
        "pod_cancel_unflagged": int(inv["pod_cancel_unflagged"].sum()),
        "cancelled_value": float(inv[inv["is_cancelled"]]["invoice_amt"].sum()),
        "open_po_invoices": int(i["open_po"].sum()),
    }


def build_all(df_raw: pd.DataFrame):
    df = normalise(df_raw)
    inv = invoice_view(df)
    return df, inv, compute(df, inv)


def build_by_month(df: pd.DataFrame, inv: pd.DataFrame, everything):
    """One result set per month, plus a combined 'All months' set."""
    out = {"All months": everything}
    months = sorted(df["month"].unique())
    if len(months) > 1:
        for m in months:
            out[m] = compute(df[df["month"] == m], inv[inv["month"] == m])
    elif months:
        out = {months[0]: everything}
    return out


def main():
    # The frame is rebuilt from scratch in normalise(), so nothing here relies
    # on a write reaching a parent object. Some pandas builds still flag the
    # pattern; silence that one message so real warnings stay visible.
    warnings.filterwarnings("ignore", message=".*ChainedAssignment.*")

    ap = argparse.ArgumentParser()
    ap.add_argument("--csv", help="read a CSV export instead of the live sheets")
    ap.add_argument("--month", default="Sample", help="month label when using --csv")
    ap.add_argument("--rows", type=int, help="only process the first N rows (for testing)")
    ap.add_argument("--xlsx", default=OUTPUT_XLSX)
    ap.add_argument("--html", default=OUTPUT_HTML)
    args = ap.parse_args()

    if args.csv:
        raw = load_from_csv(args.csv, args.month)
    else:
        print("Reading Google Sheets...")
        raw = load_from_sheets()

    if args.rows:
        raw = raw.head(args.rows)
        print(f"Limited to the first {len(raw):,} rows.", flush=True)

    df, inv, res = build_all(raw)
    print(f"\nLines: {len(df)}   Invoices: {len(inv)}   Months: {inv['month'].nunique()}")
    print("\nFILL RATE")
    print(res["fill_overall"].to_string(index=False))
    print("\nBY PLATFORM")
    print(res["fill_platform"].to_string(index=False))
    print("\nCLAIMS (unique invoices)")
    print(res["claims_summary"].to_string(index=False))
    print("\nPOD TRACKER")
    print(res["pod_summary"].to_string(index=False))
    print("\nCANCELLED")
    print(res["cancel_flags"].to_string(index=False))
    for k, v in res["exceptions"].items():
        print(f"-- {k}: {len(v)} rows")

    sheets = {
        "Fill_Overall": res["fill_overall"], "Fill_Month": res["fill_month"],
        "Fill_Platform": res["fill_platform"], "Fill_Location": res["fill_location"],
        "Fill_SKU": res["fill_sku"], "Claims_Summary": res["claims_summary"],
        "Claims_Ageing": res["claims_ageing"], "Claims_Pending": res["claims_detail"],
        "POD_Summary": res["pod_summary"], "POD_ByPlatform": res["pod_platform"],
        "Cancelled_Flags": res["cancel_flags"], "Cancelled_Detail": res["cancel_detail"],
        "Invoice_Master": inv,
    }
    for k, v in res["exceptions"].items():
        sheets[k[:31]] = v
    print("Writing workbook...", flush=True)
    with pd.ExcelWriter(args.xlsx, engine="openpyxl") as xw:
        for name, frame in sheets.items():
            frame.to_excel(xw, sheet_name=str(name)[:31], index=False)
    print(f"\nWritten: {args.xlsx}")

    from dashboard import build_dashboard
    build_dashboard(build_by_month(df, inv, res), args.html, line_df=df)
    print(f"Written: {args.html}")


if __name__ == "__main__":
    main()
