"""
ShipBob D2C engine - reads the order Dump sheets from the G: drive workbooks,
classifies every shipment and builds the numbers the dashboard shows.

Everything comes from the Excel (no invented brackets):
  Delivered                  Delivery Status starts with "Delivered", or the team marked Sub Bucket = Delivered
  In Transit                 not delivered, <= transit_days (15) since Import Date
                               (blank Delivery Status = no carrier scan yet = In Transit)
  Potential In-Transit Loss  the In Transit orders at day alert_from_day..transit_days (11-15)
                               -> these go to ShipBob in the mailer (same orders stay in In Transit)
  Under Dispute              not delivered, > transit_days since Import Date, split by the team's
                               Sub Bucket column into
                               Pending to claim / Claim raised but not received / Claim raised and received
                               and "Claim Window Expired" (only what the team marked in Sub Bucket)
                             An order the team marked "Claim raised and received" or "Claim Window Expired"
                               stays here even if the carrier later shows Delivered
  RTO, Cancelled             from the team's Main/Sub Bucket columns, Cancelled also from Order Status
  The "Claim status" column (AB) is NOT used - it is not maintained.

Counting unit: the Excel has one row per line item, so an order with 2 products has 2 rows.
  unit="lines"  counts rows (same as the Excel pivot);  unit="orders" counts each shipment once.
"""
import glob
import json
import os
import pickle
import re
from datetime import datetime

import numpy as np
import pandas as pd

BASE = os.path.dirname(os.path.abspath(__file__))
CACHE_DIR = os.path.join(BASE, "cache")

MAIN_ORDER = ["Delivered", "In Transit", "Under Dispute", "RTO", "Cancelled"]
POTENTIAL = "Potential In-Transit Loss"
EXPIRED = "Claim Window Expired"
SUB_ORDER = ["Pending to claim", "Claim raised but not received", "Claim raised and received"]
SUB_ALL = SUB_ORDER + [EXPIRED]
STAGE_ORDER = ["Blank (no scan yet)", "PreTransit", "InTransit", "OutForDelivery", "DeliveryAttemptFailed"]
MAIL_LOG = os.path.join(BASE, "state", "mailer_log.csv")

COLS = ["Shipment ID", "Store Order ID", "Import Date", "Purchase Date", "Source", "Order Status",
        "Carrier", "Tracking Information", "Fulfillment Center", "Fulfillment Cost", "Line Item Name",
        "Line Item Qty", "Total Item Quantity", "Delivery Status", "Ingestion Channel", "Claim status",
        "Amt", "Expected Claim amount", "Claims remark", "Sub Bucket", "Main Bucket"]


def load_config():
    with open(os.path.join(BASE, "config.json")) as f:
        return json.load(f)


# ------------------------------------------------------------------ sources
def find_sources(cfg):
    """Pinned closed-period files + the most recently modified other workbook
    (same rule as recon-dashboard/upload_shipbob_d2c.py)."""
    folder = cfg["source_folder"]
    pinned = [os.path.join(folder, n) for n in cfg["pinned_files"]]
    missing = [p for p in pinned if not os.path.exists(p)]
    if missing:
        raise FileNotFoundError(f"Pinned file(s) not found: {missing}")
    others = [f for f in glob.glob(os.path.join(folder, "*.xlsx"))
              if not os.path.basename(f).startswith("~$") and f not in pinned]
    if not others:
        raise FileNotFoundError(f"No current ShipBob workbook in {folder}")
    return pinned + [max(others, key=os.path.getmtime)]


def _read_dump(path):
    """Read the Dump sheet, cached by file mtime (a 20 MB workbook takes ~1 min)."""
    os.makedirs(CACHE_DIR, exist_ok=True)
    stem = re.sub(r"[^\w]+", "_", os.path.basename(path))
    key = os.path.join(CACHE_DIR, f"{stem}_{int(os.path.getmtime(path))}.pkl")
    if os.path.exists(key):
        with open(key, "rb") as f:
            return pickle.load(f)
    xl = pd.ExcelFile(path)
    dump = pd.read_excel(xl, sheet_name="Dump", usecols=lambda c: str(c).strip() in COLS)
    dump.columns = [str(c).strip() for c in dump.columns]
    prices = {}
    price_sheet = next((s for s in xl.sheet_names if s.lower().startswith("product price")), None)
    if price_sheet:
        p = pd.read_excel(xl, sheet_name=price_sheet, header=None).dropna()
        prices = {_norm(n): float(v) for n, v in p.values if isinstance(v, (int, float))}
    out = (dump, prices)
    for old in glob.glob(os.path.join(CACHE_DIR, f"{stem}_*.pkl")):
        os.remove(old)
    with open(key, "wb") as f:
        pickle.dump(out, f)
    return out


# ------------------------------------------------------------------ Google Sheet source (read-only)
SHEET_TAB = "ShipBob D2C Claims"
SYNC_TAB = "ShipBob D2C Sync Info"
PRICES_FILE = os.path.join(BASE, "product_prices.csv")


def uses_sheet(cfg):
    return str(cfg.get("data_source", "excel")).lower() == "sheet"


def _sheets_values(cfg, rng, render="UNFORMATTED_VALUE"):
    """values.get only - this module never writes to the sheet."""
    import sys
    app_dir = os.path.dirname(BASE)
    if app_dir not in sys.path:
        sys.path.insert(0, app_dir)
    from auth import get_sheets_credentials
    from googleapiclient.discovery import build
    creds = get_sheets_credentials(cfg.get("sheets_credentials_file", cfg.get("credentials_file")),
                                   cfg.get("sheets_token_file", "C:/Users/Admin/recon-token.json"))
    svc = build("sheets", "v4", credentials=creds, cache_discovery=False)
    r = svc.spreadsheets().values().get(spreadsheetId=cfg["sheet_id"], range=rng,
                                        valueRenderOption=render).execute()
    return r.get("values", [])


def sheet_sync_info(cfg):
    """(synced_at, file name, row count) from the 'ShipBob D2C Sync Info' tab - written by the upload script."""
    v = _sheets_values(cfg, f"'{SYNC_TAB}'!A1:C1", "FORMATTED_VALUE")
    row = (v[0] if v else []) + ["", "", ""]
    return str(row[0]), str(row[1]), str(row[2])


def _read_prices():
    if not os.path.exists(PRICES_FILE):
        return {}
    p = pd.read_csv(PRICES_FILE)
    return {_norm(n): float(v) for n, v in zip(p["Product"], p["Price"]) if pd.notna(v)}


def _read_sheet(cfg):
    """All rows of the 'ShipBob D2C Claims' tab, columns matched by header name (order does not matter)."""
    v = _sheets_values(cfg, f"'{SHEET_TAB}'!A1:Z200000")
    if len(v) < 2:
        raise RuntimeError(f"Sheet tab '{SHEET_TAB}' is empty")
    hdr = [str(c).strip() for c in v[0]]
    need = ["Shipment ID", "Import Date", "Line Item Name", "Delivery Status", "Sub Bucket", "Main Bucket"]
    miss = [c for c in need if c not in hdr]
    if miss:
        raise RuntimeError(f"Sheet tab is missing column(s): {miss}")
    w = len(hdr)
    d = pd.DataFrame([(r + [""] * w)[:w] for r in v[1:]], columns=hdr)
    d = d.replace("", np.nan)
    for c in COLS:                      # columns the sheet does not carry (not used by any number) stay empty
        if c not in d.columns:
            d[c] = np.nan
    d["Shipment ID"] = pd.to_numeric(d["Shipment ID"], errors="coerce")
    d = d[d["Shipment ID"].notna()].copy()
    d["Shipment ID"] = d["Shipment ID"].astype("int64")
    d["Import Date"] = pd.to_datetime(d["Import Date"], errors="coerce")
    return d[COLS]


def _norm(name):
    """Match key for product names: ignores case, spacing, punctuation, the 'I'
    used in place of '|', a trailing 's' and the Keratin/Keratine spelling."""
    s = re.sub(r"\s+i\s+", " ", str(name).lower()).replace("keratine", "keratin")
    s = re.sub(r"[^a-z0-9]", "", s)
    return s[:-1] if s.endswith("s") else s


# ------------------------------------------------------------------ classify
def build_frame(cfg):
    from_sheet = uses_sheet(cfg)
    parts, prices = [], {}
    if from_sheet:
        sources = [f"Google Sheet - {SHEET_TAB}"]
        dump = _read_sheet(cfg)
        dump["_src"] = 0
        parts.append(dump)
        prices = _read_prices()
    else:
        sources = find_sources(cfg)
        for i, path in enumerate(sources):
            dump, pr = _read_dump(path)
            dump = dump.copy()
            dump["_src"] = i
            parts.append(dump)
            prices.update(pr)
    raw = pd.concat(parts, ignore_index=True)
    raw = raw[raw["Shipment ID"].notna()]
    # a shipment appearing in two files: keep the newer file's rows
    newest = raw.groupby("Shipment ID")["_src"].transform("max")
    raw = raw[raw["_src"] == newest].copy()

    for c in ["Amt", "Expected Claim amount", "Fulfillment Cost", "Line Item Qty"]:
        raw[c] = pd.to_numeric(raw[c], errors="coerce")
    raw["_price"] = raw["Line Item Name"].map(lambda n: prices.get(_norm(n), np.nan))
    # Expected claim = Product price x Line Item Qty, summed over the shipment's line items
    raw["_cogs"] = raw["_price"] * raw["Line Item Qty"].fillna(1)
    raw["_no_price"] = raw["_price"].isna()
    raw["_item"] = raw["Line Item Name"].astype(str) + " x" + raw["Line Item Qty"].fillna(1).astype(int).astype(str)

    first = ["Store Order ID", "Import Date", "Purchase Date", "Source", "Order Status", "Carrier",
             "Tracking Information", "Fulfillment Center", "Fulfillment Cost", "Delivery Status",
             "Claim status", "Claims remark", "Sub Bucket", "Main Bucket", "Total Item Quantity"]
    g = raw.groupby("Shipment ID", sort=False)
    df = g[first].first()
    df["Lines"] = g.size()                      # rows (line items) per shipment
    df["Items"] = g["_item"].agg(" + ".join)
    df["Amt"] = g["Amt"].sum(min_count=1)
    df["Expected"] = g["Expected Claim amount"].sum(min_count=1)
    df["Expected Calc"] = g["_cogs"].sum(min_count=1)
    df["Price Missing"] = g["_no_price"].any()
    df = df.reset_index()

    # one row per line item, for the "Line items" view of the lists
    lines = pd.DataFrame({"Shipment ID": raw["Shipment ID"].values, "Item": raw["_item"].values,
                          "L Expected": raw["_cogs"].fillna(0).round(2).values,
                          "L Sheet": raw["Expected Claim amount"].fillna(0).round(2).values,
                          "L Received": raw["Amt"].fillna(0).round(2).values})

    df["Import Date"] = pd.to_datetime(df["Import Date"], errors="coerce")
    as_of = df["Import Date"].max()
    ds = df["Delivery Status"].fillna("").astype(str).str.strip()
    df["Stage"] = ds.str.split(" at ").str[0].replace("", "Blank (no scan yet)")
    df["Age"] = ((as_of - df["Import Date"]).dt.total_seconds() // 86400).astype("Int64")
    df["Month"] = df["Import Date"].dt.month_name()
    df["Month No"] = df["Import Date"].dt.year * 100 + df["Import Date"].dt.month

    low = lambda c: df[c].fillna("").astype(str).str.strip().str.lower()
    mm, ms = low("Main Bucket"), low("Sub Bucket")
    T, A = cfg["transit_days"], cfg["alert_from_day"]
    age = df["Age"].fillna(0).astype(int)
    # the claim can only be raised after the in-transit window, so the claim raised date is taken as import date + T days
    df["Claim Age"] = (age - T).clip(lower=0)

    cancelled = (df["Order Status"] == "Cancelled") | (mm == "cancelled") | (ms == "cancelled")
    rto = (mm == "rto") | (ms == "rto")
    delivered = (df["Stage"] == "Delivered") | (ms == "delivered")
    # claim money already received / window marked expired by the team stays in Under Dispute
    kept = (ms == "claim raised and received") | (ms == "claim window expired")
    df["Main"] = np.select([cancelled, rto, kept, delivered, age <= T],
                           ["Cancelled", "RTO", "Under Dispute", "Delivered", "In Transit"], "Under Dispute")
    df["Potential Loss"] = (df["Main"] == "In Transit") & (age >= A)

    # dispute sub-bucket: only from the team's Sub Bucket column (blank / anything else = Pending to claim)
    sub = np.select([ms == "claim raised and received", ms == "claim raised but not received",
                     ms == "claim window expired"],
                    ["Claim raised and received", "Claim raised but not received", EXPIRED], "Pending to claim")
    df["Sub"] = np.where(df["Main"] == "Under Dispute", sub,
                         np.where(df["Potential Loss"], POTENTIAL, df["Main"]))

    # price x qty from the Product price tab; if a product has no price, fall back
    # to the team's Expected Claim amount from the sheet (if entered)
    sheet_exp = df["Expected"].fillna(0)
    use_sheet = df["Price Missing"] & (sheet_exp > 0)
    calc = np.where(use_sheet, sheet_exp, df["Expected Calc"].fillna(0)).round(2)
    df["Expected Calc $"] = calc                                    # price x qty (Product price tab)
    open_stage = df["Sub"].isin([SUB_ORDER[0], SUB_ORDER[1], EXPIRED])
    # Pending / Raised / Expired use the Excel's own Expected Claim amount (same numbers as the recon dashboard);
    # every other order keeps price x qty.
    df["Expected Claim $"] = np.where(open_stage, sheet_exp.round(2), calc)
    df["Expected Missing"] = df["Sub"].isin([SUB_ORDER[0], SUB_ORDER[1]]) & (sheet_exp <= 0)
    df["Received $"] = df["Amt"].fillna(0).round(2)

    # orders already highlighted to ShipBob through the mailer
    mlog = load_mail_log()
    df["Mail Kind"] = mail_kind_of(df)
    entries = [mlog.get(k, {}).get(int(i)) if k else None for k, i in zip(df["Mail Kind"], df["Shipment ID"])]
    df["Mailed"] = [_mailed_text(e) for e in entries]
    df["Mailed Status"] = [(e or {}).get("status", "") for e in entries]

    miss = raw[raw["_no_price"]].merge(df[["Shipment ID", "Main"]], on="Shipment ID")
    missing_prices = (miss.groupby("Line Item Name")
                      .agg(line_items=("Shipment ID", "size"),
                           in_dispute=("Main", lambda m: int((m == "Under Dispute").sum())))
                      .sort_values("line_items", ascending=False).reset_index()
                      .rename(columns={"Line Item Name": "product"}).to_dict("records"))
    if from_sheet:
        at, fname, _n = sheet_sync_info(cfg)
        try:
            synced = datetime.strptime(at, "%Y-%m-%d %H:%M:%S")
        except ValueError:
            synced = datetime.now()
        files = [{"name": f"Google Sheet ({fname})" if fname else sources[0], "modified": synced}]
        sources = [f["name"] for f in files]
    else:
        files = [{"name": os.path.basename(f), "modified": datetime.fromtimestamp(os.path.getmtime(f))} for f in sources]
    inp = []   # newest export dropped in Input folder (it is moved to Processed once refreshed)
    for folder in (() if from_sheet else ("Input folder", os.path.join("Input folder", "Processed"))):
        for ext in ("*.csv", "*.xlsx"):
            inp += glob.glob(os.path.join(cfg["source_folder"], folder, ext))
    inp = [f for f in inp if not os.path.basename(f).startswith("~$")]
    latest_input = max(inp, key=os.path.getmtime) if inp else None
    meta = {"missing_prices": missing_prices, "as_of": as_of, "sources": [os.path.basename(s) for s in sources], "data_source": "sheet" if from_sheet else "excel",
            "files": files, "lines": lines,
            "latest_input": ({"name": os.path.basename(latest_input),
                              "modified": datetime.fromtimestamp(os.path.getmtime(latest_input))}
                             if latest_input else None),
            "loaded_at": datetime.now(),
            "settings": {"transit_days": T, "alert_from_day": A, "claim_sla_days": cfg.get("claim_sla_days", [45, 60])}}
    return df, meta


# ------------------------------------------------------------------ mail log
MAIL_KINDS = ("potential", "pending", "raised")


def _read_log():
    if not os.path.exists(MAIL_LOG):
        return pd.DataFrame(columns=["Shipment ID", "Sent On", "Status", "Kind"])
    log = pd.read_csv(MAIL_LOG)
    if "Status" not in log.columns:
        log["Status"] = "drafted"
    if "Kind" not in log.columns:
        log["Kind"] = "potential"
    return log


def load_mail_log():
    """{kind: {shipment id: {"date": first date mailed (ISO), "status": "sent" or "drafted"}}}.
    A shipment that was ever marked sent (for that mail) counts as sent."""
    log = _read_log()
    out = {k: {} for k in MAIL_KINDS}
    if log.empty:
        return out
    log["_rank"] = (log["Status"] == "sent").astype(int)
    log = log.sort_values(["_rank", "Sent On"], ascending=[False, True]).drop_duplicates(["Kind", "Shipment ID"])
    for _, r in log.iterrows():
        out.setdefault(r["Kind"], {})[int(r["Shipment ID"])] = {"date": r["Sent On"], "status": r["Status"]}
    return out


def append_mail_log(shipment_ids, sent_on, status="drafted", kind="potential"):
    os.makedirs(os.path.dirname(MAIL_LOG), exist_ok=True)
    new = pd.DataFrame({"Shipment ID": list(shipment_ids), "Sent On": sent_on, "Status": status, "Kind": kind})
    log = pd.concat([_read_log(), new], ignore_index=True)
    log.drop_duplicates(["Shipment ID", "Sent On", "Status", "Kind"]).to_csv(MAIL_LOG, index=False)


def _mailed_text(entry):
    if not entry:
        return ""
    d = datetime.strptime(str(entry["date"])[:10], "%Y-%m-%d").strftime("%d %b")
    return f'{"Sent" if entry["status"] == "sent" else "Drafted"} {d}'


def mail_kind_of(df):
    """Which mail an order belongs to right now."""
    return np.where(df["Potential Loss"], "potential",
                    np.where(df["Sub"] == "Pending to claim", "pending",
                             np.where(df["Sub"] == "Claim raised but not received", "raised", "")))


# ------------------------------------------------------------------ lists
def _rows(d, cols):
    out = d[cols].copy()
    for c in out.columns:
        if pd.api.types.is_datetime64_any_dtype(out[c]):
            out[c] = out[c].dt.strftime("%Y-%m-%d")
    return json.loads(out.to_json(orient="records"))


BASE_COLS = ["Shipment ID", "Store Order ID", "Import Date", "Age", "Carrier", "Tracking Information",
             "Delivery Status", "Sub"]
LIST_DEF = {  # list id: (filter, extra columns, (sort column, ascending))
    "potential": (lambda d: d["Potential Loss"], ["Mailed"], ("Age", False)),
    "pending": (lambda d: d["Sub"] == "Pending to claim",
                ["Expected Claim $", "Mailed", "Claims remark"], ("Age", False)),
    "raised": (lambda d: d["Sub"] == "Claim raised but not received",
               ["Claim Age", "Expected Claim $", "Mailed", "Claims remark"], ("Age", False)),
    "expired": (lambda d: d["Sub"] == EXPIRED, ["Expected Claim $", "Claims remark"], ("Age", True)),
    "received": (lambda d: d["Sub"] == "Claim raised and received",
                 ["Expected Claim $", "Received $", "Claims remark"], ("Age", True)),
    "in_transit": (lambda d: d["Main"] == "In Transit", ["Mailed"], ("Age", False)),
    "rto": (lambda d: d["Main"] == "RTO", ["Claims remark"], ("Age", True)),
}
SHOW_LIMIT = 500   # rows sent to the page per list; Download Excel always has all rows


def list_frame(df, name, meta, unit="lines", limit=None):
    """Rows of one order list. unit="lines": one row per line item (matches the Excel);
    unit="orders": one row per shipment."""
    f, extra, (col, asc) = LIST_DEF[name]
    sel = df[f(df)].sort_values(col, ascending=asc)
    if limit:
        sel = sel.head(limit)
    if unit != "lines":
        return sel[BASE_COLS + extra]
    out = sel.merge(meta["lines"], on="Shipment ID", how="left")
    if "Expected Claim $" in extra:
        out["Expected Claim $"] = out["L Sheet"] if name in ("pending", "raised", "expired") else out["L Expected"]
    if "Received $" in extra:
        out["Received $"] = out["L Received"]
    cols = BASE_COLS + ["Item"] + extra
    out = out[cols]
    return out.head(limit) if limit else out


# ------------------------------------------------------------------ freshness
def freshness(meta):
    """What the 'data last updated' popup and badge show."""
    now = datetime.now()
    f = lambda t: t.strftime("%d %b %Y, %H:%M")
    newest = max(meta["files"], key=lambda x: x["modified"])
    log = load_mail_log()
    last_mail = max((v["date"] for k in log.values() for v in k.values()), default=None)
    data_age = (now.date() - meta["as_of"].date()).days
    file_age = (now.date() - newest["modified"].date()).days
    inp = meta["latest_input"]
    return {
        "data_as_of": meta["as_of"].strftime("%d %b %Y"), "as_of_iso": meta["as_of"].strftime("%Y-%m-%d"),
        "data_age_days": data_age,
        "file_name": newest["name"], "file_modified": f(newest["modified"]), "file_age_days": file_age,
        "input_name": inp["name"] if inp else None, "input_modified": f(inp["modified"]) if inp else None,
        "input_age_days": (now.date() - inp["modified"].date()).days if inp else None,
        "last_mailer": last_mail, "dashboard_loaded": f(meta["loaded_at"]),
        "status": "ok" if data_age <= 1 else "warn" if data_age <= 3 else "stale",
    }


# ------------------------------------------------------------------ dashboard payload
def build_payload(df, meta, unit="lines"):
    S = meta["settings"]
    W = df["Lines"].astype(int) if unit == "lines" else pd.Series(1, index=df.index)
    cnt = lambda mask: int(W[mask].sum())
    by = lambda m: cnt(df["Main"] == m)
    disp = df["Main"] == "Under Dispute"
    closed = by("Delivered") + by("Under Dispute")
    mailed = df["Mailed"] != ""
    exp_sum = lambda mask, col="Expected Claim $": round(float(df.loc[mask, col].sum()), 2)

    kpis = {
        "total": int(W.sum()), "delivered": by("Delivered"), "in_transit": by("In Transit"),
        "potential": cnt(df["Potential Loss"]), "potential_mailed": cnt(df["Potential Loss"] & mailed),
        "pending_oldest": int(df.loc[disp & (df["Sub"] == SUB_ORDER[0]), "Age"].max() or 0) if (disp & (df["Sub"] == SUB_ORDER[0])).any() else 0,
        "dispute": by("Under Dispute"), "rto": by("RTO"), "cancelled": by("Cancelled"),
        "delivery_rate": round(by("Delivered") / max(closed, 1) * 100, 2),
        "loss_rate": round(by("Under Dispute") / max(closed, 1) * 100, 2),
        "sub_counts": {s: cnt(disp & (df["Sub"] == s)) for s in SUB_ALL},
        "sub_value": {s: exp_sum(disp & (df["Sub"] == s)) for s in SUB_ALL},
        "received_amt": exp_sum(disp & (df["Sub"] == SUB_ORDER[2]), "Received $"),
    }

    months = df.dropna(subset=["Month No"]).groupby(["Month No", "Month"]).size().reset_index()
    months = months.sort_values("Month No")
    label = lambda m: f'{m["Month"][:3]} {int(m["Month No"]) // 100}'
    monthly, claims_monthly = [], []
    for _, m in months.iterrows():
        inm = df["Month No"] == m["Month No"]
        row = {"month": label(m), "total": int(W[inm].sum())}
        row.update({k: cnt(inm & (df["Main"] == k)) for k in MAIN_ORDER})
        row["potential"] = cnt(inm & df["Potential Loss"])
        row.update({s: cnt(inm & disp & (df["Sub"] == s)) for s in SUB_ALL})
        cl = row["Delivered"] + row["Under Dispute"]
        row["delivery_rate"] = round(row["Delivered"] / max(cl, 1) * 100, 2)
        row["loss_rate"] = round(row["Under Dispute"] / max(cl, 1) * 100, 2)
        row["received_amt"] = exp_sum(inm & disp & (df["Sub"] == SUB_ORDER[2]), "Received $")
        row["open_claim_value"] = exp_sum(inm & disp & df["Sub"].isin(SUB_ORDER[:2]))
        monthly.append(row)
        if row["Under Dispute"]:
            claims_monthly.append({"label": label(m), "cells": _claim_cells(df, W, inm & disp)})

    bands = [("16-30", 16, 30), ("31-45", 31, 45), ("46-60", 46, 60), ("61-90", 61, 90), ("91+", 91, 99999)]
    claims_ageing = [{"label": b, "cells": _claim_cells(df, W, disp & (df["Age"] >= lo) & (df["Age"] <= hi))}
                     for b, lo, hi in bands]

    carriers = []
    for c, d in df[df["Carrier"].notna()].groupby("Carrier"):
        if len(d) < 20:
            continue
        w = W[d.index]
        m = lambda cond: int(w[cond].sum())
        closed_c = m(d["Main"].isin(["Delivered", "Under Dispute"]))
        carriers.append({"carrier": c, "shipments": int(w.sum()), "delivered": m(d["Main"] == "Delivered"),
                         "in_transit": m(d["Main"] == "In Transit"), "potential": m(d["Potential Loss"]),
                         "dispute": m(d["Main"] == "Under Dispute"),
                         "loss_rate": round(m(d["Main"] == "Under Dispute") / max(closed_c, 1) * 100, 2)})
    carriers.sort(key=lambda r: -r["shipments"])

    # accurate ageing (days since Import Date) for the two claim lists still open
    age_bands = [("16-30", 16, 30), ("31-45", 31, 45), ("46-60", 46, 60), ("61-90", 61, 90), ("91+", 91, 99999)]
    sla_lo = S["claim_sla_days"][0]
    ageing_cards = {}
    for key, sub_name in (("pending", SUB_ORDER[0]), ("raised", SUB_ORDER[1])):
        m = disp & (df["Sub"] == sub_name)
        d = df[m]
        ages = d["Age"].fillna(0).astype(int)
        ageing_cards[key] = {
            "count": cnt(m), "orders": int(len(d)),
            "avg": round(float(ages.mean()), 1) if len(d) else 0, "median": round(float(ages.median()), 1) if len(d) else 0,
            "oldest": int(ages.max()) if len(d) else 0, "newest": int(ages.min()) if len(d) else 0,
            "value": exp_sum(m),
            "past_sla": cnt(m & (df["Claim Age"] >= sla_lo)),
            "bands": [{"label": b, "count": cnt(m & (df["Age"] >= lo) & (df["Age"] <= hi)),
                       "value": exp_sum(m & (df["Age"] >= lo) & (df["Age"] <= hi))} for b, lo, hi in age_bands],
            "oldest_orders": [{"id": int(r["Shipment ID"]), "order": str(r["Store Order ID"]), "age": int(r["Age"]),
                               "carrier": "" if pd.isna(r["Carrier"]) else str(r["Carrier"])}
                              for _, r in d.sort_values("Age", ascending=False).head(5).iterrows()],
        }

    # claims value by month (dollars): Received = Amt column, the rest = Excel's Expected Claim amount
    mnames = [label(m) for _, m in months.iterrows()]
    vrows = {"received": [], "raised": [], "pending": [], "expired": []}
    kept_months = []
    for (_, m), nm in zip(months.iterrows(), mnames):
        inm = (df["Month No"] == m["Month No"]) & disp
        if not inm.any():
            continue
        kept_months.append(nm)
        vrows["received"].append(round(float(df.loc[inm & (df["Sub"] == SUB_ORDER[2]), "Received $"].sum()), 2))
        vrows["raised"].append(round(float(df.loc[inm & (df["Sub"] == SUB_ORDER[1]), "Expected Claim $"].sum()), 2))
        vrows["pending"].append(round(float(df.loc[inm & (df["Sub"] == SUB_ORDER[0]), "Expected Claim $"].sum()), 2))
        vrows["expired"].append(round(float(df.loc[inm & (df["Sub"] == EXPIRED), "Expected Claim $"].sum()), 2))
    vtot = {k: round(sum(v), 2) for k, v in vrows.items()}
    eligible = [round(a + b + c, 2) for a, b, c in zip(vrows["received"], vrows["raised"], vrows["pending"])]
    et = round(vtot["received"] + vtot["raised"] + vtot["pending"], 2)
    miss = df["Expected Missing"]
    value_table = {"months": kept_months, "rows": vrows, "totals": vtot, "eligible": eligible, "eligible_total": et,
                   "ytd": {k: round(vtot[k] / et * 100, 1) if et else 0 for k in ("received", "raised", "pending")},
                   "missing": {"count": int(miss.sum()), "calc": round(float(df.loc[miss, "Expected Calc $"].sum()), 2)}}

    # today's actions (T = in-transit days, i.e. the day an order enters dispute)
    T = S["transit_days"]
    raised_m = disp & (df["Sub"] == SUB_ORDER[1])
    pending_m = disp & (df["Sub"] == SUB_ORDER[0])
    actions = {
        "past_sla": cnt(raised_m & (df["Claim Age"] >= sla_lo)),                 # claims to chase with ShipBob
        "new_dispute": cnt(pending_m & (df["Age"] <= T + 2)),                    # just crossed the dispute day: raise the claim
        "soon": cnt(df["Potential Loss"] & (df["Age"] >= T - 1)),                # 1-2 days from entering dispute
        "new_dispute_max_age": T + 2, "soon_min_age": T - 1, "sla_claim_days": sla_lo,
    }

    lists, totals = {}, {}
    for name, (f, _, _) in LIST_DEF.items():
        totals[name] = cnt(f(df))
        lists[name] = _rows(lf := list_frame(df, name, meta, unit, SHOW_LIMIT), list(lf.columns))
    return {
        "unit": unit, "freshness": freshness(meta),
        "meta": {"as_of": meta["as_of"].strftime("%d %b %Y, %H:%M"), "sources": meta["sources"],
                 "loaded_at": meta["loaded_at"].strftime("%d %b %Y, %H:%M"), "settings": S,
                 "missing_prices": meta["missing_prices"], "show_limit": SHOW_LIMIT},
        "kpis": kpis, "monthly": monthly, "claims_monthly": claims_monthly, "claims_ageing": claims_ageing,
        "ageing_cards": ageing_cards, "actions": actions, "value_table": value_table,
        "carriers": carriers, "checks": build_checks(df, meta), "lists": lists, "list_totals": totals,
        "sub_order": SUB_ORDER, "sub_all": SUB_ALL,
    }


def build_checks(df, meta):
    """Things in the Excel that look inconsistent, so you can fix them at the source. Counted in orders.
    (Orders you mark Delivered by hand after confirming with ShipBob support / the carrier portal are trusted,
    and a received claim closes the order, so neither is flagged.)"""
    ids = lambda m: [int(x) for x in df.loc[m, "Shipment ID"].head(12)]
    ms = df["Sub Bucket"].fillna("").astype(str).str.strip().str.lower()
    known = ms.isin(["pending to claim", "claim raised but not received", "claim raised and received",
                     "claim window expired"])
    unknown = (df["Main"] == "Under Dispute") & ~known
    miss = meta["missing_prices"]
    return [
        {"id": "unknown_label", "level": "warn", "count": int(unknown.sum()), "ids": ids(unknown),
         "title": "Under Dispute orders with no claim stage in Sub Bucket",
         "help": "Counted as Pending to claim. Set the Sub Bucket in the Excel."},
        {"id": "missing_expected", "level": "info", "count": int(df["Expected Missing"].sum()), "ids": ids(df["Expected Missing"]),
         "title": "Pending / Raised orders with no Expected Claim amount in the Excel",
         "help": "They are not counted in the expected $. By price x qty they would add about $%.0f." % df.loc[df["Expected Missing"], "Expected Calc $"].sum()},
        {"id": "missing_prices", "level": "info", "count": len(miss), "ids": [],
         "title": "Products without a price in the Product price tab",
         "help": "Add the price so the expected claim $ is complete: " + "; ".join(m["product"][:48] for m in miss[:9]) + ("..." if len(miss) > 9 else "")},
    ]


def _claim_cells(df, W, mask):
    out = {}
    for s in SUB_ALL:
        m = mask & (df["Sub"] == s)
        out[s] = {"count": int(W[m].sum()), "expected": round(float(df.loc[m, "Expected Claim $"].sum()), 2),
                  "received": round(float(df.loc[m, "Received $"].sum()), 2)}
    return out


if __name__ == "__main__":
    import time
    t = time.time()
    df, meta = build_frame(load_config())
    for unit in ("lines", "orders"):
        p = build_payload(df, meta, unit)
        print(f"\n[{unit}] loaded {len(df):,} shipments in {time.time() - t:.0f}s, as of {p['meta']['as_of']}")
        print(json.dumps({k: p["kpis"][k] for k in ("total", "delivered", "in_transit", "potential", "dispute",
                                                    "rto", "cancelled", "sub_counts")}, indent=1))
