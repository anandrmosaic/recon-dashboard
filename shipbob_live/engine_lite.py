"""
ShipBob D2C engine, lite version: pure Python (no pandas / numpy), reads the Google Sheet tab
"ShipBob D2C Claims" in chunks and builds the same dashboard payload as engine.py.

It exists so the page can run on Render's small free instance (pandas broke production there: memory).
Rules are the same as engine.py (see its docstring); nothing here writes to the sheet.
"""
import json
import math
import os
import re
import statistics
import sys
from datetime import datetime

BASE = os.path.dirname(os.path.abspath(__file__))
PRICES_FILE = os.path.join(BASE, "product_prices.csv")
SHEET_TAB = "ShipBob D2C Claims"
SYNC_TAB = "ShipBob D2C Sync Info"
CHUNK = 5000

MAIN_ORDER = ["Delivered", "In Transit", "Under Dispute", "RTO", "Cancelled"]
POTENTIAL = "Potential In-Transit Loss"
EXPIRED = "Claim Window Expired"
SUB_ORDER = ["Pending to claim", "Claim raised but not received", "Claim raised and received"]
SUB_ALL = SUB_ORDER + [EXPIRED]
SHOW_LIMIT = 500
UNITS = ("lines", "orders")

DEFAULTS = {"transit_days": 15, "alert_from_day": 11, "claim_sla_days": [45, 60]}


# ------------------------------------------------------------------ config + sheet access (read-only)
def load_config():
    cfg = dict(DEFAULTS)
    path = os.path.join(BASE, "config.json")
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            cfg.update(json.load(f))
    if not cfg.get("sheet_id"):
        root = os.path.join(os.path.dirname(BASE), "config.json")
        with open(root, encoding="utf-8") as f:
            cfg["sheet_id"] = json.load(f)["recon_sheet_id"]
    return cfg


_SVC = {}


def _service(cfg):
    if "svc" not in _SVC:
        import sys
        app_dir = os.path.dirname(BASE)
        if app_dir not in sys.path:
            sys.path.insert(0, app_dir)
        from auth import get_sheets_credentials
        from googleapiclient.discovery import build
        creds = get_sheets_credentials(cfg.get("sheets_credentials_file", cfg.get("credentials_file")),
                                       cfg.get("sheets_token_file", "C:/Users/Admin/recon-token.json"))
        _SVC["svc"] = build("sheets", "v4", credentials=creds, cache_discovery=False)
    return _SVC["svc"]


def _values(cfg, rng, render="UNFORMATTED_VALUE"):
    r = _service(cfg).spreadsheets().values().get(spreadsheetId=cfg["sheet_id"], range=rng,
                                                  valueRenderOption=render).execute()
    return r.get("values", [])


def sheet_sync_info(cfg):
    v = _values(cfg, f"'{SYNC_TAB}'!A1:C1", "FORMATTED_VALUE")
    row = (v[0] if v else []) + ["", "", ""]
    return str(row[0]), str(row[1]), str(row[2])


def source_sig(cfg):
    return ("sheet",) + sheet_sync_info(cfg)


# ------------------------------------------------------------------ small helpers
def _norm(name):
    s = re.sub(r"\s+i\s+", " ", str(name).lower()).replace("keratine", "keratin")
    s = re.sub(r"[^a-z0-9]", "", s)
    return s[:-1] if s.endswith("s") else s


def _read_prices():
    out = {}
    if os.path.exists(PRICES_FILE):
        with open(PRICES_FILE, encoding="utf-8") as f:
            next(f, None)
            for line in f:
                name, _, price = line.rstrip("\r\n").rpartition(",")
                try:
                    out[_norm(name.strip().strip('"'))] = float(price)
                except ValueError:
                    pass
    return out


def _num(v):
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return None if (isinstance(v, float) and math.isnan(v)) else float(v)
    try:
        return float(str(v).strip().replace(",", ""))
    except ValueError:
        return None


def _blank(v):
    return v is None or v == "" or (isinstance(v, float) and math.isnan(v))


def _date(v):
    if _blank(v):
        return None
    if isinstance(v, (int, float)):        # serial date, in case the cell is a real date
        from datetime import timedelta
        return datetime(1899, 12, 30) + timedelta(days=float(v))
    s = str(v).strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%m/%d/%Y %H:%M:%S", "%m/%d/%Y"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


def _round(x, n=2):
    return round(x, n) + 0.0


def _low(v):
    return "" if _blank(v) else str(v).strip().lower()


class Ship:
    __slots__ = ("id", "store", "import_dt", "order_status", "carrier", "tracking", "delivery", "remark",
                 "sub_bucket", "main_bucket", "lines", "items", "amt", "exp", "calc", "price_missing", "litems",
                 "age", "month", "month_no", "main", "potential", "sub", "claim_age", "exp_claim", "exp_calc",
                 "exp_missing", "received", "stage")


# ------------------------------------------------------------------ load
def build(cfg):
    """Read the sheet in chunks and return (ships, meta). ships = list of Ship in sheet order of first appearance."""
    prices = _read_prices()
    head = _values(cfg, f"'{SHEET_TAB}'!A1:Z1")
    if not head:
        raise RuntimeError(f"Sheet tab '{SHEET_TAB}' is empty")
    hdr = [str(c).strip() for c in head[0]]
    ix = {h: i for i, h in enumerate(hdr)}
    need = ["Shipment ID", "Import Date", "Line Item Name", "Delivery Status", "Sub Bucket", "Main Bucket"]
    miss = [c for c in need if c not in ix]
    if miss:
        raise RuntimeError(f"Sheet tab is missing column(s): {miss}")
    ncol = len(hdr)
    g = lambda row, name: (row[ix[name]] if name in ix and ix[name] < len(row) else None)

    by_id, order = {}, []
    miss_products = {}                                   # product -> [line items, line item ids for dispute count]
    start = 2
    while True:
        rows = _values(cfg, f"'{SHEET_TAB}'!A{start}:Z{start + CHUNK - 1}")
        for row in rows:
            sid = _num(g(row, "Shipment ID"))
            if sid is None:
                continue
            sid = int(sid)
            s = by_id.get(sid)
            if s is None:
                s = by_id[sid] = Ship()
                order.append(s)
                s.id = sid
                s.store = s.import_dt = s.order_status = s.carrier = s.tracking = s.delivery = None
                s.remark = s.sub_bucket = s.main_bucket = None
                s.lines, s.items, s.amt, s.exp, s.calc, s.price_missing, s.litems = 0, [], None, None, None, False, []
            # pandas groupby.first(): first non-empty value per column
            for attr, name in (("store", "Store Order ID"), ("order_status", "Order Status"), ("carrier", "Carrier"),
                               ("tracking", "Tracking Information"), ("delivery", "Delivery Status"),
                               ("remark", "Claims remark"), ("sub_bucket", "Sub Bucket"),
                               ("main_bucket", "Main Bucket")):
                if getattr(s, attr) is None:
                    v = g(row, name)
                    if not _blank(v):
                        setattr(s, attr, v)
            if s.import_dt is None and not _blank(g(row, "Import Date")):
                s.import_dt = _date(g(row, "Import Date")) or False
            s.lines += 1
            name = g(row, "Line Item Name")
            qty = _num(g(row, "Line Item Qty"))
            amt = _num(g(row, "Amt"))
            exp = _num(g(row, "Expected Claim amount"))
            q = 1.0 if qty is None else qty
            item = ("nan" if _blank(name) else str(name)) + " x" + str(int(q))
            price = None if _blank(name) else prices.get(_norm(name))
            cogs = None if price is None else price * q
            s.items.append(item)
            if amt is not None:
                s.amt = (s.amt or 0.0) + amt
            if exp is not None:
                s.exp = (s.exp or 0.0) + exp
            if cogs is not None:
                s.calc = (s.calc or 0.0) + cogs
            if price is None:
                s.price_missing = True
                if not _blank(name):
                    miss_products.setdefault(str(name), []).append(s)
            s.litems.append((item, _round(cogs or 0.0), _round(exp or 0.0), _round(amt or 0.0)))
        if len(rows) < CHUNK:
            break
        start += CHUNK
        del rows

    ships = order
    T, A = cfg["transit_days"], cfg["alert_from_day"]
    dts = [s.import_dt for s in ships if s.import_dt]
    if not dts:
        raise RuntimeError("No valid Import Date in the sheet")
    as_of = max(dts)
    for s in ships:
        s.import_dt = s.import_dt or None
        s.items = " + ".join(s.items)
        ds = "" if _blank(s.delivery) else str(s.delivery).strip()
        s.stage = ds.split(" at ")[0] or "Blank (no scan yet)"
        if s.import_dt:
            s.age = int((as_of - s.import_dt).total_seconds() // 86400)
            s.month = s.import_dt.strftime("%B")
            s.month_no = s.import_dt.year * 100 + s.import_dt.month
        else:
            s.age, s.month, s.month_no = None, None, None
        a = s.age or 0
        mm, ms = _low(s.main_bucket), _low(s.sub_bucket)
        s.claim_age = max(a - T, 0)
        cancelled = s.order_status == "Cancelled" or mm == "cancelled" or ms == "cancelled"
        rto = mm == "rto" or ms == "rto"
        delivered = s.stage == "Delivered" or ms == "delivered"
        kept = ms in ("claim raised and received", "claim window expired")
        if cancelled:
            s.main = "Cancelled"
        elif rto:
            s.main = "RTO"
        elif kept:
            s.main = "Under Dispute"
        elif delivered:
            s.main = "Delivered"
        elif a <= T:
            s.main = "In Transit"
        else:
            s.main = "Under Dispute"
        s.potential = s.main == "In Transit" and a >= A
        if s.main == "Under Dispute":
            s.sub = {"claim raised and received": SUB_ORDER[2], "claim raised but not received": SUB_ORDER[1],
                     "claim window expired": EXPIRED}.get(ms, SUB_ORDER[0])
        else:
            s.sub = POTENTIAL if s.potential else s.main
        sheet_exp = s.exp or 0.0
        calc = _round(sheet_exp if (s.price_missing and sheet_exp > 0) else (s.calc or 0.0))
        s.exp_calc = calc
        open_stage = s.sub in (SUB_ORDER[0], SUB_ORDER[1], EXPIRED)
        s.exp_claim = _round(sheet_exp) if open_stage else calc
        s.exp_missing = s.sub in (SUB_ORDER[0], SUB_ORDER[1]) and sheet_exp <= 0
        s.received = _round(s.amt or 0.0)
        for attr in ("carrier", "order_status", "sub_bucket", "main_bucket"):     # few distinct values: share one copy
            v = getattr(s, attr)
            if isinstance(v, str):
                setattr(s, attr, sys.intern(v))
        if s.main in ("Delivered", "Cancelled"):          # not shown in any list: drop the heavy per-order text
            s.store = s.tracking = s.delivery = s.remark = None
            s.items, s.litems = "", ()

    missing_prices = []
    for name, lst in miss_products.items():
        missing_prices.append({"product": name, "line_items": len(lst),
                               "in_dispute": sum(1 for s in lst if s.main == "Under Dispute")})
    missing_prices.sort(key=lambda r: (-r["line_items"], r["product"]))

    at, fname, _n = sheet_sync_info(cfg)
    try:
        synced = datetime.strptime(at, "%Y-%m-%d %H:%M:%S")
    except ValueError:
        synced = datetime.now()
    name = f"Google Sheet ({fname})" if fname else f"Google Sheet - {SHEET_TAB}"
    meta = {"missing_prices": missing_prices, "as_of": as_of, "sources": [name], "data_source": "sheet",
            "files": [{"name": name, "modified": synced}], "latest_input": None, "loaded_at": datetime.now(),
            "settings": {"transit_days": T, "alert_from_day": A, "claim_sla_days": cfg.get("claim_sla_days", [45, 60])}}
    return ships, meta


# ------------------------------------------------------------------ lists
BASE_COLS = ["Shipment ID", "Store Order ID", "Import Date", "Age", "Carrier", "Tracking Information",
             "Delivery Status", "Sub"]
LIST_DEF = {  # list id: (filter, extra columns, (sort column, descending))
    "potential": (lambda s: s.potential, ["Mailed"], True),
    "pending": (lambda s: s.sub == "Pending to claim", ["Expected Claim $", "Mailed", "Claims remark"], True),
    "raised": (lambda s: s.sub == "Claim raised but not received",
               ["Claim Age", "Expected Claim $", "Mailed", "Claims remark"], True),
    "expired": (lambda s: s.sub == EXPIRED, ["Expected Claim $", "Claims remark"], False),
    "received": (lambda s: s.sub == "Claim raised and received",
                 ["Expected Claim $", "Received $", "Claims remark"], False),
    "in_transit": (lambda s: s.main == "In Transit", ["Mailed"], True),
    "rto": (lambda s: s.main == "RTO", ["Claims remark"], False),
}


def _sorted_by_age(sel, desc):
    have = [s for s in sel if s.age is not None]
    none = [s for s in sel if s.age is None]
    return sorted(have, key=lambda s: s.age, reverse=desc) + none      # blank age last, like pandas


def _val(s, col):
    return {"Shipment ID": lambda: s.id, "Store Order ID": lambda: s.store,
            "Import Date": lambda: s.import_dt.strftime("%Y-%m-%d") if s.import_dt else None,
            "Age": lambda: s.age, "Carrier": lambda: s.carrier, "Tracking Information": lambda: s.tracking,
            "Delivery Status": lambda: s.delivery, "Sub": lambda: s.sub, "Mailed": lambda: "",
            "Expected Claim $": lambda: s.exp_claim, "Claims remark": lambda: s.remark,
            "Claim Age": lambda: s.claim_age, "Received $": lambda: s.received}[col]()


def list_rows(ships, name, unit="lines", limit=None):
    f, extra, desc = LIST_DEF[name]
    sel = _sorted_by_age([s for s in ships if f(s)], desc)
    if limit:
        sel = sel[:limit]
    rows = []
    if unit != "lines":
        for s in sel:
            rows.append({c: _val(s, c) for c in BASE_COLS + extra})
    else:
        for s in sel:
            for item, l_exp, l_sheet, l_rec in s.litems:
                r = {c: _val(s, c) for c in BASE_COLS}
                r["Item"] = item
                for c in extra:
                    if c == "Expected Claim $":
                        r[c] = l_sheet if name in ("pending", "raised", "expired") else l_exp
                    elif c == "Received $":
                        r[c] = l_rec
                    else:
                        r[c] = _val(s, c)
                rows.append(r)
                if limit and len(rows) >= limit:
                    return rows
    return rows


# ------------------------------------------------------------------ payload
def freshness(meta):
    now = datetime.now()
    f = lambda t: t.strftime("%d %b %Y, %H:%M")
    newest = max(meta["files"], key=lambda x: x["modified"])
    data_age = (now.date() - meta["as_of"].date()).days
    return {
        "data_as_of": meta["as_of"].strftime("%d %b %Y"), "as_of_iso": meta["as_of"].strftime("%Y-%m-%d"),
        "data_age_days": data_age,
        "file_name": newest["name"], "file_modified": f(newest["modified"]),
        "file_age_days": (now.date() - newest["modified"].date()).days,
        "input_name": None, "input_modified": None, "input_age_days": None,
        "last_mailer": None, "dashboard_loaded": f(meta["loaded_at"]),
        "status": "ok" if data_age <= 1 else "warn" if data_age <= 3 else "stale",
    }


def payload(ships, meta, unit="lines"):
    S = meta["settings"]
    w = (lambda s: s.lines) if unit == "lines" else (lambda s: 1)
    cnt = lambda pred, rows=ships: int(sum(w(s) for s in rows if pred(s)))
    esum = lambda pred, attr="exp_claim", rows=ships: _round(math.fsum(getattr(s, attr) for s in rows if pred(s)))
    disp = lambda s: s.main == "Under Dispute"
    by = lambda m: cnt(lambda s: s.main == m)
    closed = by("Delivered") + by("Under Dispute")
    pend = [s for s in ships if disp(s) and s.sub == SUB_ORDER[0]]

    kpis = {
        "total": int(sum(w(s) for s in ships)), "delivered": by("Delivered"), "in_transit": by("In Transit"),
        "potential": cnt(lambda s: s.potential), "potential_mailed": 0,
        "pending_oldest": int(max((s.age for s in pend if s.age is not None), default=0) or 0),
        "dispute": by("Under Dispute"), "rto": by("RTO"), "cancelled": by("Cancelled"),
        "delivery_rate": round(by("Delivered") / max(closed, 1) * 100, 2),
        "loss_rate": round(by("Under Dispute") / max(closed, 1) * 100, 2),
        "sub_counts": {x: cnt(lambda s, x=x: disp(s) and s.sub == x) for x in SUB_ALL},
        "sub_value": {x: esum(lambda s, x=x: disp(s) and s.sub == x) for x in SUB_ALL},
        "received_amt": esum(lambda s: disp(s) and s.sub == SUB_ORDER[2], "received"),
    }

    month_keys = sorted({(s.month_no, s.month) for s in ships if s.month_no is not None})
    label = lambda mk: f"{mk[1][:3]} {mk[0] // 100}"
    in_month = {mk: [s for s in ships if s.month_no == mk[0]] for mk in month_keys}
    monthly, claims_monthly = [], []
    for mk in month_keys:
        rows = in_month[mk]
        row = {"month": label(mk), "total": int(sum(w(s) for s in rows))}
        row.update({k: cnt(lambda s, k=k: s.main == k, rows) for k in MAIN_ORDER})
        row["potential"] = cnt(lambda s: s.potential, rows)
        row.update({x: cnt(lambda s, x=x: disp(s) and s.sub == x, rows) for x in SUB_ALL})
        cl = row["Delivered"] + row["Under Dispute"]
        row["delivery_rate"] = round(row["Delivered"] / max(cl, 1) * 100, 2)
        row["loss_rate"] = round(row["Under Dispute"] / max(cl, 1) * 100, 2)
        row["received_amt"] = esum(lambda s: disp(s) and s.sub == SUB_ORDER[2], "received", rows)
        row["open_claim_value"] = esum(lambda s: disp(s) and s.sub in SUB_ORDER[:2], "exp_claim", rows)
        monthly.append(row)
        if row["Under Dispute"]:
            claims_monthly.append({"label": label(mk), "cells": _claim_cells(rows, w, disp)})

    bands = [("16-30", 16, 30), ("31-45", 31, 45), ("46-60", 46, 60), ("61-90", 61, 90), ("91+", 91, 99999)]
    inb = lambda s, lo, hi: s.age is not None and lo <= s.age <= hi
    claims_ageing = [{"label": b, "cells": _claim_cells(ships, w, lambda s, lo=lo, hi=hi: disp(s) and inb(s, lo, hi))}
                     for b, lo, hi in bands]

    carriers, groups = [], {}
    for s in ships:
        if s.carrier is not None:
            groups.setdefault(s.carrier, []).append(s)
    for c in sorted(groups):
        d = groups[c]
        if len(d) < 20:
            continue
        m = lambda pred: cnt(pred, d)
        closed_c = m(lambda s: s.main in ("Delivered", "Under Dispute"))
        carriers.append({"carrier": c, "shipments": int(sum(w(s) for s in d)), "delivered": m(lambda s: s.main == "Delivered"),
                         "in_transit": m(lambda s: s.main == "In Transit"), "potential": m(lambda s: s.potential),
                         "dispute": m(lambda s: s.main == "Under Dispute"),
                         "loss_rate": round(m(lambda s: s.main == "Under Dispute") / max(closed_c, 1) * 100, 2)})
    carriers.sort(key=lambda r: -r["shipments"])

    sla_lo = S["claim_sla_days"][0]
    ageing_cards = {}
    for key, sub_name in (("pending", SUB_ORDER[0]), ("raised", SUB_ORDER[1])):
        pm = lambda s, n=sub_name: disp(s) and s.sub == n
        d = [s for s in ships if pm(s)]
        ages = [s.age or 0 for s in d]
        top = sorted([s for s in d if s.age is not None], key=lambda s: s.age, reverse=True)[:5]
        ageing_cards[key] = {
            "count": cnt(pm), "orders": len(d),
            "avg": round(sum(ages) / len(ages), 1) if d else 0, "median": round(float(statistics.median(ages)), 1) if d else 0,
            "oldest": max(ages) if d else 0, "newest": min(ages) if d else 0,
            "value": esum(pm),
            "past_sla": cnt(lambda s: pm(s) and s.claim_age >= sla_lo),
            "bands": [{"label": b, "count": cnt(lambda s, lo=lo, hi=hi: pm(s) and inb(s, lo, hi)),
                       "value": esum(lambda s, lo=lo, hi=hi: pm(s) and inb(s, lo, hi))} for b, lo, hi in bands],
            "oldest_orders": [{"id": s.id, "order": str(s.store) if s.store is not None else "nan", "age": s.age,
                               "carrier": "" if s.carrier is None else str(s.carrier)} for s in top],
        }

    vrows = {"received": [], "raised": [], "pending": [], "expired": []}
    kept_months = []
    for mk in month_keys:
        rows = [s for s in in_month[mk] if disp(s)]
        if not rows:
            continue
        kept_months.append(label(mk))
        vrows["received"].append(esum(lambda s: s.sub == SUB_ORDER[2], "received", rows))
        vrows["raised"].append(esum(lambda s: s.sub == SUB_ORDER[1], "exp_claim", rows))
        vrows["pending"].append(esum(lambda s: s.sub == SUB_ORDER[0], "exp_claim", rows))
        vrows["expired"].append(esum(lambda s: s.sub == EXPIRED, "exp_claim", rows))
    vtot = {k: _round(sum(v)) for k, v in vrows.items()}
    eligible = [_round(a + b + c) for a, b, c in zip(vrows["received"], vrows["raised"], vrows["pending"])]
    et = _round(vtot["received"] + vtot["raised"] + vtot["pending"])
    miss = [s for s in ships if s.exp_missing]
    value_table = {"months": kept_months, "rows": vrows, "totals": vtot, "eligible": eligible, "eligible_total": et,
                   "ytd": {k: round(vtot[k] / et * 100, 1) if et else 0 for k in ("received", "raised", "pending")},
                   "missing": {"count": len(miss), "calc": _round(math.fsum(s.exp_calc for s in miss))}}

    T = S["transit_days"]
    actions = {
        "past_sla": cnt(lambda s: disp(s) and s.sub == SUB_ORDER[1] and s.claim_age >= sla_lo),
        "new_dispute": cnt(lambda s: disp(s) and s.sub == SUB_ORDER[0] and s.age is not None and s.age <= T + 2),
        "soon": cnt(lambda s: s.potential and s.age is not None and s.age >= T - 1),
        "new_dispute_max_age": T + 2, "soon_min_age": T - 1, "sla_claim_days": sla_lo,
    }

    lists, totals = {}, {}
    for name, (f, _, _) in LIST_DEF.items():
        totals[name] = cnt(f)
        lists[name] = list_rows(ships, name, unit, SHOW_LIMIT)
    return {
        "unit": unit, "freshness": freshness(meta),
        "meta": {"as_of": meta["as_of"].strftime("%d %b %Y, %H:%M"), "sources": meta["sources"],
                 "loaded_at": meta["loaded_at"].strftime("%d %b %Y, %H:%M"), "settings": S,
                 "missing_prices": meta["missing_prices"], "show_limit": SHOW_LIMIT},
        "kpis": kpis, "monthly": monthly, "claims_monthly": claims_monthly, "claims_ageing": claims_ageing,
        "ageing_cards": ageing_cards, "actions": actions, "value_table": value_table,
        "carriers": carriers, "checks": checks(ships, meta), "lists": lists, "list_totals": totals,
        "sub_order": SUB_ORDER, "sub_all": SUB_ALL,
    }


def checks(ships, meta):
    ids = lambda pred: [s.id for s in ships if pred(s)][:12]
    known = ("pending to claim", "claim raised but not received", "claim raised and received", "claim window expired")
    unknown = lambda s: s.main == "Under Dispute" and _low(s.sub_bucket) not in known
    miss = meta["missing_prices"]
    miss_exp = [s for s in ships if s.exp_missing]
    return [
        {"id": "unknown_label", "level": "warn", "count": sum(1 for s in ships if unknown(s)), "ids": ids(unknown),
         "title": "Under Dispute orders with no claim stage in Sub Bucket",
         "help": "Counted as Pending to claim. Set the Sub Bucket in the Excel."},
        {"id": "missing_expected", "level": "info", "count": len(miss_exp), "ids": ids(lambda s: s.exp_missing),
         "title": "Pending / Raised orders with no Expected Claim amount in the Excel",
         "help": "They are not counted in the expected $. By price x qty they would add about $%.0f." % math.fsum(s.exp_calc for s in miss_exp)},
        {"id": "missing_prices", "level": "info", "count": len(miss), "ids": [],
         "title": "Products without a price in the Product price tab",
         "help": "Add the price so the expected claim $ is complete: " + "; ".join(m["product"][:48] for m in miss[:9]) + ("..." if len(miss) > 9 else "")},
    ]


def _claim_cells(rows, w, mask):
    out = {}
    for x in SUB_ALL:
        sel = [s for s in rows if mask(s) and s.sub == x]
        out[x] = {"count": int(sum(w(s) for s in sel)), "expected": _round(math.fsum(s.exp_claim for s in sel)),
                  "received": _round(math.fsum(s.received for s in sel))}
    return out
