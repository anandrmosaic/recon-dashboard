"""
ShipBob D2C Live: Flask blueprint mounted at /shipbob inside the recon dashboard (localhost:5000).

* Reads the order Excel from the G: drive folder (shipbob_live/config.json), classifies the orders and serves the
  dashboard page + its API.
* A watcher checks the Excel every 30 s; when it changes the data is rebuilt, the open page updates itself, and the two
  ShipBob mails are saved as Gmail DRAFTS (never sent).
* It only works on the computer that has the G: drive. Anywhere else (e.g. Render) it switches itself off and the page
  says so, so the rest of the recon dashboard is not affected.

All routes sit behind the recon dashboard's login (its before_request gate).
"""
import io
import json
import os
import threading
import time
import traceback

import pandas as pd
from flask import Blueprint, Response, jsonify, request, send_file

from . import engine, mailer

bp = Blueprint("shipbob_live", __name__, url_prefix="/shipbob")

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "templates", "shipbob_live.html")
UNITS = ("lines", "orders")     # lines = Excel rows (matches the Excel pivot), orders = one per shipment
WATCH_SECONDS = 30
START = int(time.time())

STATE = {"df": None, "meta": None, "payloads": None, "error": None, "loading": False,
         "version": 0, "sig": None, "last_auto": None, "auto_log": []}
LOCK = threading.Lock()
_watcher_started = False

OFF_HTML = ("<!doctype html><meta charset=utf-8><body style=\"font:15px system-ui;padding:48px;color:#475467\">"
            "<h2 style=\"margin:0 0 8px;color:#101828\">ShipBob D2C Live</h2>"
            "This view reads the order Excel from the G: drive and saves Gmail drafts, so it only runs on the computer that "
            "has the drive (localhost:5000). It is switched off here.</body>")


def enabled():
    try:
        return os.path.isdir(engine.load_config()["source_folder"])
    except Exception:
        return False


def build_id():
    """Changes when the page file is edited or the server restarts, so an open tab reloads itself."""
    return f"{int(os.path.getmtime(TEMPLATE))}-{START}"


# ------------------------------------------------------------------ data + automation
def source_sig():
    try:
        cfg = engine.load_config()
        if engine.uses_sheet(cfg):          # the upload script rewrites this tab's A1:C1 after every sync
            return ("sheet",) + engine.sheet_sync_info(cfg)
        return tuple((p, os.path.getmtime(p)) for p in engine.find_sources(cfg))
    except Exception:
        return None


def refresh():
    """Rebuild everything from the Excel. Returns True on success."""
    with LOCK:
        STATE["loading"] = True
        try:
            sig = source_sig()
            df, meta = engine.build_frame(engine.load_config())
            payloads = {u: json.dumps(engine.build_payload(df, meta, u), default=str) for u in UNITS}
            STATE.update(df=df, meta=meta, payloads=payloads, error=None, sig=sig, version=STATE["version"] + 1)
            return True
        except Exception as e:
            traceback.print_exc()
            STATE["error"] = str(e)
            return False
        finally:
            STATE["loading"] = False


def run_auto(force=False):
    """Build the two ShipBob mails and, if Gmail is connected, save the due ones as drafts."""
    stamp = time.strftime("%d %b %H:%M")
    try:
        results = mailer.auto_run(STATE["df"], STATE["meta"], force=force)
    except Exception as e:
        traceback.print_exc()
        results = [{"kind": "all", "action": "error", "reason": str(e)}]
    for r in results:
        r["at"] = stamp
    STATE["last_auto"] = results
    STATE["auto_log"] = ([r for r in results if r["action"] != "skipped"] + STATE["auto_log"])[:12]
    if any(r["action"] == "drafted" for r in results):
        refresh()      # so the Mailed column shows the new drafts
    return results


def _watch():
    if refresh():
        run_auto()
    while True:
        time.sleep(WATCH_SECONDS)
        sig = source_sig()
        if sig and sig != STATE["sig"]:
            print("[ShipBob Live] Excel changed, rebuilding...")
            time.sleep(5)            # let the file finish saving / syncing
            if refresh():
                run_auto()


def start_watcher():
    """Start the background watcher once (only where the G: drive exists)."""
    global _watcher_started
    if _watcher_started or not enabled():
        return False
    _watcher_started = True
    threading.Thread(target=_watch, daemon=True, name="shipbob-live-watcher").start()
    print(f"[ShipBob Live] watching the Excel every {WATCH_SECONDS}s (page at /shipbob/)")
    return True


def _off():
    return jsonify({"loading": False, "error": "ShipBob D2C Live is switched off on this computer"}), 503


# ------------------------------------------------------------------ page + data
@bp.route("/")
def index():
    if not enabled():
        return Response(OFF_HTML, mimetype="text/html")
    with open(TEMPLATE, encoding="utf-8") as f:       # read every time, so an edited page shows up without a restart
        html = f.read()
    resp = Response(html.replace("__BUILD__", build_id()).replace("__BASE__", request.script_root + "/shipbob"),
                    mimetype="text/html")
    resp.headers["X-Frame-Options"] = "SAMEORIGIN"      # the recon app sets DENY by default; this page lives in an iframe
    resp.headers["Cache-Control"] = "no-store"
    return resp


@bp.route("/api/data")
def data():
    if not enabled():
        return _off()
    if STATE["payloads"] is None:
        return jsonify({"loading": True, "error": STATE["error"]}), 503
    unit = request.args.get("unit", "lines")
    return Response(STATE["payloads"][unit if unit in UNITS else "lines"], mimetype="application/json")


@bp.route("/api/status")
def status():
    if not enabled():
        return _off()
    meta = STATE["meta"]
    return jsonify({
        "version": STATE["version"], "loading": STATE["loading"], "error": STATE["error"],
        "as_of": meta["as_of"].strftime("%d %b %Y, %H:%M") if meta else None,
        "mailer": mailer.status(), "last_auto": STATE["last_auto"], "auto_log": STATE["auto_log"],
        "watch_seconds": WATCH_SECONDS, "build": build_id(),
    })


@bp.route("/api/refresh", methods=["POST"])
def api_refresh():
    if not enabled():
        return _off()
    if not refresh():
        return jsonify({"ok": False, "error": STATE["error"]}), 500
    run_auto()
    return jsonify({"ok": True})


@bp.route("/download/<name>")
def download(name):
    if not enabled() or STATE["df"] is None:
        return "Data not loaded yet", 404
    if name not in engine.LIST_DEF:
        return "Unknown list", 404
    unit = request.args.get("unit", "lines")
    d = engine.list_frame(STATE["df"], name, STATE["meta"], unit if unit in UNITS else "lines")
    d = d.rename(columns={"Age": "Days Since Import", "Sub": "Bucket"})
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        d.to_excel(xw, index=False, sheet_name=name[:31])
    buf.seek(0)
    stamp = STATE["meta"]["as_of"].strftime("%Y-%m-%d")
    return send_file(buf, as_attachment=True, download_name=f"ShipBob_D2C_{name}_{stamp}.xlsx")


@bp.route("/download/mailer/<kind>")
def download_mail(kind):
    path = mailer.kstate(kind).get("excel_path") if enabled() and kind in mailer.KINDS else None
    if not path or not os.path.exists(path):
        return "No Excel built yet for this mail", 404
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))


# ------------------------------------------------------------------ mailer actions
def _fail(e):
    traceback.print_exc()
    return jsonify({"ok": False, "error": str(e)}), 500


def _kind(kind):
    if kind not in mailer.KINDS:
        raise ValueError(f"unknown mail '{kind}'")
    return kind


@bp.route("/api/mailer/<kind>/build", methods=["POST"])
def mailer_build(kind):
    if not enabled():
        return _off()
    try:
        r = mailer.build(_kind(kind), STATE["df"], STATE["meta"])
        return jsonify({"ok": True, "new": r["new"], "all": r["all"], "path": r["path"]})
    except Exception as e:
        return _fail(e)


@bp.route("/api/mailer/<kind>/draft", methods=["POST"])
def mailer_draft(kind):
    if not enabled():
        return _off()
    try:
        r = mailer.build(_kind(kind), STATE["df"], STATE["meta"])
        d = mailer.create_draft(r)
        refresh()
        return jsonify({"ok": True, "new": r["new"], "all": r["all"], **d})
    except Exception as e:
        return _fail(e)


@bp.route("/api/mailer/<kind>/mark-sent", methods=["POST"])
def mailer_mark_sent(kind):
    if not enabled():
        return _off()
    try:
        n = mailer.mark_sent(_kind(kind))
        refresh()
        return jsonify({"ok": True, "marked": n})
    except Exception as e:
        return _fail(e)


@bp.route("/api/mailer/connect-gmail", methods=["POST"])
def mailer_connect():
    if not enabled():
        return _off()
    try:
        mailer.connect_gmail()
        return jsonify({"ok": True})
    except Exception as e:
        return _fail(e)


@bp.route("/api/mailer/open-folder", methods=["POST"])
def open_folder():
    if not enabled():
        return _off()
    os.makedirs(mailer.OUT_DIR, exist_ok=True)
    os.startfile(mailer.OUT_DIR)
    return jsonify({"ok": True})
