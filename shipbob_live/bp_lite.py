"""
ShipBob D2C Live, lite blueprint for Render (no pandas): same page and API as bp.py, but the data comes from the
Google Sheet tab "ShipBob D2C Claims" (engine_lite.py) and it is view only: no Gmail drafts, no Excel folder.

app.py picks this one when pandas is not installed (or SHIPBOB_MODE=lite). All routes sit behind the login gate.
"""
import io
import json
import os
import threading
import time
import traceback

from flask import Blueprint, Response, jsonify, request, send_file
from openpyxl import Workbook

from . import engine_lite as engine

bp = Blueprint("shipbob_live", __name__, url_prefix="/shipbob")

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "templates", "shipbob_live.html")
WATCH_SECONDS = 120
START = int(time.time())

STATE = {"ships": None, "meta": None, "payloads": None, "error": None, "loading": False, "version": 0, "sig": None}
LOCK = threading.Lock()
_started = False


def build_id():
    return f"{int(os.path.getmtime(TEMPLATE))}-{START}"


def refresh():
    if not LOCK.acquire(blocking=False):          # a rebuild is already running
        return False
    try:
        STATE["loading"] = True
        cfg = engine.load_config()
        sig = engine.source_sig(cfg)
        ships, meta = engine.build(cfg)
        payloads = {u: json.dumps(engine.payload(ships, meta, u), default=str) for u in engine.UNITS}
        STATE.update(ships=ships, meta=meta, payloads=payloads, error=None, sig=sig, version=STATE["version"] + 1)
        return True
    except Exception as e:
        traceback.print_exc()
        STATE["error"] = str(e)
        return False
    finally:
        STATE["loading"] = False
        LOCK.release()


def _watch():
    refresh()
    while True:
        time.sleep(WATCH_SECONDS)
        try:
            sig = engine.source_sig(engine.load_config())
        except Exception:
            continue
        if sig != STATE["sig"]:
            print("[ShipBob Live] sheet changed, rebuilding...")
            refresh()


def start_watcher():
    global _started
    if _started:
        return False
    _started = True
    threading.Thread(target=_watch, daemon=True, name="shipbob-live-lite").start()
    print(f"[ShipBob Live] lite mode: reading the Google Sheet, checking every {WATCH_SECONDS}s (page at /shipbob/)")
    return True


@bp.before_request
def _lazy_start():          # gunicorn never runs app.py's __main__ block, so start on the first request
    start_watcher()


# ------------------------------------------------------------------ page + data
@bp.route("/")
def index():
    with open(TEMPLATE, encoding="utf-8") as f:
        html = f.read()
    resp = Response(html.replace("__BUILD__", build_id()).replace("__BASE__", request.script_root + "/shipbob"),
                    mimetype="text/html")
    resp.headers["X-Frame-Options"] = "SAMEORIGIN"
    resp.headers["Cache-Control"] = "no-store"
    return resp


@bp.route("/api/data")
def data():
    if STATE["payloads"] is None:
        return jsonify({"loading": True, "error": STATE["error"]}), 503
    unit = request.args.get("unit", "lines")
    return Response(STATE["payloads"][unit if unit in engine.UNITS else "lines"], mimetype="application/json")


@bp.route("/api/status")
def status():
    meta = STATE["meta"]
    return jsonify({
        "version": STATE["version"], "loading": STATE["loading"], "error": STATE["error"],
        "as_of": meta["as_of"].strftime("%d %b %Y, %H:%M") if meta else None,
        "mailer": {"readonly": True, "gmail_connected": False, "kinds": {}}, "last_auto": [], "auto_log": [],
        "watch_seconds": WATCH_SECONDS, "build": build_id(),
    })


@bp.route("/api/refresh", methods=["POST"])
def api_refresh():
    """Rebuild in the background (it takes ~30 s, longer than a web request may run); the page picks up the new version."""
    threading.Thread(target=refresh, daemon=True).start()
    return jsonify({"ok": True, "started": True})


@bp.route("/download/<name>")
def download(name):
    if STATE["ships"] is None:
        return "Data not loaded yet", 404
    if name not in engine.LIST_DEF:
        return "Unknown list", 404
    unit = request.args.get("unit", "lines")
    rows = engine.list_rows(STATE["ships"], name, unit if unit in engine.UNITS else "lines")
    wb = Workbook(write_only=True)
    ws = wb.create_sheet(name[:31])
    cols = list(rows[0].keys()) if rows else engine.BASE_COLS
    names = {"Age": "Days Since Import", "Sub": "Bucket"}
    ws.append([names.get(c, c) for c in cols])
    for r in rows:
        ws.append([r.get(c) for c in cols])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    stamp = STATE["meta"]["as_of"].strftime("%Y-%m-%d")
    return send_file(buf, as_attachment=True, download_name=f"ShipBob_D2C_{name}_{stamp}.xlsx")


@bp.route("/download/mailer/<kind>")
@bp.route("/api/mailer/<kind>/<act>", methods=["POST"])
@bp.route("/api/mailer/<kind>", methods=["POST"])
@bp.route("/api/mailer/connect-gmail", methods=["POST"])
@bp.route("/api/mailer/open-folder", methods=["POST"])
def mail_off(kind=None, act=None):
    return jsonify({"ok": False, "error": "Mail drafts are created from the office PC (localhost:5000), not here."}), 404
