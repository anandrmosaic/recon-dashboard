import json
import socket
import threading
from datetime import datetime
from flask import Flask, jsonify, render_template, request

# IPv6 is unavailable on this network; force IPv4 for all outbound connections
# (httplib2 / googleapiclient pick AAAA records first, which then time out)
_orig_getaddrinfo = socket.getaddrinfo
def _ipv4_first(host, port, family=0, type=0, proto=0, flags=0):
    results = _orig_getaddrinfo(host, port, family, type, proto, flags)
    return sorted(results, key=lambda r: r[0] != socket.AF_INET)
socket.getaddrinfo = _ipv4_first
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger
import pytz

from auth import get_sheets_credentials, get_gmail_credentials
from sheets_service import get_sheet_data
from email_service import send_weekly_report
from provision_engine import get_sheet_carriers, process_provision

app = Flask(__name__)

with open('config.json') as f:
    CONFIG = json.load(f)

_cache = {'data': None, 'last_updated': None}

CF = CONFIG.get('credentials_file')
TF = CONFIG.get('token_file')


def refresh_data():
    try:
        creds = get_sheets_credentials(CF, TF)
        data = get_sheet_data(creds, CONFIG['sheet_id'], CONFIG['sheet_tab'], CONFIG.get('recon_tab'))
        _cache['data'] = data
        _cache['last_updated'] = datetime.now().strftime('%d %b %Y, %I:%M %p IST')
        print(f"[Data] Refreshed at {_cache['last_updated']}")
    except Exception as e:
        print(f"[Data] Refresh error: {e}")


def send_scheduled_email():
    try:
        if not _cache['data']:
            refresh_data()
        creds = get_gmail_credentials(CF, TF)
        send_weekly_report(
            creds,
            _cache['data'],
            CONFIG['email_recipients'],
            CONFIG['email_from'],
            CONFIG['dashboard_url']
        )
    except Exception as e:
        print(f"[Email] Error: {e}")


@app.route('/')
def dashboard():
    return render_template('dashboard.html')


@app.route('/api/data')
def api_data():
    if not _cache['data']:
        refresh_data()
    return jsonify({
        'data': _cache['data'],
        'last_updated': _cache['last_updated']
    })


@app.route('/api/refresh')
def api_refresh():
    refresh_data()
    return jsonify({'status': 'ok', 'last_updated': _cache['last_updated']})


@app.route('/api/send-test-email')
def api_send_test_email():
    try:
        send_scheduled_email()
        return jsonify({'status': 'ok', 'message': f"Email sent to {CONFIG['email_recipients']}"})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/provision')
def provision():
    return render_template('provision.html')


@app.route('/api/provision/detect-carriers', methods=['GET', 'POST'])
def api_detect_carriers():
    try:
        month = int(request.args.get('month') or request.form.get('month', 0))
        year  = int(request.args.get('year')  or request.form.get('year',  0))
        if not month or not year:
            return jsonify({'status': 'ok', 'carriers': []})
        carriers = get_sheet_carriers(get_sheets_credentials(CF, TF), CONFIG['sheet_id'], month, year)
        return jsonify({'status': 'ok', 'carriers': carriers})
    except Exception as e:
        import traceback
        print(f"[Provision] detect-carriers error: {traceback.format_exc()}")
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/api/provision/generate', methods=['POST'])
def api_provision_generate():
    f = request.files.get('finance_file')
    if not f:
        return jsonify({'status': 'error', 'message': 'No Finance file uploaded'}), 400
    try:
        month = int(request.form.get('month', 1))
        year  = int(request.form.get('year', 2026))
        carrier_rates = json.loads(request.form.get('carrier_rates', '{}'))
        result = process_provision(
            f.read(), month, year,
            get_sheets_credentials(CF, TF), CONFIG['sheet_id'], carrier_rates
        )
        return jsonify({'status': 'ok', **result})
    except Exception as e:
        import traceback
        print(f"[Provision] ERROR: {traceback.format_exc()}")
        return jsonify({'status': 'error', 'message': str(e)}), 500


if __name__ == '__main__':
    # Initial data load
    print("[Startup] Loading data from Google Sheets...")
    threading.Thread(target=refresh_data, daemon=True).start()

    # Set up weekly scheduler
    sched = CONFIG['schedule']
    scheduler = BackgroundScheduler(timezone=pytz.timezone(sched['timezone']))
    scheduler.add_job(
        send_scheduled_email,
        CronTrigger(
            day_of_week=sched['day_of_week'],
            hour=sched['hour'],
            minute=sched['minute'],
            timezone=pytz.timezone(sched['timezone'])
        )
    )
    scheduler.start()
    print(f"[Scheduler] Weekly email set for every {sched['day_of_week'].upper()} {sched['hour']}:{sched['minute']:02d} {sched['timezone']}")

    print("[Server] Dashboard running at http://localhost:5000")
    print("[Server] Send test email at http://localhost:5000/api/send-test-email")
    app.run(debug=False, host='0.0.0.0', port=5000)
