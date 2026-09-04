"""
sync_shipbob.py — Sync ShipBob claim rows to "ShipBob D2C Claims" Sheets tab

Reads from local Excel (G: drive), extracts claim rows only, writes a clean
18-column tab so the dashboard always has accurate numbers.

Run every Thursday (or set as daily scheduled task).

Usage:
  python sync_shipbob.py
"""
import json, os, sys
sys.path.insert(0, r'C:\Users\Admin\recon-dashboard')
from auth import get_sheets_credentials
from sheets_service import get_shipbob_d2c_from_excel
from googleapiclient.discovery import build

with open(r'C:\Users\Admin\recon-dashboard\config.json') as f:
    cfg = json.load(f)

FOLDER   = cfg.get('shipbob_d2c_folder', r'G:\My Drive\Shibob D2C Claim')
SHEET_ID = cfg['recon_sheet_id']
TAB_NAME = cfg.get('shipbob_d2c_tab', 'ShipBob D2C Claims')

# ── Read Excel via the same function the dashboard uses ───────────────────────
print(f"Reading from: {FOLDER}")
d2c = get_shipbob_d2c_from_excel(FOLDER)
if not d2c:
    print("ERROR: No Excel file found or could not be read.")
    sys.exit(1)

rows = d2c['rows']
kpis = d2c['kpis']
print(f"Claim rows loaded : {len(rows)}")
print(f"  rec  = Rs.{kpis['rec']['amt']}  ({kpis['rec']['count']} shipments)")
print(f"  prog = Rs.{kpis['prog']['exp']}  ({kpis['prog']['count']} shipments)")
print(f"  pend = Rs.{kpis['pend']['exp']}  ({kpis['pend']['count']} shipments)")

# ── Build clean 18-column output ──────────────────────────────────────────────
# Column order matches what get_shipbob_d2c_data() expects (header-name based)
HEADER = [
    'Month', 'Shipment ID', 'Sales Channel', 'SKU', 'Line Item Name',
    'Sub Bucket', 'Main Bucket', 'Claim status',
    'Amt', 'Expected Claim amount', 'Claims remark', 'Carrier',
]

def fmt(v):
    if v is None or v == '': return ''
    if isinstance(v, float) and v == int(v): return str(int(v))
    return str(v)

sheet_rows = [HEADER]
for r in rows:
    sheet_rows.append([
        fmt(r.get('month')),
        fmt(r.get('shipment_id')),
        fmt(r.get('channel')),
        fmt(r.get('sku')),
        fmt(r.get('line_name')),
        fmt(r.get('sub_bucket')),
        fmt(r.get('sub_bucket')),   # main bucket = sub bucket (already normalised)
        fmt(r.get('claim_status')),
        fmt(r.get('amt')),
        fmt(r.get('exp')),
        fmt(r.get('remark')),
        fmt(r.get('carrier')),
    ])

print(f"Rows to write     : {len(sheet_rows) - 1} (+ 1 header)")

# ── Connect to Sheets ─────────────────────────────────────────────────────────
creds   = get_sheets_credentials(cfg['credentials_file'], cfg['token_file'])
service = build('sheets', 'v4', credentials=creds)
ss      = service.spreadsheets()

# Create tab if missing; clear if exists
meta      = ss.get(spreadsheetId=SHEET_ID).execute()
tab_names = [s['properties']['title'] for s in meta['sheets']]

if TAB_NAME not in tab_names:
    print(f"Creating tab '{TAB_NAME}'...")
    ss.batchUpdate(
        spreadsheetId=SHEET_ID,
        body={'requests': [{'addSheet': {'properties': {'title': TAB_NAME}}}]}
    ).execute()
else:
    print(f"Clearing '{TAB_NAME}' tab...")
    ss.values().clear(
        spreadsheetId=SHEET_ID,
        range=f"'{TAB_NAME}'!A1:Z10000"
    ).execute()

# ── Write in one shot (< 300 rows, tiny) ─────────────────────────────────────
ss.values().update(
    spreadsheetId=SHEET_ID,
    range=f"'{TAB_NAME}'!A1",
    valueInputOption='RAW',
    body={'values': sheet_rows}
).execute()

print(f"\nSync complete — {len(sheet_rows)-1} rows in '{TAB_NAME}'")
print(f"Hit Refresh on the dashboard to see updated numbers.")
