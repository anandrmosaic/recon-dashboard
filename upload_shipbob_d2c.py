# upload_shipbob_d2c.py
# ─────────────────────
# Reads ShipBob Excel source file(s) from G:\My Drive\Shibob D2C Claim\
# Writes ALL rows with exactly 17 columns (A–Q) to the "ShipBob D2C Claims"
# tab in the recon Google Sheet, in the order below.
#
# Source resolution: PINNED_FILES (closed historical periods that are never
# re-exported, e.g. "April-May") are always included by exact name. Everything
# else is auto-detected — among the remaining .xlsx files directly in FOLDER,
# the single most-recently-modified one is treated as the current rolling
# file. This means a new weekly export (whatever it's named/dated) is always
# picked up automatically; no code change needed each week. Move a retired
# current-period file into the "Old File" subfolder once its replacement has
# been dropped in, so an older file sitting around doesn't confuse things if
# its mtime is ever touched (e.g. by a sync tool) after the fact.
#
# Run after dropping a new file in the folder:
#     python upload_shipbob_d2c.py
# Or double-click: "Sync to Dashboard.bat" in G:\My Drive\Shibob D2C Claim\

import glob
import json
import os
import sys
import openpyxl

# ── Config ────────────────────────────────────────────────────────────────────
FOLDER      = r'G:\My Drive\Shibob D2C Claim'
SHEET_ID    = '1N8qozEIZUg2FWYRqdO4UGHvcHiTtfoZr_BK_UojDSSA'
TAB_NAME    = 'ShipBob D2C Claims'
CONFIG_FILE = os.path.join(os.path.dirname(__file__), 'config.json')

# Closed historical periods — always included, never auto-replaced because
# no new export for these months will ever appear.
PINNED_FILES = [
    'Shipbob Order Data -- April-May 2026.xlsx',
]

# Exactly 17 columns, strictly in this order (A → Q)
COLUMNS = [
    'Shipment ID',
    'Import Date',
    'Month',
    'Days of order',
    'Fulfillment Cost',
    'Line Item Name',
    'Line Item Qty',
    'Total Item Quantity',
    'Delivery Status',
    'Delivery remark',
    'Claim status',
    'Amt',
    'Expected Claim amount',
    'Claims remark',
    'Sub Bucket',
    'Main Bucket',
    'Row Status',
]


# ── Resolve source files: pinned by name + auto-detected latest current file ──
def find_source_files(folder, pinned_names):
    pinned_paths = []
    missing = []
    for name in pinned_names:
        path = os.path.join(folder, name)
        if os.path.exists(path):
            pinned_paths.append(path)
        else:
            missing.append(name)
    if missing:
        raise FileNotFoundError(
            f'Pinned file(s) not found in {folder}: {missing}\n'
            f'Update PINNED_FILES in upload_shipbob_d2c.py if these were renamed/moved.'
        )

    pinned_basenames = {os.path.basename(p) for p in pinned_paths}
    all_xlsx = [
        f for f in glob.glob(os.path.join(folder, '*.xlsx'))
        if not os.path.basename(f).startswith('~$')
    ]
    candidates = [f for f in all_xlsx if os.path.basename(f) not in pinned_basenames]
    if not candidates:
        raise FileNotFoundError(
            f'No current-period ShipBob file found in {folder} (besides pinned files).'
        )

    latest = max(candidates, key=os.path.getmtime)
    others = [f for f in candidates if f != latest]
    if others:
        print(f'  (ignoring {len(others)} older file(s) in the folder: '
              f'{[os.path.basename(f) for f in others]} — using most recently edited only)')

    return pinned_paths + [latest]


# ── Read Excel → list of rows ─────────────────────────────────────────────────
def read_excel(path):
    print(f'Reading: {os.path.basename(path)}')
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)

    if 'Dump' not in wb.sheetnames:
        raise ValueError(f"Sheet 'Dump' not found. Available: {wb.sheetnames}")

    ws = wb['Dump']
    rows_iter = ws.iter_rows(min_row=1, values_only=True)

    # Build column index map from header row (case-insensitive match)
    raw_headers = [str(v).strip() if v is not None else '' for v in next(rows_iter)]
    col_idx = {}
    for col in COLUMNS:
        for i, h in enumerate(raw_headers):
            if h.lower() == col.lower():
                col_idx[col] = i
                break

    missing = [c for c in COLUMNS if c not in col_idx]
    if missing:
        print(f'  ⚠ Columns not found in Excel (will be blank): {missing}')

    # Excel formula error strings — treat as blank
    EXCEL_ERRORS = {'#VALUE!', '#REF!', '#N/A', '#DIV/0!', '#NAME?', '#NULL!', '#NUM!'}

    def cell_val(raw, col):
        if col not in col_idx:
            return ''
        i = col_idx[col]
        if i >= len(raw):
            return ''
        v = raw[i]
        if v is None:
            return ''
        if hasattr(v, 'strftime'):          # datetime / date cell → YYYY-MM-DD
            return v.strftime('%Y-%m-%d')
        if isinstance(v, float):
            if v == int(v):
                return str(int(v))          # whole number → no decimal
            if col == 'Days of order':
                return str(round(v, 1))     # days → 1 decimal place
            return str(round(v, 4))         # other floats → 4 dp
        s = str(v).strip()
        if s in EXCEL_ERRORS:               # formula error → blank
            return ''
        return s

    data_rows = []
    for raw in rows_iter:
        if not any(raw):
            continue
        data_rows.append([cell_val(raw, col) for col in COLUMNS])

    wb.close()
    print(f'  Rows read: {len(data_rows):,}')
    return data_rows


# ── Write to Google Sheets ────────────────────────────────────────────────────
def upload_to_sheets(data_rows):
    from auth import get_sheets_credentials
    from googleapiclient.discovery import build

    with open(CONFIG_FILE) as f:
        config = json.load(f)

    creds   = get_sheets_credentials(config.get('credentials_file'), config.get('token_file'))
    service = build('sheets', 'v4', credentials=creds)
    ss      = service.spreadsheets()

    # ── Ensure tab exists ─────────────────────────────────────────────────
    meta       = ss.get(spreadsheetId=SHEET_ID).execute()
    tab_names  = [s['properties']['title'] for s in meta['sheets']]

    if TAB_NAME not in tab_names:
        print(f'  Creating tab: "{TAB_NAME}"')
        ss.batchUpdate(
            spreadsheetId=SHEET_ID,
            body={'requests': [{'addSheet': {'properties': {'title': TAB_NAME}}}]}
        ).execute()
        # Re-fetch after creating
        meta = ss.get(spreadsheetId=SHEET_ID).execute()

    # ── Expand grid to fit all rows ───────────────────────────────────────
    sheet_id_int = next(
        s['properties']['sheetId']
        for s in meta['sheets']
        if s['properties']['title'] == TAB_NAME
    )
    needed_rows = len(data_rows) + 10      # header + data + small buffer
    needed_cols = len(COLUMNS) + 1
    ss.batchUpdate(
        spreadsheetId=SHEET_ID,
        body={'requests': [{'updateSheetProperties': {
            'properties': {
                'sheetId': sheet_id_int,
                'gridProperties': {
                    'rowCount': needed_rows,
                    'columnCount': needed_cols,
                },
            },
            'fields': 'gridProperties.rowCount,gridProperties.columnCount',
        }}]}
    ).execute()
    print(f'  Grid expanded to {needed_rows:,} rows × {needed_cols} cols')

    # ── Clear existing content (wipe full grid width to remove any old columns) ──
    ss.values().clear(
        spreadsheetId=SHEET_ID,
        range=f"'{TAB_NAME}'!A1:Z{needed_rows}",
    ).execute()

    # ── Write header + data in 5 000-row batches ──────────────────────────
    all_rows = [COLUMNS] + data_rows
    CHUNK    = 5_000
    for start in range(0, len(all_rows), CHUNK):
        chunk = all_rows[start:start + CHUNK]
        ss.values().update(
            spreadsheetId=SHEET_ID,
            range=f"'{TAB_NAME}'!A{start + 1}",
            valueInputOption='RAW',
            body={'values': chunk},
        ).execute()
        end = start + len(chunk)
        print(f'  → batch {start // CHUNK + 1}: rows {start + 1}–{end}')

    print(f'\n✅ Done — {len(data_rows):,} rows written to "{TAB_NAME}"')
    print('Go to the dashboard and hit Refresh.')


# ── Main ──────────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    try:
        paths = find_source_files(FOLDER, PINNED_FILES)

        data_rows = []
        for path in paths:
            data_rows.extend(read_excel(path))

        print(f'\nCombined rows from {len(paths)} file(s): {len(data_rows):,}')

        if not data_rows:
            print('No rows found in Excel — nothing uploaded.')
            sys.exit(0)

        upload_to_sheets(data_rows)

    except Exception as e:
        print(f'\n❌ Error: {e}')
        import traceback; traceback.print_exc()
        sys.exit(1)
