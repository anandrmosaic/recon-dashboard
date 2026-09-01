# sync_shipbob_d2c.py
# ─────────────────────
# Reads claim rows from Source Google Sheet (Dump tab)
# Filters to Main Bucket = "Under Dispute" (~200 rows)
# Writes them into the "ShipBob D2C Claims" tab of the Recon Google Sheet.
#
# Run whenever you finish updating the source sheet:
#     python sync_shipbob_d2c.py
#
# Source  : https://docs.google.com/spreadsheets/d/1ykDp3piZ1MkkEVUakxVPSSaB2EU3ZhzG9sFSskBJWrE
# Dest    : https://docs.google.com/spreadsheets/d/1N8qozEIZUg2FWYRqdO4UGHvcHiTtfoZr_BK_UojDSSA

import json
import os
import sys

CONFIG_FILE = os.path.join(os.path.dirname(__file__), 'config.json')

SOURCE_ID  = '1ykDp3piZ1MkkEVUakxVPSSaB2EU3ZhzG9sFSskBJWrE'
SOURCE_TAB = 'Dump'
DEST_ID    = '1N8qozEIZUg2FWYRqdO4UGHvcHiTtfoZr_BK_UojDSSA'
DEST_TAB   = 'ShipBob D2C Claims'

# Columns to copy (no PII — name/email/address excluded)
KEEP = [
    'Month', 'Shipment ID', 'Store Order ID', 'Sales Channel', 'SKU',
    'Line Item Name', 'Sub Bucket', 'Main Bucket', 'Claim status', 'Amt',
    'Expected Claim amount', 'Claims remark', 'Carrier',
    'Delivery Status', 'Delivery remark', 'Row Status',
]


def main():
    with open(CONFIG_FILE) as f:
        config = json.load(f)

    from auth import get_sheets_credentials
    from googleapiclient.discovery import build

    creds   = get_sheets_credentials(config.get('credentials_file'), config.get('token_file'))
    service = build('sheets', 'v4', credentials=creds)

    # ── 1. Read source sheet ───────────────────────────────────────────────────
    print(f'Reading source: "{SOURCE_TAB}" tab...')
    result = service.spreadsheets().values().get(
        spreadsheetId=SOURCE_ID,
        range=f"'{SOURCE_TAB}'!A1:BZ100000",
        valueRenderOption='UNFORMATTED_VALUE',
        dateTimeRenderOption='FORMATTED_STRING',
    ).execute()

    all_rows = result.get('values', [])
    if not all_rows:
        print('❌ Source sheet is empty.')
        sys.exit(1)

    # ── 2. Map column names → indices ──────────────────────────────────────────
    raw_headers = [str(v).strip() if v else '' for v in all_rows[0]]

    def find_col(name):
        """Case-insensitive exact match."""
        nl = name.lower()
        for i, h in enumerate(raw_headers):
            if h.lower() == nl:
                return i
        return None

    col_idx = {col: find_col(col) for col in KEEP}
    missing  = [c for c, i in col_idx.items() if i is None]
    if missing:
        print(f'  Warning — columns not found (will be blank): {missing}')

    # Also need Main Bucket index for filtering
    mb_idx = find_col('Main Bucket')
    if mb_idx is None:
        print('❌ "Main Bucket" column not found in source. Cannot filter.')
        sys.exit(1)

    # ── 3. Filter to claim rows (Under Dispute) ────────────────────────────────
    claim_rows = []
    total_data = 0
    for raw in all_rows[1:]:
        if not any(raw):
            continue
        total_data += 1

        mb = str(raw[mb_idx]).strip().lower() if mb_idx < len(raw) else ''
        if mb != 'under dispute':
            continue

        row = []
        for col in KEEP:
            idx = col_idx.get(col)
            if idx is not None and idx < len(raw):
                row.append(str(raw[idx]) if raw[idx] is not None else '')
            else:
                row.append('')
        claim_rows.append(row)

    print(f'  Total rows scanned : {total_data:,}')
    print(f'  Claim rows found   : {len(claim_rows)} (Main Bucket = "Under Dispute")')

    if not claim_rows:
        print('No claim rows found — nothing written.')
        sys.exit(0)

    # ── 4. Write to destination sheet ─────────────────────────────────────────
    print(f'\nWriting to destination: "{DEST_TAB}" tab...')

    # Ensure tab exists
    meta = service.spreadsheets().get(spreadsheetId=DEST_ID).execute()
    existing_tabs = [s['properties']['title'] for s in meta['sheets']]
    if DEST_TAB not in existing_tabs:
        print(f'  Creating tab "{DEST_TAB}"...')
        service.spreadsheets().batchUpdate(
            spreadsheetId=DEST_ID,
            body={'requests': [{'addSheet': {'properties': {'title': DEST_TAB}}}]}
        ).execute()
    else:
        print(f'  Tab "{DEST_TAB}" exists — overwriting...')

    # Clear existing data
    service.spreadsheets().values().clear(
        spreadsheetId=DEST_ID,
        range=f"'{DEST_TAB}'!A1:ZZ100000"
    ).execute()

    # Write header + data in one call
    all_write = [KEEP] + claim_rows
    service.spreadsheets().values().update(
        spreadsheetId=DEST_ID,
        range=f"'{DEST_TAB}'!A1",
        valueInputOption='RAW',
        body={'values': all_write}
    ).execute()

    print(f'  ✅ Uploaded {len(claim_rows)} rows + header to "{DEST_TAB}"')
    print('\nDone! Go to the dashboard and hit Refresh.')


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print(f'\n❌ Error: {e}')
        import traceback; traceback.print_exc()
        sys.exit(1)
