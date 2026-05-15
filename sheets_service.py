from googleapiclient.discovery import build
from collections import defaultdict

MONTH_ORDER = ['January','February','March','April','May','June',
               'July','August','September','October','November','December']


def safe_float(val):
    try:
        return float(str(val).replace(',', '').strip()) if val else 0.0
    except:
        return 0.0


def get_sheet_data(creds, sheet_id, awb_tab, recon_tab=None):
    service = build('sheets', 'v4', credentials=creds)

    # Read AWB tracker raw data
    awb_result = service.spreadsheets().values().get(
        spreadsheetId=sheet_id,
        range=f"'{awb_tab}'!A1:AQ10000"
    ).execute()
    awb_values = awb_result.get('values', [])

    # Optionally read remarks from recon pivot tab
    remarks = {}
    if recon_tab:
        try:
            recon_result = service.spreadsheets().values().get(
                spreadsheetId=sheet_id,
                range=f"'{recon_tab}'!A1:R80"
            ).execute()
            recon_values = recon_result.get('values', [])
            remarks_start = next(
                (i for i, r in enumerate(recon_values) if r and 'Channel remark' in str(r[0])), None
            )
            if remarks_start is not None:
                remarks = parse_remarks(recon_values, remarks_start)
        except Exception as e:
            print(f"[Sheets] Could not read remarks: {e}")

    return parse_awb_data(awb_values, remarks)


def parse_awb_data(values, remarks=None):
    # Find the header row (col 0 = "Month", col 1 = "Year")
    header_idx = None
    for i, row in enumerate(values):
        if row and str(row[0]).strip().lower() == 'month' and len(row) > 1 and str(row[1]).strip().lower() == 'year':
            header_idx = i
            break

    if header_idx is None:
        print("[Sheets] Header row not found in AWB tracker")
        return {'channel_data': {}, 'transporter_data': {}, 'remarks': remarks or {}, 'kpis': {}}

    # Aggregate raw rows
    # Key: (month_str, year_int)  →  channel  →  metrics
    ch_agg = defaultdict(lambda: defaultdict(lambda: {
        'qty_sent': 0.0, 'lost_stock': 0.0, 'expected_reimburs': 0.0, 'actual_reimbursed': 0.0
    }))
    tr_agg = defaultdict(lambda: defaultdict(lambda: {'qty_sent': 0.0, 'lost_stock': 0.0}))

    period_set = set()  # (month_str, year_int)
    discrepancies = []  # individual shipment rows where lost_stock > 0

    for row in values[header_idx + 1:]:
        if not row or not str(row[0]).strip():
            continue
        month = str(row[0]).strip()
        if month.lower() in ['month', 'grand total']:
            continue
        if month not in MONTH_ORDER:
            continue

        year_raw = str(row[1]).strip() if len(row) > 1 else ''
        if not year_raw.isdigit():
            continue
        year = int(year_raw)

        qty_sent       = safe_float(row[15] if len(row) > 15 else 0)
        lost_stock     = safe_float(row[27] if len(row) > 27 else 0)
        expected       = safe_float(row[28] if len(row) > 28 else 0)
        actual         = safe_float(row[29] if len(row) > 29 else 0)
        channel        = str(row[33]).strip() if len(row) > 33 else ''
        transporter    = str(row[8]).strip()  if len(row) > 8  else ''

        period_set.add((month, year))

        if channel:
            d = ch_agg[(month, year)][channel]
            d['qty_sent']          += qty_sent
            d['lost_stock']        += lost_stock
            d['expected_reimburs'] += expected
            d['actual_reimbursed'] += actual

        if transporter:
            t = tr_agg[(month, year)][transporter]
            t['qty_sent']   += qty_sent

        if lost_stock > 0 and channel:
            discrepancies.append({
                'month':                f"{month} {year}",
                'awb':                  str(row[4]).strip()  if len(row) > 4  else '',
                'platform_label':       str(row[5]).strip()  if len(row) > 5  else '',
                'transporter':          transporter,
                'channel':              channel,
                'qty_sent':             int(qty_sent),
                'lost_stock':           int(lost_stock),
                'expected_reimburs':    round(expected, 2),
                'actual_reimbursed':    round(actual, 2),
                'pending':              round(expected - actual, 2),
                'reimbursement_status': str(row[30]).strip() if len(row) > 30 else '',
                'case_status':          str(row[39]).strip() if len(row) > 39 else '',
                'remark':               str(row[32]).strip() if len(row) > 32 else '',
            })
            t['lost_stock'] += lost_stock

    # Sort periods chronologically: year ASC, then calendar month order
    sorted_periods = sorted(period_set, key=lambda x: (x[1], MONTH_ORDER.index(x[0])))
    period_labels  = [f"{m} {y}" for m, y in sorted_periods]  # e.g. "April 2024"

    # Collect all unique channels and transporters
    all_channels     = sorted({ch for period_data in ch_agg.values() for ch in period_data})
    all_transporters = sorted({tr for period_data in tr_agg.values() for tr in period_data if tr})

    # Build per-channel monthly arrays
    channels = {}
    for ch in all_channels:
        channels[ch] = [
            {
                'qty_sent':          ch_agg[p].get(ch, {}).get('qty_sent', 0.0),
                'lost_stock':        ch_agg[p].get(ch, {}).get('lost_stock', 0.0),
                'expected_reimburs': ch_agg[p].get(ch, {}).get('expected_reimburs', 0.0),
                'actual_reimbursed': ch_agg[p].get(ch, {}).get('actual_reimbursed', 0.0),
            }
            for p in sorted_periods
        ]

    # Grand total per period
    grand_total = []
    for p in sorted_periods:
        row = {'qty_sent': 0.0, 'lost_stock': 0.0, 'expected_reimburs': 0.0, 'actual_reimbursed': 0.0}
        for ch_data in channels.values():
            r = ch_data[sorted_periods.index(p)]
            for k in row:
                row[k] += r[k]
        grand_total.append(row)

    totals = {k: sum(r[k] for r in grand_total)
              for k in ['qty_sent', 'lost_stock', 'expected_reimburs', 'actual_reimbursed']}

    # Build per-transporter monthly arrays
    transporters = {}
    for tr in all_transporters:
        transporters[tr] = [
            {
                'qty_sent':   tr_agg[p].get(tr, {}).get('qty_sent', 0.0),
                'lost_stock': tr_agg[p].get(tr, {}).get('lost_stock', 0.0),
            }
            for p in sorted_periods
        ]

    channel_data = {
        'months':      period_labels,
        'channels':    channels,
        'grand_total': grand_total,
        'totals':      totals,
    }
    transporter_data = {
        'months':       period_labels,
        'transporters': transporters,
        'totals':       {},
    }

    return {
        'channel_data':     channel_data,
        'transporter_data': transporter_data,
        'remarks':          remarks or {},
        'kpis':             calculate_kpis(channel_data),
        'discrepancies':    discrepancies,
    }


def parse_remarks(values, start):
    remarks = {}
    current_channel = current_month = None
    months = ['January', 'February', 'March', 'April', 'May', 'June',
              'July', 'August', 'September', 'October', 'November', 'December']

    for row in values[start + 1:]:
        if not row or not str(row[0]).strip():
            continue
        cell = str(row[0]).strip()
        if cell.lower() in ['tiktok', 'shipbob', 'amazon']:
            current_channel = 'TikTok' if cell.lower() == 'tiktok' else cell
            remarks.setdefault(current_channel, {})
        elif cell in months:
            current_month = cell
        elif current_channel and current_month and len(cell) > 5:
            remarks[current_channel][current_month] = cell

    return remarks


def calculate_kpis(channel_data):
    t = channel_data.get('totals', {})
    total_shipped = t.get('qty_sent', 0)
    total_lost    = t.get('lost_stock', 0)
    expected      = t.get('expected_reimburs', 0)
    actual        = t.get('actual_reimbursed', 0)
    pending       = expected - actual
    recovery_rate = round((actual / expected * 100), 1) if expected > 0 else 0
    loss_rate     = round((total_lost / total_shipped * 100), 3) if total_shipped > 0 else 0

    return {
        'total_shipped':    int(total_shipped),
        'total_lost':       int(total_lost),
        'expected_recovery': round(expected, 2),
        'actual_recovered': round(actual, 2),
        'pending_recovery': round(pending, 2),
        'recovery_rate':    recovery_rate,
        'loss_rate':        loss_rate,
    }
