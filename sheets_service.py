from googleapiclient.discovery import build
from collections import defaultdict
from datetime import datetime, date as date_type, timedelta

MONTH_ORDER = ['January','February','March','April','May','June',
               'July','August','September','October','November','December']


def safe_float(val):
    try:
        return float(str(val).replace(',', '').strip()) if val else 0.0
    except:
        return 0.0


def get_sheet_data(creds, sheet_id, awb_tab, recon_tab=None, data_since=None):
    service = build('sheets', 'v4', credentials=creds)

    # Read AWB tracker raw data
    awb_result = service.spreadsheets().values().get(
        spreadsheetId=sheet_id,
        range=f"'{awb_tab}'!A1:AQ8000"
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

    return parse_awb_data(awb_values, remarks, data_since=data_since)


def _parse_date_only(date_str):
    if not date_str or not str(date_str).strip():
        return None
    ds = str(date_str).strip()
    for fmt in ('%d/%m/%Y', '%m/%d/%Y', '%Y-%m-%d', '%d-%m-%Y', '%d %b %Y', '%d %B %Y'):
        try:
            return datetime.strptime(ds, fmt).date()
        except ValueError:
            continue
    return None


def _parse_aging(date_str):
    """Return (days_open, bucket_label) from a case raise date string."""
    d = _parse_date_only(date_str)
    if d is None:
        return None, None
    days = (date_type.today() - d).days
    if days < 0:
        return days, None
    if days <= 30:   return days, '0-30 days'
    elif days <= 60: return days, '31-60 days'
    elif days <= 90: return days, '61-90 days'
    else:            return days, '90+ days'


def _parse_resolution(raise_date_str, close_date_str, status):
    """Return (days_to_close, is_closed, is_rejected).
    Uses close_date - raise_date when both dates present.
    Falls back to status keywords for is_closed when close date is missing.
    is_rejected = True when status contains 'reject' (closed but claim denied).
    'No Discrepancy' statuses are excluded — close date present but no real resolution."""
    status_lower = status.lower() if status else ''
    is_rejected = 'reject' in status_lower
    # No discrepancy = case closed because no issue was found, not a real win or loss
    no_discrepancy = 'no' in status_lower and 'discrep' in status_lower
    is_closed_status = is_rejected or bool(any(
        k in status_lower for k in ['close', 'done', 'reimb', 'resolved', 'completed']
    ))
    raise_d = _parse_date_only(raise_date_str)
    close_d = _parse_date_only(close_date_str)
    if raise_d and close_d:
        return max(0, (close_d - raise_d).days), not no_discrepancy, is_rejected
    return None, is_closed_status and not no_discrepancy, is_rejected


def parse_awb_data(values, remarks=None, data_since=None):
    # Cutoff: skip rows before (since_year, since_month)
    since_year  = int(data_since['year'])  if data_since else None
    since_month = int(data_since['month']) if data_since else None
    # Find the header row (col 0 = "Month", col 1 = "Year")
    header_idx = None
    for i, row in enumerate(values):
        if row and str(row[0]).strip().lower() == 'month' and len(row) > 1 and str(row[1]).strip().lower() == 'year':
            header_idx = i
            break

    if header_idx is None:
        print("[Sheets] Header row not found in AWB tracker")
        return {'channel_data': {}, 'transporter_data': {}, 'remarks': remarks or {}, 'kpis': {}}

    # Detect key columns dynamically from header row (exact / keyword matching)
    header_row = values[header_idx]
    headers_lower = [str(c).strip().lower() for c in header_row]

    def col_exact(name, fallback):
        """Find column whose header is exactly `name` (case-insensitive)."""
        try:
            return headers_lower.index(name.lower())
        except ValueError:
            return fallback

    def col_contains(keyword, fallback):
        """Find first column whose header contains `keyword` (case-insensitive)."""
        for i, h in enumerate(headers_lower):
            if keyword in h:
                return i
        return fallback

    # Channel = exact "Channel" header (not "Channel Shipment Status" etc.)
    channel_col      = col_exact('channel', fallback=34)
    # Case Raise Date = first header containing "case raise"
    case_raise_col   = col_contains('case raise', fallback=None)
    case_close_col   = col_contains('case close', fallback=None)
    # Current Status (ship partner portal) — used to exclude Abandon / RTO from shipment count
    ship_status_col  = col_contains('ship partner portal', fallback=None)
    if ship_status_col is None:
        ship_status_col = col_contains('current status', fallback=None)
    # Pickup Date — used for week-on-week comparison
    pickup_date_col  = col_contains('pick up', fallback=None)

    EXCLUDE_STATUSES = {'abandon', 'rto'}

    print(f"[Sheets] Columns — channel:{channel_col}  case_raise:{case_raise_col}  case_close:{case_close_col}  ship_status:{ship_status_col}  pickup_date:{pickup_date_col}")

    # Weekly buckets: this_week = last 7 days, last_week = 7-14 days ago (by pickup date)
    today_d    = date_type.today()
    week_start = today_d - timedelta(days=7)
    prev_start = today_d - timedelta(days=14)
    weekly = {
        'this_week': {'shipments': 0, 'lost_stock': 0, 'expected': 0.0, 'actual': 0.0},
        'last_week': {'shipments': 0, 'lost_stock': 0, 'expected': 0.0, 'actual': 0.0},
    }

    # Aggregate raw rows
    # Key: (month_str, year_int)  →  channel  →  metrics
    ch_agg = defaultdict(lambda: defaultdict(lambda: {
        'qty_sent': 0.0, 'lost_stock': 0.0, 'expected_reimburs': 0.0, 'actual_reimbursed': 0.0, 'shipment_count': 0
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

        # Skip rows before the configured cutoff (archive filter)
        if since_year is not None:
            month_idx = MONTH_ORDER.index(month) + 1  # 1-based
            if (year < since_year) or (year == since_year and month_idx < since_month):
                continue

        qty_sent        = safe_float(row[15] if len(row) > 15 else 0)
        lost_stock      = safe_float(row[27] if len(row) > 27 else 0)
        expected        = safe_float(row[28] if len(row) > 28 else 0)
        actual          = safe_float(row[29] if len(row) > 29 else 0)
        channel         = str(row[channel_col]).strip() if len(row) > channel_col else ''
        transporter     = str(row[8]).strip()  if len(row) > 8  else ''
        platform_label  = str(row[5]).strip()  if len(row) > 5  else ''
        product_name    = str(row[11]).strip() if len(row) > 11 else ''
        invoice_no      = str(row[14]).strip() if len(row) > 14 else ''

        period_set.add((month, year))

        if channel:
            d = ch_agg[(month, year)][channel]
            d['qty_sent']          += qty_sent
            d['lost_stock']        += lost_stock
            d['expected_reimburs'] += expected
            d['actual_reimbursed'] += actual
            ship_status = (
                str(row[ship_status_col]).strip().lower()
                if ship_status_col is not None and len(row) > ship_status_col
                else ''
            )
            if ship_status not in EXCLUDE_STATUSES:
                d['shipment_count'] += 1

        # Weekly bucketing by pickup date
        if pickup_date_col is not None and len(row) > pickup_date_col:
            pickup_d = _parse_date_only(str(row[pickup_date_col]).strip())
            if pickup_d is not None:
                if week_start <= pickup_d < today_d:
                    bucket = weekly['this_week']
                elif prev_start <= pickup_d < week_start:
                    bucket = weekly['last_week']
                else:
                    bucket = None
                if bucket is not None:
                    if platform_label:
                        bucket['shipments'] += 1
                    bucket['lost_stock'] += lost_stock
                    bucket['expected']   += expected
                    bucket['actual']     += actual

        if transporter:
            t = tr_agg[(month, year)][transporter]
            t['qty_sent']   += qty_sent

        # Extract case raise date before the condition so it can be used as a trigger
        case_raise_raw = (
            str(row[case_raise_col]).strip()
            if case_raise_col is not None and len(row) > case_raise_col
            else ''
        )

        # Include any row where a case was raised (has raise date) or has lost stock —
        # covers Excess Receive / Inventory Relocated rows that have lost_stock=0
        if (lost_stock > 0 or case_raise_raw) and channel:
            case_close_raw = (
                str(row[case_close_col]).strip()
                if case_close_col is not None and len(row) > case_close_col
                else ''
            )
            reimb_status = str(row[30]).strip() if len(row) > 30 else ''
            days_open, aging_bucket = _parse_aging(case_raise_raw)
            days_to_close, is_closed, is_rejected = _parse_resolution(case_raise_raw, case_close_raw, reimb_status)
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
                'reimbursement_status': reimb_status,
                'remark':               str(row[32]).strip() if len(row) > 32 else '',
                'product_name':         product_name,
                'invoice_no':           invoice_no,
                'case_raise_date':      case_raise_raw,
                'case_close_date':      case_close_raw,
                'days_open':            days_open,
                'aging_bucket':         aging_bucket,
                'days_to_close':        days_to_close,
                'is_closed':            is_closed,
                'is_rejected':          is_rejected,
            })
            if transporter:
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
                'shipment_count':    ch_agg[p].get(ch, {}).get('shipment_count', 0),
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
        'weekly':           weekly,
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


def get_ups_claims_data(creds, sheet_id):
    """Read UPS Claim tab + AWB Master (UPS) tab and return structured claims data."""
    if not sheet_id:
        return {'summary': {}, 'claims': []}
    service = build('sheets', 'v4', credentials=creds)

    # Read AWB Master (UPS) - col A = AWB, col B = TRUE/FALSE, row 1 has counts in cols F-H
    master_result = service.spreadsheets().values().get(
        spreadsheetId=sheet_id,
        range="'AWB Master (ups)'!A1:H1000"
    ).execute()
    master_rows = master_result.get('values', [])

    total_awbs = 0
    claim_filed = 0
    not_filed = 0
    if master_rows:
        header = master_rows[0]
        # Counts are in row 1 cols F(5), G(6), H(7)
        if len(master_rows) > 1 and len(master_rows[1]) >= 8:
            try: claim_filed = int(str(master_rows[1][5]).replace(',','').strip())
            except: pass
            try: not_filed = int(str(master_rows[1][6]).replace(',','').strip())
            except: pass
            try: total_awbs = int(str(master_rows[1][7]).replace(',','').strip())
            except: pass

    # Read UPS Claim tab
    claim_result = service.spreadsheets().values().get(
        spreadsheetId=sheet_id,
        range="'UPS Claim'!A1:G500"
    ).execute()
    claim_rows = claim_result.get('values', [])

    claims = []
    if len(claim_rows) > 1:
        for row in claim_rows[1:]:
            if not row or not str(row[0]).strip():
                continue
            padded = row + [''] * max(0, 7 - len(row))
            parent_awb    = str(padded[0]).strip()
            lost_awb      = str(padded[1]).strip()
            lost_qty      = str(padded[2]).strip()
            claim_amount  = str(padded[3]).strip()
            form_received = str(padded[4]).strip()
            approved_date = str(padded[5]).strip()
            settled_date  = str(padded[6]).strip()

            # Determine state
            if claim_amount:
                state = 'amount_received'
            elif form_received:
                state = 'filed_pending'
            else:
                state = 'not_filed'

            claims.append({
                'parent_awb':    parent_awb,
                'lost_awb':      lost_awb,
                'lost_qty':      lost_qty,
                'claim_amount':  claim_amount,
                'form_received': form_received,
                'approved_date': approved_date,
                'settled_date':  settled_date,
                'state':         state,
            })

    amount_received_count = sum(1 for c in claims if c['state'] == 'amount_received')
    filed_pending_count   = sum(1 for c in claims if c['state'] == 'filed_pending')

    return {
        'summary': {
            'total_awbs':           total_awbs,
            'claim_filed':          claim_filed,
            'not_filed':            not_filed,
            'amount_received_count': amount_received_count,
            'filed_pending_count':  filed_pending_count,
        },
        'claims': claims,
    }


def get_outward_loss_data(creds, sheet_id):
    if not sheet_id:
        return {'headers': [], 'rows': []}
    from googleapiclient.discovery import build
    service = build('sheets', 'v4', credentials=creds)
    result = service.spreadsheets().values().get(
        spreadsheetId=sheet_id,
        range="Sheet1!A1:F5000"
    ).execute()
    values = result.get('values', [])
    if not values:
        return {'headers': [], 'rows': []}
    return {'headers': values[0], 'rows': values[1:]}
