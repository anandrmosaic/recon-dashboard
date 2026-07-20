import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
from googleapiclient.discovery import build


# ── helpers ────────────────────────────────────────────────────────────────

def _usd(n):
    return f"${n:,.2f}"

def _num(n):
    return f"{int(n):,}"

def _badge(text, bg, color):
    return f'<span style="background:{bg};color:{color};padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600">{text}</span>'

def _arrow(curr, prev, higher_is_bad=True):
    if prev == 0:
        return ''
    pct = ((curr - prev) / abs(prev)) * 100
    up = pct > 0
    symbol = '↑' if up else '↓'
    bad = (up and higher_is_bad) or (not up and not higher_is_bad)
    color = '#c62828' if bad else '#2e7d32'
    return f'<span style="font-size:11px;color:{color};margin-left:6px">{symbol} {abs(pct):.1f}%</span>'

def _section_title(text):
    return f'<p style="margin:0 0 10px;font-size:11px;color:#888;text-transform:uppercase;letter-spacing:0.6px;font-weight:700">{text}</p>'

def _th(text, align='center'):
    return f'<th style="padding:8px 10px;text-align:{align};font-weight:500;white-space:nowrap">{text}</th>'

def _td(text, align='center', style=''):
    return f'<td style="padding:7px 10px;border-bottom:1px solid #f0f0f0;text-align:{align};{style}">{text}</td>'


def send_weekly_report(creds, data, recipients, sender, dashboard_url):
    service = build('gmail', 'v1', credentials=creds)
    subject = f"Weekly Inventory Reconciliation Report — {datetime.now().strftime('%d %b %Y')}"
    html = build_enhanced_email_html(data, dashboard_url)

    for recipient in recipients:
        msg = MIMEMultipart('alternative')
        msg['Subject'] = subject
        msg['From'] = sender
        msg['To'] = recipient
        msg.attach(MIMEText(html, 'html'))
        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
        service.users().messages().send(userId='me', body={'raw': raw}).execute()

    print(f"[Email] Sent to {len(recipients)} recipients at {datetime.now()}")


def build_email_html(kpis, data, dashboard_url):
    months = data.get('channel_data', {}).get('months', [])
    channels = data.get('channel_data', {}).get('channels', {})
    totals = data.get('channel_data', {}).get('totals', {})

    channel_rows = ''
    for ch, rows in channels.items():
        total_lost = sum(r['lost_stock'] for r in rows)
        total_expected = sum(r['expected_reimburs'] for r in rows)
        total_actual = sum(r['actual_reimbursed'] for r in rows)
        channel_rows += f"""
        <tr>
            <td style="padding:8px;border:1px solid #ddd;font-weight:500">{ch}</td>
            <td style="padding:8px;border:1px solid #ddd;text-align:center;color:#c62828">{int(total_lost)}</td>
            <td style="padding:8px;border:1px solid #ddd;text-align:center">${total_expected:,.2f}</td>
            <td style="padding:8px;border:1px solid #ddd;text-align:center;color:#2e7d32">${total_actual:,.2f}</td>
            <td style="padding:8px;border:1px solid #ddd;text-align:center;color:#e65100">${(total_expected - total_actual):,.2f}</td>
        </tr>"""

    remarks_html = ''
    for ch, months_data in data.get('remarks', {}).items():
        for month, remark in months_data.items():
            remarks_html += f'<li><strong>{ch} – {month}:</strong> {remark}</li>'

    return f"""
<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;background:#f5f5f5;margin:0;padding:20px">
  <div style="max-width:700px;margin:0 auto;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.1)">

    <!-- Header -->
    <div style="background:#1a237e;padding:24px;text-align:center">
      <h1 style="color:#fff;margin:0;font-size:20px">India → USA Inventory Reconciliation</h1>
      <p style="color:#90caf9;margin:6px 0 0">{datetime.now().strftime('%A, %d %B %Y')}</p>
    </div>

    <!-- KPIs -->
    <div style="display:grid;grid-template-columns:repeat(2,1fr);gap:0;padding:0">
      <div style="background:#e3f2fd;padding:20px;text-align:center;border:1px solid #e0e0e0">
        <div style="font-size:28px;font-weight:700;color:#1565c0">{kpis.get('total_shipped', 0):,}</div>
        <div style="color:#666;font-size:13px;margin-top:4px">Total Units Shipped</div>
      </div>
      <div style="background:#ffebee;padding:20px;text-align:center;border:1px solid #e0e0e0">
        <div style="font-size:28px;font-weight:700;color:#c62828">{kpis.get('total_lost', 0):,}</div>
        <div style="color:#666;font-size:13px;margin-top:4px">Total Units Lost ({kpis.get('loss_rate', 0)}%)</div>
      </div>
      <div style="background:#fff8e1;padding:20px;text-align:center;border:1px solid #e0e0e0">
        <div style="font-size:28px;font-weight:700;color:#e65100">${kpis.get('pending_recovery', 0):,.2f}</div>
        <div style="color:#666;font-size:13px;margin-top:4px">Pending Recovery</div>
      </div>
      <div style="background:#e8f5e9;padding:20px;text-align:center;border:1px solid #e0e0e0">
        <div style="font-size:28px;font-weight:700;color:#2e7d32">${kpis.get('actual_recovered', 0):,.2f}</div>
        <div style="color:#666;font-size:13px;margin-top:4px">Recovered ({kpis.get('recovery_rate', 0)}%)</div>
      </div>
    </div>

    <!-- Channel Table -->
    <div style="padding:24px">
      <h2 style="color:#1a237e;font-size:16px;margin:0 0 12px">Channel Breakdown</h2>
      <table style="width:100%;border-collapse:collapse;font-size:13px">
        <thead>
          <tr style="background:#1a237e;color:#fff">
            <th style="padding:10px;text-align:left">Channel</th>
            <th style="padding:10px;text-align:center">Lost Units</th>
            <th style="padding:10px;text-align:center">Expected ($)</th>
            <th style="padding:10px;text-align:center">Recovered ($)</th>
            <th style="padding:10px;text-align:center">Pending ($)</th>
          </tr>
        </thead>
        <tbody>{channel_rows}</tbody>
      </table>
    </div>

    <!-- Remarks -->
    {"<div style='padding:0 24px 24px'><h2 style='color:#1a237e;font-size:16px;margin:0 0 12px'>Key Remarks</h2><ul style='color:#444;font-size:13px;line-height:1.8'>" + remarks_html + "</ul></div>" if remarks_html else ""}

    <!-- CTA -->
    <div style="background:#f5f5f5;padding:20px;text-align:center;border-top:1px solid #e0e0e0">
      <a href="{dashboard_url}" style="background:#1a237e;color:#fff;padding:12px 28px;border-radius:4px;text-decoration:none;font-size:14px;font-weight:600">
        View Live Dashboard
      </a>
      <p style="color:#999;font-size:11px;margin:12px 0 0">Auto-generated every Monday 12:01 PM IST</p>
    </div>
  </div>
</body>
</html>"""


# ── Enhanced email (preview + new mailer) ─────────────────────────────────

def build_enhanced_email_html(data, dashboard_url):
    # Filter channel data to 2026 only for the email
    cd_full  = data.get('channel_data', {})
    all_months = cd_full.get('months', [])
    idx_2026 = [i for i, m in enumerate(all_months) if '2026' in m]

    cd = {}
    if idx_2026:
        cd['months']      = [all_months[i] for i in idx_2026]
        cd['grand_total'] = [cd_full.get('grand_total', [])[i] for i in idx_2026]
        cd['channels']    = {
            ch: [rows[i] for i in idx_2026]
            for ch, rows in cd_full.get('channels', {}).items()
        }
        # Build recon disc lookup by month and by (month, channel)
        disc_by_month = {}
        disc_by_mc    = {}
        for d in (data.get('discrepancies') or []):
            m  = d.get('month', '')
            ch = d.get('channel', '')
            if not m: continue
            if m not in disc_by_month:
                disc_by_month[m] = {'lost_stock': 0.0, 'expected_reimburs': 0.0, 'actual_reimbursed': 0.0}
            disc_by_month[m]['lost_stock']        += d.get('lost_stock', 0) or 0
            disc_by_month[m]['expected_reimburs'] += d.get('expected_reimburs', 0) or 0
            disc_by_month[m]['actual_reimbursed'] += d.get('actual_reimbursed', 0) or 0
            if ch:
                k2 = (m, ch)
                if k2 not in disc_by_mc:
                    disc_by_mc[k2] = {'lost_stock': 0.0, 'expected_reimburs': 0.0, 'actual_reimbursed': 0.0}
                disc_by_mc[k2]['lost_stock']        += d.get('lost_stock', 0) or 0
                disc_by_mc[k2]['expected_reimburs'] += d.get('expected_reimburs', 0) or 0
                disc_by_mc[k2]['actual_reimbursed'] += d.get('actual_reimbursed', 0) or 0

        # Override grand_total per month with recon data
        for i, month in enumerate(cd['months']):
            if month in disc_by_month and i < len(cd['grand_total']):
                cd['grand_total'][i] = {**cd['grand_total'][i], **disc_by_month[month]}

        # Override per-channel monthly data with recon data
        for ch, rows in cd['channels'].items():
            for i, month in enumerate(cd['months']):
                k2 = (month, ch)
                if k2 in disc_by_mc and i < len(rows):
                    rows[i] = {**rows[i], **disc_by_mc[k2]}

        # Recompute totals from updated grand_total
        totals_2026 = {'qty_sent': 0.0, 'lost_stock': 0.0, 'expected_reimburs': 0.0, 'actual_reimbursed': 0.0}
        for row in cd['grand_total']:
            for k in totals_2026:
                totals_2026[k] += row.get(k, 0)
        # Override totals with recon sheet direct sums
        recon_totals = data.get('channel_data', {}).get('totals', {})
        if recon_totals.get('actual_reimbursed'):
            totals_2026['actual_reimbursed'] = recon_totals['actual_reimbursed']
        if recon_totals.get('expected_reimburs'):
            totals_2026['expected_reimburs']  = recon_totals['expected_reimburs']
        if recon_totals.get('lost_stock'):
            totals_2026['lost_stock']         = recon_totals['lost_stock']
        cd['totals'] = totals_2026
    else:
        cd = cd_full

    from sheets_service import calculate_kpis
    kpis     = calculate_kpis(cd) if idx_2026 else data.get('kpis', {})
    td       = data.get('transporter_data', {})
    months   = cd.get('months', [])
    channels = cd.get('channels', {})
    grand    = cd.get('grand_total', [])
    transporters = td.get('transporters', {})

    today = datetime.now().strftime('%A, %d %B %Y')

    # ── 1. Month-over-month ────────────────────────────────────────────────
    # Skip only the current in-progress month (last entry), compare previous two complete months
    mom_html = ''
    if len(grand) >= 3:
        cur  = grand[-2]
        prev = grand[-3]
        cur_label  = months[-2] if len(months) >= 2 else ''
        prev_label = months[-3] if len(months) >= 3 else ''
    elif len(grand) == 2:
        cur  = grand[-1]
        prev = grand[-2]
        cur_label  = months[-1] if months else ''
        prev_label = months[-2] if len(months) >= 2 else ''

    if len(grand) >= 2:

        def mom_cell(label, cur_val, prev_val, higher_is_bad=True, fmt='num'):
            diff = cur_val - prev_val
            pct  = ((diff) / abs(prev_val) * 100) if prev_val else 0
            up   = diff > 0
            bad  = (up and higher_is_bad) or (not up and not higher_is_bad)
            color = '#c62828' if bad else '#2e7d32'
            symbol = '↑' if up else '↓'
            val_str = _usd(cur_val) if fmt == 'usd' else _num(cur_val)
            pct_str = f'{symbol} {abs(pct):.1f}% vs {prev_label}'
            return f'''
            <div style="border:1px solid #e0e0e0;border-radius:6px;padding:14px;text-align:center;flex:1;min-width:140px">
              <div style="font-size:18px;font-weight:700;color:{color}">{val_str}</div>
              <div style="font-size:11px;color:#888;margin-top:3px">{label}</div>
              <div style="font-size:10px;color:{color};margin-top:3px">{pct_str}</div>
            </div>'''

        mom_html = f'''
        <div style="padding:20px 24px 0">
          {_section_title(f'▲▼ This Month ({cur_label}) vs Last Month')}
          <div style="display:flex;gap:8px;flex-wrap:wrap">
            {mom_cell('Lost Units',        cur.get('lost_stock',0),        prev.get('lost_stock',0),        higher_is_bad=True,  fmt='num')}
            {mom_cell('Pending Recovery',  cur.get('expected_reimburs',0) - cur.get('actual_reimbursed',0),
                                           prev.get('expected_reimburs',0) - prev.get('actual_reimbursed',0), higher_is_bad=True, fmt='usd')}
            {mom_cell('Recovered',         cur.get('actual_reimbursed',0), prev.get('actual_reimbursed',0), higher_is_bad=False, fmt='usd')}
          </div>
        </div>'''

    # ── 1b. Week-on-week ──────────────────────────────────────────────────
    wow_html = ''
    weekly = data.get('weekly', {})
    tw = weekly.get('this_week', {})
    lw = weekly.get('last_week', {})
    tw_ship = tw.get('shipments', 0)
    lw_ship = lw.get('shipments', 0)
    tw_lost = tw.get('lost_stock', 0)
    lw_lost = lw.get('lost_stock', 0)
    tw_pend = tw.get('expected', 0.0) - tw.get('actual', 0.0)
    lw_pend = lw.get('expected', 0.0) - lw.get('actual', 0.0)
    tw_rec  = tw.get('actual', 0.0)
    lw_rec  = lw.get('actual', 0.0)

    def wow_cell(label, cur_val, prev_val, higher_is_bad=True, fmt='num'):
        diff = cur_val - prev_val
        pct  = (diff / abs(prev_val) * 100) if prev_val else 0
        up   = diff > 0
        bad  = (up and higher_is_bad) or (not up and not higher_is_bad)
        color  = '#c62828' if bad else '#2e7d32'
        symbol = '↑' if up else ('↓' if diff < 0 else '→')
        val_str = _usd(cur_val) if fmt == 'usd' else _num(cur_val)
        pct_str = f'{symbol} {abs(pct):.1f}% vs last week' if prev_val else 'No prior data'
        return f'''
        <div style="border:1px solid #e0e0e0;border-radius:6px;padding:14px;text-align:center;flex:1;min-width:130px">
          <div style="font-size:18px;font-weight:700;color:{color}">{val_str}</div>
          <div style="font-size:11px;color:#888;margin-top:3px">{label}</div>
          <div style="font-size:10px;color:{color};margin-top:3px">{pct_str}</div>
        </div>'''

    from datetime import date as _date, timedelta as _tdelta
    _today = _date.today()
    _ws    = (_today - _tdelta(days=7)).strftime('%d %b')
    _we    = (_today - _tdelta(days=1)).strftime('%d %b')
    _ps    = (_today - _tdelta(days=14)).strftime('%d %b')
    _pe    = (_today - _tdelta(days=8)).strftime('%d %b')

    wow_html = f'''
    <div style="padding:20px 24px 0">
      {_section_title(f'⚡ This Week ({_ws}–{_we}) vs Last Week ({_ps}–{_pe})')}
      <div style="display:flex;gap:8px;flex-wrap:wrap">
        {wow_cell('Shipments Picked Up', tw_ship, lw_ship, higher_is_bad=False)}
        {wow_cell('Lost Units',          tw_lost, lw_lost, higher_is_bad=True)}
        {wow_cell('Pending Recovery',    tw_pend, lw_pend, higher_is_bad=True,  fmt='usd')}
        {wow_cell('Recovered',           tw_rec,  lw_rec,  higher_is_bad=False, fmt='usd')}
      </div>
    </div>'''

    # ── 2. Last 6 months trend ─────────────────────────────────────────────
    trend_rows = ''
    last6_months = months[-6:] if len(months) >= 6 else months
    last6_grand  = grand[-6:]  if len(grand)  >= 6 else grand
    for i, (m, g) in enumerate(zip(last6_months, last6_grand)):
        sent     = g.get('qty_sent', 0)
        lost     = g.get('lost_stock', 0)
        expected = g.get('expected_reimburs', 0)
        actual   = g.get('actual_reimbursed', 0)
        pending  = expected - actual
        loss_pct = round(lost / sent * 100, 2) if sent else 0
        lc = '#c62828' if loss_pct > 1 else ('#e65100' if loss_pct > 0.5 else '#2e7d32')
        bg = '#fafafa' if i % 2 else '#fff'
        bold = 'font-weight:600' if m == months[-1] else ''
        if pending < 0:
            pending_cell = _badge(f'+{_usd(abs(pending))} extra', '#e8f5e9', '#2e7d32')
        else:
            pending_cell = f'<span style="color:#e65100">{_usd(pending)}</span>'
        trend_rows += f'''<tr style="background:{bg}">
          {_td(m, 'left', bold)}
          {_td(_num(sent))}
          {_td(_badge(_num(lost), '#ffebee', '#c62828'))}
          {_td(f'<span style="color:{lc}">{loss_pct}%</span>')}
          {_td(_usd(expected))}
          {_td(_badge(_usd(actual), '#e8f5e9', '#2e7d32'))}
          {_td(pending_cell)}
        </tr>'''

    trend_html = f'''
    <div style="padding:20px 24px 0">
      {_section_title('📈 Last 6 Months Trend')}
      <table style="width:100%;border-collapse:collapse;font-size:12px">
        <thead><tr style="background:#1a237e;color:#fff">
          {_th('Month','left')}{_th('Shipped')}{_th('Lost')}{_th('Loss%')}{_th('Expected')}{_th('Recovered')}{_th('Pending')}
        </tr></thead>
        <tbody>{trend_rows}</tbody>
      </table>
    </div>'''

    # ── 3. Channel breakdown ───────────────────────────────────────────────
    ch_rows = ''
    for ch, rows in channels.items():
        lost     = sum(r.get('lost_stock', 0) for r in rows)
        expected = sum(r.get('expected_reimburs', 0) for r in rows)
        actual   = sum(r.get('actual_reimbursed', 0) for r in rows)
        pending  = expected - actual
        rec_rate = round(actual / expected * 100, 1) if expected else 0
        if rec_rate >= 60:
            rb = _badge(f'{rec_rate}%', '#e8f5e9', '#2e7d32')
        elif rec_rate >= 45:
            rb = _badge(f'{rec_rate}%', '#fff8e1', '#e65100')
        else:
            rb = _badge(f'{rec_rate}%', '#ffebee', '#c62828')
        if pending < 0:
            pending_str = _badge(f'+{_usd(abs(pending))} extra', '#e8f5e9', '#2e7d32')
        else:
            pending_str = f'<span style="color:#e65100">{_usd(pending)}</span>'
        ch_rows += f'''<tr>
          {_td(f'<strong>{ch}</strong>', 'left')}
          {_td(f'<span style="color:#c62828">{_num(lost)}</span>')}
          {_td(_usd(expected))}
          {_td(f'<span style="color:#2e7d32">{_usd(actual)}</span>')}
          {_td(pending_str)}
          {_td(rb)}
        </tr>'''

    channel_html = f'''
    <div style="padding:20px 24px 0">
      {_section_title('🏪 Channel Breakdown')}
      <table style="width:100%;border-collapse:collapse;font-size:12px">
        <thead><tr style="background:#1a237e;color:#fff">
          {_th('Channel','left')}{_th('Lost Units')}{_th('Expected')}{_th('Recovered')}{_th('Pending')}{_th('Recovery %')}
        </tr></thead>
        <tbody>{ch_rows}</tbody>
      </table>
    </div>'''

    # ── 4. Transporter breakdown — month-wise ─────────────────────────────
    tr_rows = ''
    row_idx = 0
    for m in months:
        month_has_data = False
        month_rows = ''
        tr_month_data = []
        for tr_name, tr_rows_list in transporters.items():
            period_idx = months.index(m)
            r = tr_rows_list[period_idx] if period_idx < len(tr_rows_list) else {}
            shipped = r.get('qty_sent', 0)
            lost    = r.get('lost_stock', 0)
            if shipped == 0 and lost == 0:
                continue
            tr_month_data.append((tr_name, shipped, lost))
        tr_month_data.sort(key=lambda x: x[2], reverse=True)
        for tr_name, shipped, lost in tr_month_data:
            loss_pct = round(lost / shipped * 100, 2) if shipped else 0
            lc = '#c62828' if loss_pct > 8 else ('#e65100' if loss_pct > 4 else '#2e7d32')
            lb = _badge(f'{loss_pct}%', '#ffebee' if loss_pct > 8 else ('#fff8e1' if loss_pct > 4 else '#e8f5e9'), lc)
            bg = '#fafafa' if row_idx % 2 else '#fff'
            month_label = m if not month_has_data else ''
            tr_rows += f'''<tr style="background:{bg}">
              {_td(f'<strong>{month_label}</strong>', 'left')}
              {_td(tr_name, 'left')}
              {_td(_num(int(shipped)))}
              {_td(f'<span style="color:#c62828">{_num(int(lost))}</span>')}
              {_td(lb)}
            </tr>'''
            month_has_data = True
            row_idx += 1

    transporter_html = f'''
    <div style="padding:20px 24px 0">
      {_section_title('🚚 Transporter Loss Breakdown — Month Wise')}
      <table style="width:100%;border-collapse:collapse;font-size:12px">
        <thead><tr style="background:#1a237e;color:#fff">
          {_th('Month','left')}{_th('Transporter','left')}{_th('Total Shipped')}{_th('Total Lost')}{_th('Loss %')}
        </tr></thead>
        <tbody>{tr_rows}</tbody>
      </table>
    </div>'''

    # ── 5. Open Cases & Aging ─────────────────────────────────────────────
    discrepancies = data.get('discrepancies', [])
    open_cases  = [d for d in discrepancies if not d.get('is_closed')]
    AGING_BUCKETS = ['0-30 days', '31-60 days', '61-90 days', '90+ days']
    aging_summary = {}
    for b in AGING_BUCKETS:
        cases_in = [d for d in open_cases if d.get('aging_bucket') == b]
        aging_summary[b] = {
            'count':   len(cases_in),
            'pending': sum(d.get('pending', 0) for d in cases_in),
        }
    total_open    = len(open_cases)
    total_pending = sum(d.get('pending', 0) for d in open_cases)

    # Aging table rows
    aging_rows = ''
    bucket_colors = {'0-30 days': '#2e7d32', '31-60 days': '#e65100', '61-90 days': '#c62828', '90+ days': '#6a1a1a'}
    for b in AGING_BUCKETS:
        info = aging_summary[b]
        if info['count'] == 0:
            continue
        color = bucket_colors[b]
        urgency = '🔴 Immediate' if b == '90+ days' else ('🟠 High' if b == '61-90 days' else ('🟡 Medium' if b == '31-60 days' else '🟢 Low'))
        aging_rows += f'''<tr>
          {_td(f'<span style="color:{color};font-weight:600">{b}</span>', 'left')}
          {_td(f'<strong>{info["count"]}</strong>')}
          {_td(f'<span style="color:#e65100;font-weight:600">{_usd(info["pending"])}</span>')}
          {_td(urgency)}
        </tr>'''

    cases_html = ''
    if total_open > 0:
        cases_html = f'''
    <div style="padding:20px 24px 0">
      {_section_title('📋 Open Cases — Aging & Pending Recovery')}
      <div style="display:flex;gap:8px;margin-bottom:12px;flex-wrap:wrap">
        <div style="background:#ffebee;border-radius:6px;padding:12px 18px;text-align:center;flex:1;min-width:120px">
          <div style="font-size:22px;font-weight:700;color:#c62828">{total_open}</div>
          <div style="font-size:11px;color:#888;margin-top:2px">Open Cases</div>
        </div>
        <div style="background:#fff8e1;border-radius:6px;padding:12px 18px;text-align:center;flex:1;min-width:120px">
          <div style="font-size:22px;font-weight:700;color:#e65100">{_usd(total_pending)}</div>
          <div style="font-size:11px;color:#888;margin-top:2px">Total Owed</div>
          <div style="font-size:10px;color:#bbb;margin-top:3px;font-style:italic">Open cases only. Higher than Pending Recovery above as it excludes over-recovery offsets from closed cases.</div>
        </div>
        <div style="background:#fce4ec;border-radius:6px;padding:12px 18px;text-align:center;flex:1;min-width:120px">
          <div style="font-size:22px;font-weight:700;color:#6a1a1a">{aging_summary["90+ days"]["count"]}</div>
          <div style="font-size:11px;color:#888;margin-top:2px">90+ Days (Critical)</div>
        </div>
      </div>
      <table style="width:100%;border-collapse:collapse;font-size:12px">
        <thead><tr style="background:#1a237e;color:#fff">
          {_th('Aging Bucket','left')}{_th('Cases')}{_th('Amount Owed')}{_th('Action Priority')}
        </tr></thead>
        <tbody>{aging_rows}</tbody>
      </table>
    </div>'''

    # ── 6. Recoveries / Wins (closed cases where actual_reimbursed > 0) ────
    closed_with_recovery = [
        d for d in discrepancies
        if d.get('is_closed') and not d.get('is_rejected') and d.get('actual_reimbursed', 0) > 0
    ]
    # Group by month
    from collections import defaultdict as _dd
    wins_by_month = _dd(lambda: {'count': 0, 'actual': 0.0, 'expected': 0.0})
    for d in closed_with_recovery:
        m = d['month']
        wins_by_month[m]['count']    += 1
        wins_by_month[m]['actual']   += d.get('actual_reimbursed', 0)
        wins_by_month[m]['expected'] += d.get('expected_reimburs', 0)

    wins_rows = ''
    for m in sorted(wins_by_month, key=lambda x: months.index(x) if x in months else 999):
        w = wins_by_month[m]
        wins_rows += f'''<tr>
          {_td(f'<strong>{m}</strong>', 'left')}
          {_td(str(w["count"]))}
          {_td(f'<span style="color:#2e7d32;font-weight:600">{_usd(w["actual"])}</span>')}
          {_td(_usd(w["expected"]))}
          {_td(_badge(f'{round(w["actual"]/w["expected"]*100,1)}%', "#e8f5e9", "#2e7d32") if w["expected"] else "—")}
        </tr>'''

    wins_total_actual   = sum(w['actual']   for w in wins_by_month.values())
    wins_total_expected = sum(w['expected'] for w in wins_by_month.values())
    wins_total_cases    = sum(w['count']    for w in wins_by_month.values())

    recoveries_html = ''
    if wins_rows:
        recoveries_html = f'''
    <div style="padding:20px 24px 0">
      {_section_title('✅ Recovered — Closed Cases (2026)')}
      <div style="display:flex;gap:8px;margin-bottom:12px;flex-wrap:wrap">
        <div style="background:#e8f5e9;border-radius:6px;padding:12px 18px;text-align:center;flex:1;min-width:120px">
          <div style="font-size:22px;font-weight:700;color:#2e7d32">{wins_total_cases}</div>
          <div style="font-size:11px;color:#888;margin-top:2px">Cases Recovered</div>
        </div>
        <div style="background:#e8f5e9;border-radius:6px;padding:12px 18px;text-align:center;flex:1;min-width:120px">
          <div style="font-size:22px;font-weight:700;color:#2e7d32">{_usd(kpis.get('actual_recovered', 0))}</div>
          <div style="font-size:11px;color:#888;margin-top:2px">Total Recovered</div>
        </div>
        <div style="background:#f3e5f5;border-radius:6px;padding:12px 18px;text-align:center;flex:1;min-width:120px">
          <div style="font-size:22px;font-weight:700;color:#6a1b9a">{kpis.get('recovery_rate', 0)}%</div>
          <div style="font-size:11px;color:#888;margin-top:2px">Recovery Rate</div>
        </div>
      </div>
      <table style="width:100%;border-collapse:collapse;font-size:12px">
        <thead><tr style="background:#2e7d32;color:#fff">
          {_th('Month','left')}{_th('Cases Closed')}{_th('Recovered')}{_th('Expected')}{_th('Rate')}
        </tr></thead>
        <tbody>{wins_rows}</tbody>
      </table>
    </div>'''

    # ── 7. Carrier Fault Cases ────────────────────────────────────────────
    CARRIER_KEYWORDS = ['freightzen loss', 'ups loss', 'dhl loss']
    carrier_cases = [
        d for d in discrepancies
        if any(kw in d.get('remark', '').lower() for kw in CARRIER_KEYWORDS)
    ]

    carrier_html = ''
    if carrier_cases:
        from collections import defaultdict as _cdd
        # Group by carrier name
        by_carrier = _cdd(lambda: {'months': set(), 'total_lost': 0, 'total_shipped': 0, 'pending': 0.0, 'cases': 0})
        for d in carrier_cases:
            c = by_carrier[d['transporter']]
            c['months'].add(d['month'])
            c['total_lost']    += d.get('lost_stock', 0)
            c['pending']       += d.get('pending', 0)
            c['cases']         += 1
        # Pull total shipped per carrier from transporter_data
        tr_data_raw = data.get('transporter_data', {}).get('transporters', {})
        for carrier_name, info in by_carrier.items():
            rows = tr_data_raw.get(carrier_name, [])
            info['total_shipped'] = int(sum(r.get('qty_sent', 0) for r in rows))

        carrier_rows = ''
        for i, (carrier_name, info) in enumerate(sorted(by_carrier.items(), key=lambda x: -x[1]['total_lost'])):
            bg = '#fafafa' if i % 2 else '#fff'
            months_str = ', '.join(sorted(info['months']))
            carrier_rows += f'''<tr style="background:{bg}">
              {_td(f'<strong>{carrier_name}</strong>', 'left')}
              {_td(months_str, 'left')}
              {_td(_num(info['total_shipped']))}
              {_td(f'<span style="color:#c62828;font-weight:700">{int(info["total_lost"])}</span>')}
              {_td(f'<span style="color:#e65100;font-weight:600">{_usd(info["pending"])}</span>')}
            </tr>'''

        carrier_html = f'''
    <div style="padding:20px 24px 0">
      {_section_title('🚚 Carrier Fault Cases')}
      <table style="width:100%;border-collapse:collapse;font-size:12px">
        <thead><tr style="background:#37474f;color:#fff">
          {_th('Carrier','left')}{_th('Months Affected','left')}{_th('Total Shipped')}{_th('Total Lost')}{_th('Pending')}
        </tr></thead>
        <tbody>{carrier_rows}</tbody>
      </table>
    </div>'''

    # ── 8. Shipbob Outward Loss (D2C) ────────────────────────────────────
    outward_loss = data.get('outward_loss', {})
    ol_headers = outward_loss.get('headers', [])
    ol_rows    = outward_loss.get('rows', [])

    outward_loss_html = ''
    if ol_headers and ol_rows:
        def _ol_status(val):
            if not val or not val.strip():
                return '—'
            clean = val.strip().lstrip('✅🔄⬜ ').strip()
            v = clean.lower()
            if 'recover' in v or v in ('resolved', 'closed', 'done'):
                return f'<span style="color:#2e7d32;font-weight:600">✅ {clean}</span>'
            elif 'progress' in v or v in ('open', 'pending'):
                return f'<span style="color:#e65100;font-weight:600">🔄 {clean}</span>'
            return clean

        # Compute summary stats (skip TOTAL row)
        data_rows = [r for r in ol_rows if r and str(r[0]).strip() not in ('', '—', '-') and str(r[1]).strip().lower() != 'total']
        total_units    = 0
        total_value    = 0.0
        recovered_cnt  = 0
        inprogress_cnt = 0
        for r in data_rows:
            padded = r + [''] * max(0, 5 - len(r))
            try: total_units += int(str(padded[2]).replace(',','').strip()) if padded[2] and padded[2] not in ('-','—') else 0
            except: pass
            try:
                v = str(padded[3]).replace(',','').replace('$','').strip()
                total_value += float(v) if v and v not in ('-','—') else 0.0
            except: pass
            st = str(padded[4]).strip().lower() if len(padded) > 4 else ''
            if 'recover' in st or st in ('resolved','closed','done'): recovered_cnt += 1
            elif 'progress' in st or st in ('open','pending'): inprogress_cnt += 1

        body_rows = ''
        for i, row in enumerate(ol_rows):
            if not row: continue
            padded = row + [''] * max(0, len(ol_headers) - len(row))
            sr     = padded[0] if padded else ''
            desc   = padded[1] if len(padded) > 1 else ''
            units  = padded[2] if len(padded) > 2 else ''
            value  = padded[3] if len(padded) > 3 else ''
            status = padded[4] if len(padded) > 4 else ''
            notes  = padded[5] if len(padded) > 5 else ''
            is_total = str(desc).strip().upper() == 'TOTAL'
            bg = '#e8f5e9' if is_total else ('#fafafa' if i % 2 else '#fff')
            fw = 'font-weight:700' if is_total else ''
            value_str = f'<span style="color:#1565c0;font-weight:600">{value}</span>' if value and value not in ('-','—') else '—'
            units_str = f'<strong>{units}</strong>' if is_total else (units or '—')
            body_rows += f'''<tr style="background:{bg}">
              <td style="padding:8px 10px;border-bottom:1px solid #f0f0f0;color:#999;font-size:11px;{fw}">{sr or '—'}</td>
              <td style="padding:8px 10px;border-bottom:1px solid #f0f0f0;{fw}">{desc or '—'}</td>
              <td style="padding:8px 10px;border-bottom:1px solid #f0f0f0;text-align:center;{fw}">{units_str}</td>
              <td style="padding:8px 10px;border-bottom:1px solid #f0f0f0;text-align:center;{fw}">{value_str}</td>
              <td style="padding:8px 10px;border-bottom:1px solid #f0f0f0;text-align:center">{_ol_status(status) if not is_total else '—'}</td>
              <td style="padding:8px 10px;border-bottom:1px solid #f0f0f0;font-size:11px;color:#777">{notes or '—'}</td>
            </tr>'''

        outward_loss_html = f'''
    <div style="padding:20px 24px 0">
      {_section_title('📦 Shipbob Outward Loss — D2C Orders')}
      <div style="display:flex;gap:8px;margin-bottom:12px;flex-wrap:wrap">
        <div style="background:#e3f2fd;border-radius:6px;padding:12px 18px;text-align:center;flex:1;min-width:110px">
          <div style="font-size:20px;font-weight:700;color:#1565c0">{total_units:,}</div>
          <div style="font-size:11px;color:#888;margin-top:2px">Total Order ID count</div>
        </div>
        <div style="background:#e8f5e9;border-radius:6px;padding:12px 18px;text-align:center;flex:1;min-width:110px">
          <div style="font-size:20px;font-weight:700;color:#2e7d32">${total_value:,.2f}</div>
          <div style="font-size:11px;color:#888;margin-top:2px">Total Value</div>
        </div>
        <div style="background:#e8f5e9;border-radius:6px;padding:12px 18px;text-align:center;flex:1;min-width:110px">
          <div style="font-size:20px;font-weight:700;color:#2e7d32">{recovered_cnt}</div>
          <div style="font-size:11px;color:#888;margin-top:2px">Batches Recovered</div>
        </div>
        <div style="background:#fff8e1;border-radius:6px;padding:12px 18px;text-align:center;flex:1;min-width:110px">
          <div style="font-size:20px;font-weight:700;color:#e65100">{inprogress_cnt}</div>
          <div style="font-size:11px;color:#888;margin-top:2px">In Progress</div>
        </div>
      </div>
      <table style="width:100%;border-collapse:collapse;font-size:12px">
        <thead><tr style="background:#1565c0;color:#fff">
          <th style="padding:8px 10px;text-align:left;font-weight:500">#</th>
          <th style="padding:8px 10px;text-align:left;font-weight:500">Description</th>
          <th style="padding:8px 10px;text-align:center;font-weight:500">Order ID count</th>
          <th style="padding:8px 10px;text-align:center;font-weight:500">Value ($)</th>
          <th style="padding:8px 10px;text-align:center;font-weight:500">Status</th>
          <th style="padding:8px 10px;text-align:left;font-weight:500">Notes</th>
        </tr></thead>
        <tbody>{body_rows}</tbody>
      </table>
    </div>'''

    # ── 9. Urgency alerts ──────────────────────────────────────────────────
    alerts = []
    for ch, rows in channels.items():
        exp = sum(r.get('expected_reimburs', 0) for r in rows)
        act = sum(r.get('actual_reimbursed', 0) for r in rows)
        rate = round(act / exp * 100, 1) if exp else 0
        if rate < 50:
            alerts.append(f'<b>{ch}</b> recovery rate is <b>{rate}%</b> — below 50% threshold, follow-up needed.')
    for tr_name, tr_rows_list in transporters.items():
        shipped = sum(r.get('qty_sent', 0) for r in tr_rows_list)
        lost    = sum(r.get('lost_stock', 0) for r in tr_rows_list)
        if not shipped:
            continue
        loss_pct = round(lost / shipped * 100, 2)
        if loss_pct > 8:
            alerts.append(f'<b>{tr_name}</b> loss rate is <b>{loss_pct}%</b> — highest carrier loss, investigate.')
    pending_total = kpis.get('pending_recovery', 0)
    if pending_total > 10000:
        alerts.append(f'Total pending recovery is <b>{_usd(pending_total)}</b> — open cases need to be followed up.')

    alert_html = ''
    if alerts:
        items = ''.join(f'<li style="margin-bottom:6px">{a}</li>' for a in alerts)
        alert_html = f'''
        <div style="padding:20px 24px 0">
          <div style="background:#fff3e0;border-left:4px solid #e65100;border-radius:4px;padding:14px 16px">
            <div style="font-weight:700;color:#e65100;font-size:13px;margin-bottom:8px">⚠️ Attention Required</div>
            <ul style="margin:0;padding-left:18px;font-size:12px;color:#555;line-height:1.8">{items}</ul>
          </div>
        </div>'''

    # ── Assemble ───────────────────────────────────────────────────────────
    return f"""<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;background:#f5f5f5;margin:0;padding:20px">
<div style="max-width:680px;margin:0 auto;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.1)">

  <!-- Header -->
  <div style="background:#1a237e;padding:24px;text-align:center">
    <h1 style="color:#fff;margin:0;font-size:20px">India → USA Inventory Reconciliation</h1>
    <p style="color:#90caf9;margin:6px 0 0;font-size:13px">Weekly Report — {today}</p>
  </div>

  <!-- KPIs -->
  <div style="padding:20px 24px 0">
    {_section_title('Overall Summary')}
    <div style="color:#999;font-size:11px;margin-bottom:8px;text-align:center">{months[0] if months else ''} – {months[-2] if len(months) >= 2 else (months[-1] if months else '')}</div>
    <div style="display:grid;grid-template-columns:repeat(2,1fr);gap:8px">
      <div style="background:#e3f2fd;padding:18px;text-align:center;border-radius:6px">
        <div style="font-size:26px;font-weight:700;color:#1565c0">{_num(kpis.get('total_shipped',0))}</div>
        <div style="color:#666;font-size:12px;margin-top:4px">Total Units Shipped</div>
        <div style="color:#aaa;font-size:10px;margin-top:2px">{months[0] if months else ''} – {months[-2] if len(months) >= 2 else ''}</div>
      </div>
      <div style="background:#ffebee;padding:18px;text-align:center;border-radius:6px">
        <div style="font-size:26px;font-weight:700;color:#c62828">{_num(kpis.get('total_lost',0))}</div>
        <div style="color:#666;font-size:12px;margin-top:4px">Total Units Lost ({kpis.get('loss_rate',0)}%)</div>
        <div style="color:#aaa;font-size:10px;margin-top:2px">{months[0] if months else ''} – {months[-2] if len(months) >= 2 else ''}</div>
      </div>
      <div style="background:#fff8e1;padding:18px;text-align:center;border-radius:6px">
        <div style="font-size:26px;font-weight:700;color:#e65100">{_usd(kpis.get('pending_recovery',0))}</div>
        <div style="color:#666;font-size:12px;margin-top:4px">Pending Recovery</div>
        <div style="color:#aaa;font-size:10px;margin-top:2px">{months[0] if months else ''} – {months[-2] if len(months) >= 2 else ''}</div>
        <div style="color:#bbb;font-size:10px;margin-top:3px;font-style:italic">Net of over-recoveries from closed cases. See Open Cases section for exact amount owed.</div>
      </div>
      <div style="background:#e8f5e9;padding:18px;text-align:center;border-radius:6px">
        <div style="font-size:26px;font-weight:700;color:#2e7d32">{_usd(kpis.get('actual_recovered',0))} <span style="font-size:14px">({kpis.get('recovery_rate',0)}%)</span></div>
        <div style="color:#666;font-size:12px;margin-top:4px">Recovered</div>
        <div style="color:#aaa;font-size:10px;margin-top:2px">{months[0] if months else ''} – {months[-2] if len(months) >= 2 else ''}</div>
      </div>
    </div>
  </div>

  {mom_html}
  {trend_html}
  {channel_html}
  {cases_html}
  {carrier_html}
  {recoveries_html}
  {outward_loss_html}
  {alert_html}

  <!-- CTA -->
  <div style="background:#f5f5f5;padding:24px;text-align:center;border-top:1px solid #e0e0e0;margin-top:20px">
    <a href="{dashboard_url}" style="background:#1a237e;color:#fff;padding:12px 32px;border-radius:4px;text-decoration:none;font-size:14px;font-weight:600">
      View Live Dashboard
    </a>
    <p style="color:#999;font-size:11px;margin:12px 0 0">Auto-generated every Monday 12:01 PM IST</p>
  </div>

</div>
</body>
</html>"""
