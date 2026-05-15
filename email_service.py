import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
from googleapiclient.discovery import build


def send_weekly_report(creds, data, recipients, sender, dashboard_url):
    service = build('gmail', 'v1', credentials=creds)
    kpis = data.get('kpis', {})
    subject = f"Weekly Inventory Reconciliation Report — {datetime.now().strftime('%d %b %Y')}"
    html = build_email_html(kpis, data, dashboard_url)

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
