# send_operational_pivot_report.py
# ─────────────────────────────────
# Builds the "Operational Pivot Views" tables (Shipment Flow / GRN Reconciliation
# / Claim Exposure) for both India->US and US->US, from July to date, and saves it
# as a Gmail DRAFT -- never sent automatically. You open Gmail > Drafts, review it,
# and send it yourself.
#
# If operational_pivot_thread.json exists (see email_service.load_reply_thread()),
# the draft is created as a reply inside that existing thread instead of a new one.
#
# TEST MODE (default): drafts are addressed only to anand.r@mosaicwellness.in.
# Once verified, add the finance team to RECIPIENTS below and set TEST_MODE=False.
#
# Scheduled: Windows Task "MosaicOperationalPivotReport", every Friday at 17:00 (changed from Saturday 13:00).
#
# Run:
#     python send_operational_pivot_report.py

import json
import os

from auth import get_sheets_credentials, get_gmail_credentials
from sheets_service import get_india_us_data, get_us2us_data
from email_service import draft_operational_pivot_email, load_reply_thread

CONFIG_FILE = os.path.join(os.path.dirname(__file__), 'config.json')

# ── Recipients ────────────────────────────────────────────────────────────────
TEST_MODE = True   # True = send only to TEST_RECIPIENT; False = send to FINANCE_RECIPIENTS
TEST_RECIPIENT = 'anand.r@mosaicwellness.in'
FINANCE_RECIPIENTS = [
    # Add finance team email addresses here once the test looks good, e.g.:
    # 'finance1@mosaicwellness.in',
    # 'finance2@mosaicwellness.in',
]

SINCE_MONTH = 'July'

with open(CONFIG_FILE) as f:
    CONFIG = json.load(f)

CF = CONFIG.get('credentials_file')
TF = CONFIG.get('token_file')
RECON_ID = CONFIG.get('recon_sheet_id')


def main():
    recipients = [TEST_RECIPIENT] if TEST_MODE else FINANCE_RECIPIENTS
    if not recipients:
        print("No recipients configured (FINANCE_RECIPIENTS is empty and TEST_MODE is False). Nothing sent.")
        return

    print("Fetching India→US data …")
    sheets_creds = get_sheets_credentials(CF, TF)
    ius = get_india_us_data(sheets_creds, RECON_ID)
    print(f"  {len(ius.get('rows', []))} rows")

    print("Fetching US→US data …")
    u2u = get_us2us_data(sheets_creds, RECON_ID)
    print(f"  {len(u2u.get('rows', []))} rows")

    thread = load_reply_thread()
    print(f"Drafting for: {recipients} (TEST_MODE={TEST_MODE}, thread={'yes, ' + thread['subject'] if thread else 'no, new conversation'})")
    gmail_creds = get_gmail_credentials(CF, TF)
    draft_operational_pivot_email(
        gmail_creds,
        ius.get('rows', []),
        u2u.get('rows', []),
        recipients,
        CONFIG.get('email_from'),
        since_month=SINCE_MONTH,
        thread=thread,
    )
    print("Done.")


if __name__ == '__main__':
    main()
