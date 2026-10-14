"""
ShipBob mails. Two mails, each with an Excel attached and saved as a Gmail DRAFT (never sent):

  potential  Potential In-Transit Loss: not delivered, 11-15 days since import. Daily, when there are new orders.
  claims     Claims follow-up, one mail for two lists:
               - Pending to claim: orders we cannot raise on the ShipBob portal (ask the reason / status)
               - Claim raised but not received: ageing-wise, follow-up on the claim SLA (45-60 days)
             The claim raised date is taken as import date + 15 days (claim age = order age - 15).

  build(kind)         Excel + email text
  create_draft(kind)  saves a Gmail draft with the Excel attached (updates the same-day draft if one exists)
  auto_run()          runs by itself when the data updates: builds both, drafts the ones that are due
  mark_sent(kind)     you confirm you sent it (draft -> sent in the log)

State (per mail: what was built / drafted) is in state/mailer_state.json,
the per-order history in state/mailer_log.csv (kind = potential / pending / raised, matching the order lists).
"""
import base64
import html as _html
import json
import os
from datetime import date, datetime
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from . import engine

OUT_DIR = os.path.join(engine.BASE, "output", "mailer")
STATE_FILE = os.path.join(engine.BASE, "state", "mailer_state.json")
GMAIL_SCOPES = ["https://www.googleapis.com/auth/gmail.compose"]

KINDS = ("potential", "claims")
TITLES = {"potential": "Potential in-transit loss", "claims": "Claims follow-up"}
FILES = {"potential": "Potential_InTransit_Loss", "claims": "Claims_Followup"}

BASE_COLS = [("Shipment ID", 13), ("Store Order ID", 18), ("Import Date", 13), ("Age", 12), ("Carrier", 16),
             ("Tracking Information", 28), ("Delivery Status", 42)]
EXTRA_COLS = {"potential": [],
              "pending": [("Expected Claim $", 14), ("Claims remark", 40)],
              "raised": [("Claim Age", 16), ("SLA status", 20), ("Expected Claim $", 14), ("Claims remark", 40)]}
HEADERS = {"Age": "Days Since Import", "Tracking Information": "Tracking #", "Expected Claim $": "Expected Claim ($)",
           "Mailed": "First mailed", "Claim Age": "Days Since Claim Raised"}
BANDS = [("16-30 days", 16, 30), ("31-60 days", 31, 60), ("61-90 days", 61, 90), ("91+ days", 91, 99999)]


# ------------------------------------------------------------------ state
def load_state():
    st = {}
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, encoding="utf-8") as f:
            st = json.load(f)
    if "kinds" not in st:
        st = {"kinds": {}}
    return st


def kstate(kind):
    return load_state()["kinds"].get(kind, {})


def save_kstate(kind, **changes):
    st = load_state()
    st["kinds"].setdefault(kind, {}).update(changes)
    os.makedirs(os.path.dirname(STATE_FILE), exist_ok=True)
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(st, f, indent=1)
    return st["kinds"][kind]


def status(cfg=None):
    """What the dashboard shows in the mail panel."""
    cfg = cfg or engine.load_config()
    out = {"gmail_connected": gmail_connected(cfg), "kinds": {}}
    for k in KINDS:
        s = dict(kstate(k))
        s.pop("ids_by_kind", None)
        s["excel_exists"] = bool(s.get("excel_path")) and os.path.exists(s["excel_path"])
        s["mail_to"] = _to(k, cfg)
        s["title"] = TITLES[k]
        out["kinds"][k] = s
    return out


def _to(kind, cfg):
    return cfg.get(f"mail_to_{kind}") or cfg.get("mail_to", "")


def _cc(kind, cfg):
    return cfg.get(f"mail_cc_{kind}") or cfg.get("mail_cc", "")


# ------------------------------------------------------------------ build
def sla_status(claim_age, cfg):
    lo, hi = cfg.get("claim_sla_days", [45, 60])
    return "Over %d days" % hi if claim_age > hi else ("%d-%d days" % (lo, hi) if claim_age >= lo else "Within %d days" % lo)


def parts(df, kind, cfg):
    """The order lists inside a mail: {part kind: frame}, oldest first. Part kinds match the dashboard lists."""
    if kind == "potential":
        return {"potential": df[df["Potential Loss"]].sort_values("Age", ascending=False)}
    pend = df[df["Sub"] == "Pending to claim"].sort_values("Age", ascending=False)
    raised = df[df["Sub"] == "Claim raised but not received"].sort_values("Age", ascending=False).copy()
    raised["SLA status"] = [sla_status(a, cfg) for a in raised["Claim Age"]]
    return {"pending": pend, "raised": raised}


def _sheet(wb, title, d, part, first=False):
    ws = wb.active if first else wb.create_sheet()
    ws.title = title[:31]
    cols = BASE_COLS + EXTRA_COLS[part] + [("Mailed", 16)]
    head_fill = PatternFill("solid", fgColor="1F3864")
    for i, (c, w) in enumerate(cols, 1):
        cell = ws.cell(row=1, column=i, value=HEADERS.get(c, c))
        cell.font = Font(name="Arial", bold=True, color="FFFFFF")
        cell.fill = head_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = w
    for r, (_, x) in enumerate(d.iterrows(), 2):
        for i, (c, _w) in enumerate(cols, 1):
            v = x[c]
            if pd.isna(v) or v == "":
                v = None
            elif isinstance(v, pd.Timestamp):
                v = v.to_pydatetime()
            elif hasattr(v, "item"):
                v = v.item()
            cell = ws.cell(row=r, column=i, value=v)
            cell.font = Font(name="Arial", size=10)
            if c == "Import Date":
                cell.number_format = "dd-mmm-yyyy"
            elif c == "Expected Claim $":
                cell.number_format = "#,##0.00"
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{max(len(d) + 1, 2)}"
    ws.row_dimensions[1].height = 30
    return ws


SHEET_NAMES = {"potential": "Potential loss", "pending": "Pending to claim", "raised": "Raised, not received"}


def count_new(kind, pp, data_date):
    """Orders not mailed before. If we already drafted this mail for the same data date, those entries
    still count as new, so refreshing the draft does not change what it says."""
    ks = kstate(kind)
    same_day = ks.get("drafted_at", "")[:10] if ks.get("drafted_for") == data_date else None
    log = engine.load_mail_log()
    n = 0
    for part, d in pp.items():
        entries = log.get(part, {})
        for sid in d["Shipment ID"]:
            e = entries.get(int(sid))
            if e is None or (same_day and e["status"] == "drafted" and str(e["date"])[:10] == same_day):
                n += 1
    return n


def build(kind, df, meta, cfg=None):
    """Build the Excel and the email text for one mail. Does not touch Gmail or the log."""
    cfg = cfg or engine.load_config()
    pp = parts(df, kind, cfg)
    n_all = sum(len(d) for d in pp.values())
    n_new = count_new(kind, pp, f"{meta['as_of']:%Y-%m-%d}")
    os.makedirs(OUT_DIR, exist_ok=True)
    path = os.path.join(OUT_DIR, f"ShipBob_{FILES[kind]}_{meta['as_of']:%Y-%m-%d}.xlsx")
    wb = Workbook()
    for i, (part, d) in enumerate(pp.items()):
        _sheet(wb, f"{SHEET_NAMES[part]} ({len(d)})", d, part, first=(i == 0))
    wb.save(path)
    subject, blocks = (_email_potential if kind == "potential" else _email_claims)(pp, n_new, meta, cfg)
    body, html = render_plain(blocks), render_html(blocks)
    with open(path.replace(".xlsx", "_email.txt"), "w", encoding="utf-8") as f:
        f.write(f"Subject: {subject}\n\n{body}\n")
    ids = {part: [int(x) for x in d["Shipment ID"]] for part, d in pp.items()}
    value = round(sum(float(d["Expected Claim $"].sum()) for p, d in pp.items() if p != "potential"), 2) if kind == "claims" else None
    lo = cfg.get("claim_sla_days", [45, 60])[0]
    aged = int((pp["raised"]["Claim Age"] >= lo).sum()) if kind == "claims" else None
    save_kstate(kind, excel_path=path, excel_built_at=datetime.now().isoformat(timespec="seconds"),
                data_date=f"{meta['as_of']:%Y-%m-%d}", new_count=n_new, all_count=n_all, value=value, aged=aged,
                pending_count=len(pp.get("pending", [])) if kind == "claims" else None,
                raised_count=len(pp.get("raised", [])) if kind == "claims" else None,
                subject=subject, body=body, html=html, ids_by_kind=ids)
    return {"kind": kind, "path": path, "all": n_all, "new": n_new, "subject": subject, "body": body, "html": html,
            "ids_by_kind": ids}


# ------------------------------------------------------------------ email text
# A mail is a list of blocks, rendered both as plain text and as HTML (headings in bold):
#   ("P", text[, tight])  paragraph       ("H", text)  bold heading      ("L", text)  bold label above a list
#   ("UL", [items])       bullet list     ("OL", [items])  a. b. c. list
# tight=True / H / L are followed directly by the next block (no blank line).
def render_plain(blocks):
    out = []
    for blk in blocks:
        kind, val = blk[0], blk[1]
        if kind in ("P", "H", "L"):
            out.append(val)
        elif kind == "UL":
            out.append("\n".join(f"  - {i}" for i in val))
        else:
            out.append("\n".join(f"  {chr(97 + n)}. {i}" for n, i in enumerate(val)))
    text = ""
    for n, blk in enumerate(blocks):
        text += out[n]
        if n < len(blocks) - 1:
            tight = blk[0] in ("H", "L") or (blk[0] == "P" and len(blk) > 2 and blk[2])
            text += "\n" if tight else "\n\n"
    return text


def render_html(blocks):
    h = lambda t: _html.escape(str(t)).replace("\n", "<br>")
    parts = []
    for blk in blocks:
        kind, val = blk[0], blk[1]
        tight = len(blk) > 2 and blk[2]
        if kind == "P":
            parts.append(f'<p style="margin:0 0 {4 if tight else 12}px">{h(val)}</p>')
        elif kind in ("H", "L"):
            parts.append(f'<p style="margin:{16 if kind == "H" else 12}px 0 4px"><b>{h(val)}</b></p>')
        elif kind == "UL":
            parts.append('<ul style="margin:0 0 12px;padding-left:22px">' + "".join(f"<li>{h(i)}</li>" for i in val) + "</ul>")
        else:
            parts.append('<ol type="a" style="margin:0 0 12px;padding-left:22px">' + "".join(f"<li>{h(i)}</li>" for i in val) + "</ol>")
    return ('<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;line-height:1.5;color:#222">'
            + "".join(parts) + "</div>")


def _counts(series):
    return [f"{k}: {v}" for k, v in series.items()]


def _band_counts(allf):
    return [f"{b}: {int(((allf['Age'] >= lo) & (allf['Age'] <= hi)).sum())}" for b, lo, hi in BANDS]


def _newtxt(n_all, n_new):
    """Empty when everything is new (the first mail), else a sentence about what is new."""
    if n_new == n_all:
        return ""
    if n_new == 0:
        return " All of them were shared earlier and are still open."
    return f" {n_new} of them are new since our last mail; the rest were shared earlier and are still open."


def _sign(cfg):
    return ("P", f"Thanks,\n{cfg.get('mail_signature', '')}")


def _email_potential(pp, n_new, meta, cfg):
    allf = pp["potential"]
    A, T = cfg["alert_from_day"], cfg["transit_days"]
    new_in_subject = f" ({n_new} new)" if 0 < n_new < len(allf) else ""
    subject = (f"ShipBob D2C - {len(allf)} orders not delivered {A}-{T} days after import"
               f"{new_in_subject} - {meta['as_of']:%d %b %Y}")
    return subject, [
        ("P", "Hi Team,"),
        ("P", f"Please find attached the D2C orders that are still not delivered {A}-{T} days after import "
              f"(data as of {meta['as_of']:%d %b %Y}).{_newtxt(len(allf), n_new)} They move into dispute after {T} days, "
              f"so we would appreciate a status check with the carriers before then."),
        ("L", "By latest carrier status:"), ("UL", _counts(allf["Stage"].value_counts())),
        ("L", "By carrier:"), ("UL", _counts(allf["Carrier"].fillna("Not assigned").value_counts())),
        ("P", "Could you please share an update / ETA for each order, and flag any that are lost so we can proceed with the claim?"),
        _sign(cfg),
    ]


def _email_claims(pp, n_new, meta, cfg):
    pend, raised = pp["pending"], pp["raised"]
    T = cfg["transit_days"]
    lo, hi = cfg.get("claim_sla_days", [45, 60])
    n_all = len(pend) + len(raised)
    subject = (f"ShipBob D2C - claims follow-up: {len(pend)} pending to claim, {len(raised)} raised not received "
               f"({meta['as_of']:%d %b %Y})")
    blocks = [("P", "Hi Team,"),
              ("P", f"Two updates on our D2C claims (data as of {meta['as_of']:%d %b %Y}). "
                    f"The order-wise lists are in the attached sheet, one tab for each.{_newtxt(n_all, n_new)}")]
    sec = 1
    if len(pend):
        blocks += [
            ("H", f"{sec}. Pending to claim: unable to raise on the portal"),
            ("P", f"We are unable to raise claims on the ShipBob portal for {len(pend)} orders. These orders are not "
                  f"delivered and are more than {T} days past import."),
            ("L", "Ageing (days since order import):"), ("UL", _band_counts(pend)),
            ("P", "Could you please:", True),
            ("OL", ["Let us know the reason we are unable to raise the claim on these orders, and",
                    "Share the current status of each order, or raise the claim from your side if that is possible."])]
        sec += 1
    if len(raised):
        within = int((raised["Claim Age"] < lo).sum())
        window = int(((raised["Claim Age"] >= lo) & (raised["Claim Age"] <= hi)).sum())
        over = int((raised["Claim Age"] > hi).sum())
        blocks += [
            ("H", f"{sec}. Claim raised, amount not received"),
            ("P", f"We have raised claims for {len(raised)} orders and the amount has not been received yet."),
            ("L", "Ageing since the claim was raised:"),
            ("UL", [f"Under {lo} days: {within}", f"{lo}-{hi} days: {window}", f"Over {hi} days: {over}"])]
        if window + over:
            blocks.append(("P", f"As per the claim SLA of {lo}-{hi} days, the timeline has passed for {window + over} of these "
                                f"orders and the amount is still not received ({over} of them are already past {hi} days). "
                                f"Could you please resolve these on priority and share the status / expected credit date?"))
        if within:
            blocks.append(("P", f"For the remaining {within} orders (still within the SLA), could you please check "
                                f"whether there is any update on them?"))
    blocks.append(_sign(cfg))
    return subject, blocks


# ------------------------------------------------------------------ gmail
def gmail_connected(cfg=None):
    cfg = cfg or engine.load_config()
    token = cfg["gmail_token_file"]
    if not os.path.exists(token):
        return False
    try:
        from google.oauth2.credentials import Credentials
        creds = Credentials.from_authorized_user_file(token, GMAIL_SCOPES)
        return bool(creds and (creds.valid or creds.refresh_token))
    except Exception:
        return False


def _gmail_creds(cfg, interactive=False):
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow

    token = cfg["gmail_token_file"]
    creds = Credentials.from_authorized_user_file(token, GMAIL_SCOPES) if os.path.exists(token) else None
    if creds and not creds.valid and creds.refresh_token:
        creds.refresh(Request())
    if not creds or not creds.valid:
        if not interactive:
            raise RuntimeError("Gmail is not connected. Click Connect Gmail on the dashboard once.")
        flow = InstalledAppFlow.from_client_secrets_file(cfg["credentials_file"], GMAIL_SCOPES)
        creds = flow.run_local_server(port=0)    # opens the browser once to approve "create drafts"
    with open(token, "w") as f:
        f.write(creds.to_json())
    return creds


def connect_gmail(cfg=None):
    _gmail_creds(cfg or engine.load_config(), interactive=True)
    return True


def create_draft(result, cfg=None):
    """Save the mail as a Gmail draft (not sent): text + HTML body (bold headings) with the Excel attached, and log
    the orders as drafted. If a draft for the same data date already exists it is updated instead of duplicated;
    To / Cc already typed into that draft are kept when config.json has none."""
    from googleapiclient.discovery import build as gbuild

    cfg = cfg or engine.load_config()
    kind = result["kind"]
    service = gbuild("gmail", "v1", credentials=_gmail_creds(cfg))
    ks = kstate(kind)
    existing = ks.get("draft_id") if ks.get("drafted_for") == ks.get("data_date") else None

    to, cc, bcc = _to(kind, cfg), _cc(kind, cfg), ""
    if existing:
        try:
            hdrs = service.users().drafts().get(userId="me", id=existing, format="metadata").execute()["message"]["payload"]["headers"]
            got = {h["name"].lower(): h["value"] for h in hdrs}
            to, cc, bcc = to or got.get("to", ""), cc or got.get("cc", ""), got.get("bcc", "")
        except Exception:
            existing = None      # draft was sent or deleted meanwhile

    msg = MIMEMultipart("mixed")
    msg["Subject"] = result["subject"]
    for name, val in (("To", to), ("Cc", cc), ("Bcc", bcc)):
        if val:
            msg[name] = val
    alt = MIMEMultipart("alternative")
    alt.attach(MIMEText(result["body"], "plain", "utf-8"))
    alt.attach(MIMEText(result["html"], "html", "utf-8"))
    msg.attach(alt)
    with open(result["path"], "rb") as f:
        part = MIMEApplication(f.read(), Name=os.path.basename(result["path"]))
    part["Content-Disposition"] = f'attachment; filename="{os.path.basename(result["path"])}"'
    msg.attach(part)
    body = {"message": {"raw": base64.urlsafe_b64encode(msg.as_bytes()).decode()}}

    draft = None
    if existing:
        try:
            draft = service.users().drafts().update(userId="me", id=existing, body=body).execute()
        except Exception:
            draft = None
    if draft is None:
        draft = service.users().drafts().create(userId="me", body=body).execute()
    for part_kind, ids in result["ids_by_kind"].items():
        engine.append_mail_log(ids, date.today().isoformat(), "drafted", part_kind)
    save_kstate(kind, draft_id=draft["id"], drafted_at=datetime.now().isoformat(timespec="seconds"),
                drafted_for=ks.get("data_date"), drafted_count=result["all"])
    return {"draft_id": draft["id"], "url": "https://mail.google.com/mail/u/0/#drafts"}


def mark_sent(kind):
    """You confirm the mail went out: the orders in its latest Excel are logged as sent."""
    ids_by_kind = kstate(kind).get("ids_by_kind") or {}
    n = 0
    for part_kind, ids in ids_by_kind.items():
        engine.append_mail_log(ids, date.today().isoformat(), "sent", part_kind)
        n += len(ids)
    save_kstate(kind, sent_at=datetime.now().isoformat(timespec="seconds"))
    return n


# ------------------------------------------------------------------ automation
def _days_since(ts):
    return (datetime.now() - datetime.fromisoformat(ts)).days if ts else 10 ** 6


def auto_run(df, meta, cfg=None, force=False):
    """Runs on its own when the data updates (and from daily_run.py). Builds both mails and saves a Gmail draft
    for the ones that are due. One pass per data date. Nothing is sent. Returns one result per mail.

      draft_mode "daily" (default): both mails get a draft every data date, as long as the list is not empty
      draft_mode "new_only": potential when there are new orders; claims on new orders or every `followup_days` days
    """
    cfg = cfg or engine.load_config()
    data_date = f"{meta['as_of']:%Y-%m-%d}"
    now = lambda: datetime.now().isoformat(timespec="seconds")
    followup = cfg.get("followup_days", 7)
    daily = cfg.get("draft_mode", "daily") == "daily"
    connected = gmail_connected(cfg)
    results = []
    for kind in KINDS:
        if not force and kstate(kind).get("auto_done_for") == data_date:
            results.append({"kind": kind, "action": "skipped", "reason": f"data of {data_date} already handled"})
            continue
        r = build(kind, df, meta, cfg)
        short = {"kind": kind, "new": r["new"], "all": r["all"], "path": r["path"]}
        last = max(filter(None, [kstate(kind).get("drafted_at"), kstate(kind).get("sent_at")]), default=None)
        if daily or force:        # a fresh draft every data date, as long as there is something to send
            due = r["all"] > 0
        else:                     # "new_only": only when there are new orders (claims mail also as a weekly reminder)
            due = r["new"] > 0 or (kind == "claims" and r["all"] > 0 and _days_since(last) >= followup)
        if not due:
            save_kstate(kind, auto_done_for=data_date, auto_result="nothing due", auto_at=now())
            results.append({**short, "action": "built", "reason": "nothing new, no draft needed"})
            continue
        if not connected:
            save_kstate(kind, auto_result="Excel ready, Gmail not connected", auto_at=now())
            results.append({**short, "action": "built", "reason": "Gmail not connected"})
            continue
        try:
            d = create_draft(r, cfg)
        except Exception as e:
            save_kstate(kind, auto_result=f"draft failed: {e}", auto_at=now())
            results.append({**short, "action": "built", "reason": f"draft failed: {e}"})
            continue
        save_kstate(kind, auto_done_for=data_date, auto_result="draft created", auto_at=now())
        results.append({**short, "action": "drafted", **d})
    return results
