# shipbob_live: ShipBob D2C Live (module of the recon dashboard)

The ShipBob D2C order / claims dashboard and the ShipBob mail drafts, running **inside** the recon dashboard
at `http://localhost:5000`, tab **"ShipBob D2C Live"**. The page itself is served at `/shipbob/` and shown in
that tab (iframe). The older **"ShipBob D2C"** tab (reads the Google Sheet tab "ShipBob D2C Claims") is untouched.

## What it does
* Reads the order Excel `Dump` sheets from `G:\My Drive\Shibob D2C Claim\` (April-May file + the newest other `.xlsx`,
  the same rule as `upload_shipbob_d2c.py`). It reads the **Excel**, not the Google Sheet.
* Buckets (all from the Excel): Delivered / In Transit (0-15 days since Import Date) / Potential In-Transit Loss (day 11-15) /
  Under Dispute (>15 days, split by the **Sub Bucket** column: Pending to claim, Claim raised but not received,
  Claim raised and received, Claim Window Expired) / RTO / Cancelled. The "Claim status" column is not used.
* Expected $ for Pending / Raised / Expired = the Excel's **Expected Claim amount** column (same numbers as the recon
  "Claims Bucket - Year-to-Date" table); Received $ = **Amt**.
* Two ShipBob mails, saved as **Gmail drafts, never sent**: (1) Potential in-transit loss (11-15 days),
  (2) Claims follow-up (Pending to claim + Claim raised but not received in one mail). HTML body, Excel attached,
  no order values in the body. A watcher checks the Excel every 30 s and refreshes the page and the drafts when it changes.

## Files
| File | Purpose |
|---|---|
| `bp.py` | Flask blueprint (`/shipbob/...`), the watcher, downloads, mail actions |
| `engine.py` | Reads the Excel, applies the rules, builds all numbers and data checks |
| `mailer.py` | Mail Excel + email text, Gmail draft (updates the same-day draft, keeps To/Cc), mail log |
| `templates/shipbob_live.html` | The page (reads `/shipbob/api/...`; `__BASE__` / `__BUILD__` are filled in by the server) |
| `config.json` | Excel folder, day thresholds, mail To/Cc, `draft_mode` (daily / new_only). **Git-ignored** |
| `state/` `cache/` `output/` | Mail log + state, cached Excel reads, built mail Excels. **Git-ignored** |

## How it is wired into the recon app (3 small edits)
* `app.py`: registers the blueprint inside a `try/except` before `if __name__ == '__main__':`, and calls
  `start_shipbob_watcher()` right before `app.run(...)`. If anything fails it only prints `[ShipBob Live] disabled`.
* `templates/dashboard.html`: one tab button (`data-tab="shipbob_live"`), one branch in `applyAndRender()`,
  and `renderShipBobLive()` which puts the iframe in `#mainContent`.
* `.gitignore`: the four git-ignored paths above.
* The blueprint sets `X-Frame-Options: SAMEORIGIN` on its page (the app default is `DENY`, which would block the iframe).

## Rules to stay in sync
* **Local only.** It needs the G: drive and Gmail. If `source_folder` does not exist (e.g. Render) every route answers
  "switched off" and the watcher never starts. Do not add pandas to `requirements.txt` for it.
* **One server only.** Mail drafts are created by whichever process runs the watcher. The old standalone copy
  (`C:\Users\Admin\shipbob-d2c-dashboard`, port 5055) is retired: do not run it together with this module.
* **Do not revert or `git stash` the uncommitted edits** above while this module is in use; commit them when approved.
* The Google Sheet tab "ShipBob D2C Claims" is written by `upload_shipbob_d2c.py` (20 columns incl. Store Order ID,
  Order Status, Carrier, Tracking Information). This module does not read that tab, so both stay independent;
  both use the same Excel as the source.
* Gmail: one-time "Connect Gmail" in the page (scope `gmail.compose`); token at the path in `config.json`.

## Data source: Google Sheet (current)
* `config.json` -> `"data_source": "sheet"` (+ `sheet_id`). The module reads the tab **ShipBob D2C Claims** (read-only, by header name)
  and takes freshness from the tab **ShipBob D2C Sync Info** (A1:C1 = synced at, file name, row count). The watcher re-reads when that cell changes.
* Set `"data_source": "excel"` to go back to reading the G: drive Excel files (the code for both stays).
* Product prices are not in the sheet: they come from `shipbob_live/product_prices.csv` (copy of the Excel 'Product price' tab). New product -> add a row there.
* Columns the sheet does not carry (Purchase Date, Source, Fulfillment Center, Claim status) are not used by any number.
* Checked: sheet mode vs Excel mode give identical Main/Sub/Age/Expected/Received for all 65,576 shipments.
