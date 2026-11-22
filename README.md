# ahg-asset-scanner-app

Phone QR scanner for asset inventory rounds at Almoosa. Scan → pick location → the Google Sheet updates.

| Folder | What |
|---|---|
| `frontend/index.html` | The scanner page (GitHub Pages). One file, no build. |
| `backend-gs/Code.gs` | Apps Script bound to the inventory sheet, deployed as a web app. |

## Apps Script deploys itself
Push a change under `backend-gs/` → the **Push to Apps Script** Action runs `clasp push`, then moves the web app
deployment to the new version (same `/exec` URL). Re-run it by hand from the Actions tab.
- `backend-gs/.clasp.json` – script ID of the Apps Script bound to the asset sheet.
- `backend-gs/appsscript.json` – manifest (web app: execute as me, anyone).
- Repo secrets: `CLASPRC_JSON` (clasp login of the Google account that owns the sheet), `SCANNER_TOKEN` (replaces `CHANGE_ME` at push time).

## Sheet tabs
- **Asset tab** (`ASSET_SHEET` in Code.gs, now `Sheet1`) – the asset list. Any columns. `Found`, `Location`, `Scanned At`, `Scanned By` are added if missing.
- **Settings** – Key / Value. Controls the app; edit here, no code change:
  - `Title`, `Locations` (comma list), `Match Columns` (columns the QR is matched against),
  - `Card Columns` (shown after a scan), `Group Column` (filter chips), `Code Pattern` (regex that pulls the code out of the QR text).
- **Scans** – every scan, append-only. Scans not on the list have `On List = No`.
- **Staff** – names. When filled, the app shows a one-tap name picker at setup.

## Add a new kind of asset
Add the rows to the asset tab with the same column names. If its asset code looks different, add it to `Code Pattern` (e.g. `…|EQ-\d{5}`). The `Group Column` chips pick it up automatically.

## Setup link (one tap per phone)
`https://<pages-url>/#api=<exec URL>&t=<token>&n=<name>`
Saved on the phone, then removed from the address bar.

## Later
- Login (Staff tab + PIN).
