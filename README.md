# Payment Verifier (Tampermonkey + Google Sheets)

Employee uploads/pastes a payment screenshot -> details saved in a Google Sheet.
Admin uploads the bank statement (Excel / CSV / PDF) -> every unverified row is checked
(UTR + amount + date). All 3 match = **Verified** (green), otherwise **Not Verified** (red).

## Setup (once, ~10 min)
### 1. Google Sheet + backend
1. Create a new Google Sheet.
2. Extensions -> Apps Script -> delete default code -> paste `Code.gs`.
3. Change `SECRET_TOKEN` (access key for employees) and `ADMIN_PASSWORD`.
4. Deploy -> New deployment -> Web app -> Execute as **Me**, Who has access **Anyone** -> Deploy -> authorise.
5. Copy the **Web app URL** (ends with `/exec`).

### 2. GitHub
1. New repo `payment-verifier`; upload `payment-verifier.user.js`, `Code.gs`, `README.md`.
2. In `payment-verifier.user.js` replace both `YOUR_GITHUB_USERNAME` (lines @namespace/@updateURL/@downloadURL).
3. Do NOT put your real access key / password in the repo (the script asks for the key in Settings).

### 3. Install
1. Install Tampermonkey (Chrome / Edge / Brave / Firefox / Safari).
2. Open `https://raw.githubusercontent.com/<you>/payment-verifier/main/payment-verifier.user.js` -> Install.
3. Chrome/Edge/Brave: enable *Developer mode* (or "Allow user scripts") for Tampermonkey if it asks.
4. On any website click the blue **Payment Verifier** button -> Settings -> paste Web App URL + Access key -> Save.

## Use
- Employee: enter email, paste (Ctrl/Cmd+V) or choose screenshot, check the auto-read fields, Save.
- Admin: enter admin password, choose statement, press Verify. Results are written to the sheet.

## Notes
- Employees never see the sheet data; the admin password is checked on the server.
- Duplicate UTRs are rejected.
- Date must equal the bank date (`DATE_TOLERANCE_DAYS` at top of the script: set 1 to allow +/-1 day).
- Scanned (image-only) PDFs are not supported – use Excel/CSV.
- If OCR is blocked on a site with strict security (e.g. github.com), use the tool from another site.
