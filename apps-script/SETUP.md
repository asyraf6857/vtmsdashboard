# VTMS VOR Secure Excel Upload — Setup

## Architecture
GitHub Pages provides the dashboard and an **Upload VOR Excel** button. The button opens a Google Apps Script HtmlService page (`?page=upload`). File bytes are submitted to Apps Script using `google.script.run`; the server verifies the signed-in Google account against an allowlist before importing. Do **not** implement an unauthenticated POST endpoint, place OAuth tokens in GitHub Pages, or publish raw vessel records.

## Required owner setup
1. Open **PTP VTMS MASTER DATABASE** > **Extensions > Apps Script**.
2. Replace the existing `Code.gs` with `apps-script/Code.gs` from this repository. Create a separate HTML file named exactly `Upload` and paste the contents of `apps-script/Upload.html` (do not include the .html extension in the Apps Script file name field).
3. Enable **Services > + > Drive API** (Advanced Google service; v3).
4. Under **Project Settings > Script properties**, add `VTMS_UPLOAD_ALLOWLIST` with authorized Google emails separated by commas, e.g. `operator1@company.example,operator2@company.example`. An empty or unknown identity is always denied.
5. Deploy **Web app** as **Execute as: User accessing the web app**. Restrict **Who has access** to your organization if your Google Workspace offers that option. Each operator needs spreadsheet access and permission to use Drive; each must authorize the script on first use.
6. **Deploy > Manage deployments > Edit > New version > Deploy**. The existing /exec URL generally remains unchanged when editing the deployment; verify it.
7. Open `<YOUR_EXEC_URL>?page=upload` while signed in as an allowlisted user and test with a backed-up sample VOR Excel file. Confirm `MONTHLY_SUMMARY`, `VISIT_REASON`, `LOA_ANALYSIS` and `UPLOAD_LOG`.
8. Open GitHub Pages and click **Upload VOR Excel**. Refresh the dashboard after a successful import.

## Important
- Only `.xlsx` files up to 10 MB are accepted.
- The importer expects the original Marine VOR workbook with a worksheet named `2025` or `2026` matching the year in the filename and known cell positions. If the source layout changes, import will fail or must be adapted.
- The importer **replaces the entire year's existing monthly aggregate rows** with months found in the new workbook; upload the latest full-year/YTD workbook, not a single-month-only workbook. Back up the master spreadsheet first.
- Data reads are currently exposed by the deployed `doGet` API as **aggregate statistics** via JSON/JSONP for the public GitHub Pages site. Obtain PTP authorization for publishing these aggregates. Upload authorization does not protect public aggregate reads.
- If Google account email is unavailable from Apps Script under your deployment mode, the upload is deliberately denied. Use a managed Workspace configuration and verify identity; never remove the server-side check to work around this.
- This code has been committed to GitHub but **cannot be activated remotely by a GitHub commit**. A Google Apps Script owner must copy files and redeploy.
