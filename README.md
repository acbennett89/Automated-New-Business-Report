# New Biz Report Automation

Work in progress. This README will be updated as we build the new automation.

## Goal
Automate collecting the required reports (currently gathered in Chrome) and produce the final output consistently.

## Prerequisites
- Windows
- Google Chrome
- Access to the required report portals
- Download folder configured (confirm default download location)
- Python 3 (or permission to install it on first run)
- Internet access to install Playwright (first run)

## Current Manual Process (Chrome)
1. Open Chrome.
2. Log in to the required portals.
3. Download the required reports.

## How To Run
1. Double-click `run_reports.bat`.
2. On first run, the script installs Python (if missing), creates a local `.venv`, installs dependencies, then opens a login window so you can authenticate.
3. The script saves a Playwright `storage_state.json` in `config/` so future runs can be headless.
4. The script navigates to Agency Reports and downloads "Opportunities by Producer" into `Working Files/` (headless).
5. Keep the terminal open while it runs.

### Desktop UI
1. Run `run_reports.bat ui` to open a Tkinter desktop launcher.
2. Use the UI to run full pipeline, partial pipeline, setup only, and view live logs.
3. You can also double-click `Launch_UI.vbs` to launch the UI directly.
4. In the UI, enter EPIC `Usercode` and `Password` and click `Save Credentials` to enable EPIC auto-login.
5. In the UI, enter Bignition `Username` and `Password` and click `Save Credentials` to enable Bignition auto-login.

### New Biz Tabs Only (Debug)
1. Run `run_reports.bat newbiz` to build only the `YYYY New Biz` tabs from the existing consolidated workbook.
2. This does not pull reports; it only updates `Consolidated New Biz Report.xlsx`.

## EPIC (Testing)
1. Run `run_reports.bat epic` to launch the EPIC flow in a headed browser.
2. EPIC full flow now runs headless by default and uses saved credentials from `config/epic_credentials.json` when available.
3. If you need visible browser debugging, run `.\.venv\Scripts\python.exe Scripts\epic_report.py --headed`.
4. The script saves `config/epic_storage_state.json` for convenience.

## Report Sources
1. Bignition portal login: `https://bignition.thewedge.net/Account/Login?ReturnUrl=%2f`
2. Dashboard: `https://bignition.thewedge.net/Home/Dashboard`
3. Agency Reports: `https://bignition.thewedge.net/Reports/AgencyReports`
4. Opportunities by Producer: `https://bignition.thewedge.net/Reports/ViewAgencyReport?id=5`

## Open Questions
1. Which reports are required (names, portals, URLs)?
2. What date range or filters are needed for each report?
3. Where should downloads be saved?
4. What is the desired final output and file naming convention?

## Notes
- Add steps here as we confirm them.
- Downloads are saved to `Working Files/`.
- Delete `config/storage_state.json` if you need to force a fresh login.
- Delete `config/epic_storage_state.json` to force EPIC re-login.
- If Python is missing, `run_reports.bat` attempts to install it via `winget` and may prompt for approval.

## Change Log
- 2026-02-17: Initialized README.
- 2026-02-17: Added `main.py` and `run_reports.bat` to launch the first login screen.
- 2026-02-17: Added Playwright automation for Bignition navigation and download using `storage_state.json`.
- 2026-02-17: Updated `run_reports.bat` to self-install dependencies on first run.
