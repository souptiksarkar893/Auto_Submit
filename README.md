# Auto Submit Tool (Excel -> Fast Index Panel)

Python automation utility that:

1. Reads links from Excel column A.
2. Removes blanks, invalid URLs, and duplicates.
3. Logs in to `https://fast-index.icu/botfarms/panel.php`.
4. Submits links into the textarea labeled "Paste Links (one per line / max 7000)".
5. Handles batch submission when links exceed the configured batch size (500 by default).
6. Optionally clears or moves submitted links inside the Excel file.

## Files

- `submit_links.py`: Main automation script.
- `.env.example`: Environment template.
- `requirements.txt`: Python dependencies.

## Setup

1. Create a virtual environment and install dependencies:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

2. Create `.env` from `.env.example` and fill values:

- `EXCEL_FILE_PATH`: Local Excel file path (example: `links.xlsx`).
- `EXCEL_SHEET_NAME`: Worksheet name to read from (example: `Sheet1`).
- `HISTORY_URL_BASE`: Base history endpoint (default: check_all.php endpoint).
- `PANEL_USERNAME`, `PANEL_PASSWORD`: Panel login credentials.

## Run

```powershell
python submit_links.py
```

## Runtime Behavior

- Logs like:
  - `History URL: ...`
  - `Fetched N links from Excel`
  - `Logging in...`
  - `Submitting batch 1/X...`
  - `Batch X submitted successfully`
- Cookie reuse:
  - Saves cookies to `COOKIES_FILE` after login.
  - Reuses cookies on next run if valid.
- Retries:
  - Retries login/submission on temporary failures.

## Config Highlights (`.env`)

- `HEADLESS=true|false`: Run without browser UI or with visible browser.
- `MAX_LINKS_PER_BATCH=500`: Batch size (default in this project).
- `POST_SUBMIT_ACTION=none|clear|move`:
  - `none`: keep original links in source sheet.
  - `clear`: clear submitted rows in source column A.
  - `move`: move links to `SUBMITTED_SHEET_NAME` (column A) and clear source rows.
- `RUN_INTERVAL_MINUTES=0`: Set `>0` to run with built-in scheduler.

## Notes

- URL validation accepts only `http://` and `https://` links.
- Duplicates are removed before submission.
- Excel files should be `.xlsx`/`.xlsm` format (openpyxl-compatible).
- If panel HTML changes, selectors in `submit_links.py` may need updates.
