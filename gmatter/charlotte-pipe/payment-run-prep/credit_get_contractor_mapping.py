"""
Google Drive Scanner → Customer Name Mapping
----------------------------------------------
Scans a Google Drive folder (and all subfolders) for Google Sheets whose name
contains a configurable keyword. From each matching sheet's "info" tab it
extracts the mapping columns needed to later locate the "customer_name" column
in the corresponding Excel files:

    file_name | sheet_name | customer_name | header_range | data_range | static_value

If the value in the "customer_name" column is wrapped in $$ (e.g. $$Acme Corp$$),
the literal text between the markers is written directly into the "static_value"
column and no Excel lookup is needed for that row.

Results are written to a tab called "CustomerNameMapping" in the output sheet.

Requirements:
    pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib gspread

Auth setup:
    1. Go to https://console.cloud.google.com/
    2. Enable the Google Drive API and Google Sheets API
    3. Create OAuth 2.0 credentials (Desktop app) and download as credentials.json
    4. Place credentials.json in the same directory as this script
"""

import pickle
import re
from pathlib import Path

import gspread
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# ─────────────────────────────────────────────
#  CONFIGURATION  ← edit these values
# ─────────────────────────────────────────────

# Google Drive folder to scan (root + all subfolders are searched).
# Found in the folder URL: drive.google.com/drive/folders/<FOLDER_ID>
SOURCE_FOLDER_ID = "1vuyxzhvo0NB9nX4-Myyj9NPZoBuzBVAA"

# Google Sheet to write mapping results into.
# Found in the sheet URL: docs.google.com/spreadsheets/d/<SHEET_ID>
OUTPUT_SHEET_ID = "1LUKPMtd41o-0uvMXb6rk_tY1HIqcme3zkAAaMh2yF-Y"

# Only Google Sheets whose name contains this string will be processed.
SEARCH_KEYWORD = "contractor_credit_unspecified"

# Tab name written in the output sheet (created if it doesn't exist).
OUTPUT_TAB_NAME = "credit_CustomerNameMapping"

# ─────────────────────────────────────────────

BASE_DIR = Path(__file__).parent

SCOPES = [
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/spreadsheets",
]


# ── Auth ──────────────────────────────────────────────────────────────────────

def get_credentials():
    """Return valid OAuth credentials, refreshing or re-authorising as needed."""
    creds = None
    token_path = BASE_DIR / "token.pkl"
    creds_path = BASE_DIR / "credentials.json"

    if token_path.exists():
        with open(token_path, "rb") as f:
            creds = pickle.load(f)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(creds_path), SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_path, "wb") as f:
            pickle.dump(creds, f)

    return creds


# ── Drive helpers ─────────────────────────────────────────────────────────────

def get_all_folder_ids(drive_service, root_folder_id: str) -> list[str]:
    """Recursively collect the root folder ID and every subfolder ID beneath it."""
    all_ids = [root_folder_id]
    queue = [root_folder_id]

    while queue:
        parent_id = queue.pop()
        page_token = None
        while True:
            resp = drive_service.files().list(
                q=(
                    f"'{parent_id}' in parents"
                    " and mimeType = 'application/vnd.google-apps.folder'"
                    " and trashed = false"
                ),
                fields="nextPageToken, files(id, name)",
                pageToken=page_token,
                includeItemsFromAllDrives=True,
                supportsAllDrives=True,
            ).execute()

            for folder in resp.get("files", []):
                print(f"  📁 Found subfolder: {folder['name']}")
                all_ids.append(folder["id"])
                queue.append(folder["id"])

            page_token = resp.get("nextPageToken")
            if not page_token:
                break

    print(f"Scanning {len(all_ids)} folder(s) total (root + subfolders).")
    return all_ids


def find_matching_sheets(drive_service, folder_id: str, keyword: str) -> list[dict]:
    """Return all Google Sheets inside folder_id (recursive) whose name contains keyword."""
    folder_ids = get_all_folder_ids(drive_service, folder_id)

    results = []
    batch_size = 50

    for i in range(0, len(folder_ids), batch_size):
        batch = folder_ids[i : i + batch_size]
        parents_clause = " or ".join(f"'{fid}' in parents" for fid in batch)
        query = (
            f"({parents_clause})"
            f" and name contains '{keyword}'"
            " and mimeType = 'application/vnd.google-apps.spreadsheet'"
            " and trashed = false"
        )
        page_token = None
        while True:
            resp = drive_service.files().list(
                q=query,
                fields="nextPageToken, files(id, name)",
                pageToken=page_token,
                includeItemsFromAllDrives=True,
                supportsAllDrives=True,
            ).execute()

            results.extend(resp.get("files", []))
            page_token = resp.get("nextPageToken")
            if not page_token:
                break

    print(f"Found {len(results)} matching file(s).")
    return results


# ── Static-value helper ───────────────────────────────────────────────────────

def extract_static_value(raw: str) -> str | None:
    """
    If raw is wrapped in $$ markers (e.g. '$$Acme Corp$$'), return the text
    between the markers. Otherwise return None.
    """
    match = re.fullmatch(r"\$\$(.+?)\$\$", raw.strip())
    return match.group(1).strip() if match else None


# ── Mapping extraction ────────────────────────────────────────────────────────

def extract_customer_name_mapping(gc: gspread.Client, file: dict) -> list[list]:
    """
    Open the Google Sheet and pull mapping rows from its 'info' tab.
    Only rows that have a value in the 'customer_name' column are kept.

    If the customer_name value is wrapped in $$ (e.g. $$Acme Corp$$), the
    unwrapped text is stored in the 'static_value' column and the
    'customer_name' column is left empty — no Excel lookup will be needed.

    Returns a list of rows:
        [source_file, file_name, sheet_name, customer_name, header_range, data_range, static_value]
    """
    rows = []
    try:
        spreadsheet = gc.open_by_key(file["id"])

        try:
            worksheet = spreadsheet.worksheet("info")
        except gspread.exceptions.WorksheetNotFound:
            print(f"    ℹ  No 'info' tab in '{file['name']}', skipping.")
            return rows

        data = worksheet.get_all_records()
        if data:
            print(f"    Columns found: {list(data[0].keys())}")

        for record in data:
            # Normalise keys to lowercase so column name capitalisation doesn't matter
            record_lower = {k.lower().strip(): v for k, v in record.items()}

            customer_name_raw = str(record_lower.get("customer_name", "")).strip()

            # Skip rows that have no customer_name mapping at all
            if not customer_name_raw:
                continue

            file_name_val    = str(record_lower.get("file_name", "")).strip()
            sheet_name_val   = str(record_lower.get("sheet_name", "")).strip()
            header_range_val = str(record_lower.get("header_range", "")).strip()
            data_range_val   = str(record_lower.get("data_range", "")).strip()

            # Check for $$literal value$$ — if found, store it and clear the
            # column identifier so Script 2 skips the Excel lookup entirely.
            static_val = extract_static_value(customer_name_raw)
            if static_val:
                customer_name_col = ""   # no column to look up
                print(f"    ℹ  Static value detected for '{file_name_val}': '{static_val}'")
            else:
                customer_name_col = customer_name_raw
                static_val = ""

            row = [
                file["name"],       # source Google Sheet name
                file_name_val,      # Excel file name
                sheet_name_val,     # Excel sheet/tab name
                customer_name_col,  # column identifier (empty when static_value is set)
                header_range_val,   # e.g. "A1:Z1"
                data_range_val,     # e.g. "A2:Z1000"
                static_val,         # literal value (empty when Excel lookup is needed)
            ]
            rows.append(row)

        static_count = sum(1 for r in rows if r[6])
        lookup_count = len(rows) - static_count
        print(f"    → {len(rows)} row(s): {static_count} static, {lookup_count} Excel lookup")

    except Exception as e:
        print(f"  ⚠  Could not read '{file['name']}': {e}")

    return rows


# ── Output ────────────────────────────────────────────────────────────────────

def write_mapping_to_sheet(gc: gspread.Client, all_rows: list[list]):
    """Write mapping rows (with header) to OUTPUT_TAB_NAME in OUTPUT_SHEET_ID."""
    spreadsheet = gc.open_by_key(OUTPUT_SHEET_ID)

    try:
        ws = spreadsheet.worksheet(OUTPUT_TAB_NAME)
        ws.clear()
    except gspread.exceptions.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=OUTPUT_TAB_NAME, rows=1, cols=7)

    header = ["source_file", "file_name", "sheet_name", "customer_name", "header_range", "data_range", "static_value"]
    ws.update([header] + all_rows)
    print(f"✓ Wrote {len(all_rows)} mapping row(s) to '{OUTPUT_TAB_NAME}' tab.")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    creds = get_credentials()
    drive_service = build("drive", "v3", credentials=creds)
    gc = gspread.authorize(creds)

    # 1. Find matching Google Sheets in the source folder
    matching_files = find_matching_sheets(drive_service, SOURCE_FOLDER_ID, SEARCH_KEYWORD)

    if not matching_files:
        print("No matching files found. Exiting.")
        return

    # 2. Extract customer_name mapping from each file's 'info' tab
    all_rows: list[list] = []
    for f in matching_files:
        print(f"  Reading: {f['name']}")
        rows = extract_customer_name_mapping(gc, f)
        all_rows.extend(rows)

    # 3. Write results to the output sheet
    if all_rows:
        write_mapping_to_sheet(gc, all_rows)
    else:
        print("No customer_name mapping rows found across all matching files.")


if __name__ == "__main__":
    main()
