#!/usr/bin/env python3
"""
file_counts_report.py
------------------------
Builds a per-file record-count report across:

  - base mapping doc     -> Google Sheet, 'info' tab, rows_count summed per
                             file_name, counting only rows where process=TRUE
  - credit mapping doc   -> same, for the credit mapping doc
  - template xlsx files  -> row count calculated by reading each file's
                             "REPORTED SALES DATA" sheet (header in row 6),
                             counting rows below the header with at least 3
                             populated columns (excludes bad-data rows that
                             have only 1-2 stray values)

Writes the combined result to a brand-new Google Sheet named
"file_counts_YYYYMMDD" (today's date), created in the same Drive folder that
actually contains the base/credit mapping docs (verified via each doc's real
parent folder, not just the month folder search result).

Columns written: file_name, count_of_records, file_type, purpose

Requires oauth_desktop_app.json in the same directory as this script (needs
full spreadsheets + drive scopes, since it creates and moves a new file —
broader than the read-only scopes the monthly pipeline scripts use, so a
separate token_file_counts.pkl is used).

Usage:
    python file_counts_report.py
"""

import pickle
from collections import defaultdict
from datetime import date
from pathlib import Path
from typing import Dict

import gspread
import pandas as pd
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

from pipeline_helpers import find_month_folder, find_mapping_sheet_id

# =============================================================================
# CONFIGURATION
# =============================================================================
MAPPING_ROOT_FOLDER_ID = "1pkSdlq7gYLPvaJIWHCjI7kpmOhTcJsEt"
BASE_INCLUDE_KEYWORD   = "unspecified"
BASE_EXCLUDE_KEYWORD   = "credit"
CREDIT_INCLUDE_KEYWORD = "credit"
CREDIT_EXCLUDE_KEYWORD = None

MAPPING_TAB = "info"

TEMPLATES_DIR       = Path("/Users/lorimartella/Documents/gmatter/charlotte_pipe/templates")
TEMPLATE_SHEET_NAME = "REPORTED SALES DATA"
TEMPLATE_HEADER_ROW = 6   # 1-indexed, as seen in Excel
TEMPLATE_EXTENSIONS = {".xlsx", ".xls", ".xlsm", ".xlsb"}

BASE_DIR = Path(__file__).parent
# Full (not read-only) scopes: this script creates a new file and moves it
# into a folder, which read-only access can't do.
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
# =============================================================================


# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------

def get_credentials():
    creds = None
    token_path = BASE_DIR / "token_file_counts.pkl"
    creds_path = BASE_DIR / "oauth_desktop_app.json"

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


# ---------------------------------------------------------------------------
# Base / credit mapping doc counts
# ---------------------------------------------------------------------------

def get_mapping_doc_counts(gc: gspread.Client, sheet_id: str) -> Dict[str, int]:
    """Read the 'info' tab and return {file_name: summed rows_count},
    counting only rows where process=TRUE — matching what actually gets
    loaded by base_mapping_qc.py / credit_mapping_qc.py."""
    spreadsheet = gc.open_by_key(sheet_id)
    worksheet = spreadsheet.worksheet(MAPPING_TAB)
    records = worksheet.get_all_records()  # list of dicts keyed by header row

    totals: Dict[str, int] = defaultdict(int)
    for row in records:
        process = row.get("process", "")
        if isinstance(process, bool):
            process_str = "TRUE" if process else "FALSE"
        else:
            process_str = str(process).strip().upper()
        if process_str not in ("TRUE", "1", "YES", "Y"):
            continue

        file_name = str(row.get("file_name", "")).strip()
        if not file_name:
            continue

        raw_count = str(row.get("rows_count", "")).strip().replace(",", "")
        try:
            count = int(raw_count) if raw_count else 0
        except ValueError:
            print(f"    WARNING: non-numeric rows_count {raw_count!r} for "
                  f"'{file_name}' — treating as 0")
            count = 0

        totals[file_name] += count

    return dict(totals)


# ---------------------------------------------------------------------------
# Template file counts (read fresh — no mapping doc / DuckDB table for these)
# ---------------------------------------------------------------------------

def get_template_counts() -> Dict[str, int]:
    """Return {file_name: complete data row count} for every Template file
    that has a 'REPORTED SALES DATA' sheet, counting rows below the row-6
    header that have at least 3 populated columns (a row with only 1-2
    stray populated cells is treated as bad data, not a real record)."""
    counts: Dict[str, int] = {}

    if not TEMPLATES_DIR.exists():
        print(f"    WARNING: templates folder not found: {TEMPLATES_DIR}")
        return counts

    excel_files = sorted(
        f for ext in TEMPLATE_EXTENSIONS
        for f in TEMPLATES_DIR.glob(f"*{ext}")
    )

    header_index = TEMPLATE_HEADER_ROW - 1  # convert to 0-indexed
    target_lower = TEMPLATE_SHEET_NAME.strip().lower()

    for file_path in excel_files:
        try:
            xl = pd.ExcelFile(file_path)
        except Exception as e:
            print(f"    WARNING: could not open '{file_path.name}': {e}")
            continue

        target_sheet = None
        for s in xl.sheet_names:
            if s.strip().lower() == target_lower:
                target_sheet = s
                break

        if target_sheet is None:
            print(f"    (skip) '{file_path.name}': no '{TEMPLATE_SHEET_NAME}' sheet")
            continue

        try:
            df = pd.read_excel(file_path, sheet_name=target_sheet, header=header_index)
        except Exception as e:
            print(f"    WARNING: could not read '{target_sheet}' in '{file_path.name}': {e}")
            continue

        # Only count rows with meaningfully complete data — a row with just
        # one stray populated cell (bad data) shouldn't count as a record.
        # Require at least 3 populated columns.
        populated_counts = df.notna().sum(axis=1)
        complete_rows = df[populated_counts >= 3]
        counts[file_path.name] = len(complete_rows)

    return counts


# ---------------------------------------------------------------------------
# Drive folder / file helpers
# ---------------------------------------------------------------------------

def get_containing_folder_id(drive_service, file_id: str) -> str:
    """Return the actual immediate parent folder ID of a Drive file."""
    meta = drive_service.files().get(
        fileId=file_id, fields="parents", supportsAllDrives=True
    ).execute()
    parents = meta.get("parents", [])
    if not parents:
        raise ValueError(f"File {file_id!r} has no parent folder on Drive.")
    return parents[0]


def create_report_sheet(gc: gspread.Client, drive_service, folder_id: str,
                         title: str) -> gspread.Spreadsheet:
    """Create a new Google Sheet and move it into folder_id."""
    spreadsheet = gc.create(title)

    file_meta = drive_service.files().get(
        fileId=spreadsheet.id, fields="parents", supportsAllDrives=True
    ).execute()
    previous_parents = ",".join(file_meta.get("parents", []))

    drive_service.files().update(
        fileId=spreadsheet.id,
        addParents=folder_id,
        removeParents=previous_parents,
        fields="id, parents",
        supportsAllDrives=True,
    ).execute()

    return spreadsheet


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    print("Authenticating with Google Sheets/Drive...")
    creds = get_credentials()
    gc = gspread.authorize(creds)
    drive_service = build("drive", "v3", credentials=creds)

    print("\nLocating this month's mapping folder...")
    folder_id, folder_name = find_month_folder(drive_service, MAPPING_ROOT_FOLDER_ID)

    print("\nLocating base mapping doc...")
    base_sheet_id, base_sheet_name = find_mapping_sheet_id(
        drive_service, folder_id,
        include_keyword=BASE_INCLUDE_KEYWORD, exclude_keyword=BASE_EXCLUDE_KEYWORD,
    )
    print(f"  Found: '{base_sheet_name}' ({base_sheet_id})")

    print("\nLocating credit mapping doc...")
    credit_sheet_id, credit_sheet_name = find_mapping_sheet_id(
        drive_service, folder_id,
        include_keyword=CREDIT_INCLUDE_KEYWORD, exclude_keyword=CREDIT_EXCLUDE_KEYWORD,
    )
    print(f"  Found: '{credit_sheet_name}' ({credit_sheet_id})")

    print("\nDetermining containing folder for the new report...")
    base_parent = get_containing_folder_id(drive_service, base_sheet_id)
    credit_parent = get_containing_folder_id(drive_service, credit_sheet_id)
    if base_parent != credit_parent:
        print(f"  WARNING: base doc's folder ({base_parent}) and credit doc's "
              f"folder ({credit_parent}) differ — using base doc's folder.")
    target_folder_id = base_parent
    print(f"  Target folder: {target_folder_id}")

    print("\nReading base mapping doc row counts...")
    base_counts = get_mapping_doc_counts(gc, base_sheet_id)
    print(f"  {len(base_counts)} file(s), {sum(base_counts.values()):,} total record(s)")

    print("\nReading credit mapping doc row counts...")
    credit_counts = get_mapping_doc_counts(gc, credit_sheet_id)
    print(f"  {len(credit_counts)} file(s), {sum(credit_counts.values()):,} total record(s)")

    print("\nReading template file row counts...")
    template_counts = get_template_counts()
    print(f"  {len(template_counts)} file(s), {sum(template_counts.values()):,} total record(s)")

    rows = [["file_name", "count_of_records", "file_type", "purpose"]]

    for file_name, count in sorted(base_counts.items()):
        rows.append([file_name, count, "unspecified", "contractor_base"])

    for file_name, count in sorted(credit_counts.items()):
        rows.append([file_name, count, "unspecified", "contractor_credit"])

    for file_name, count in sorted(template_counts.items()):
        rows.append([file_name, count, "templated_api", "contractor_base"])

    title = f"file_counts_{date.today().strftime('%Y%m%d')}"
    print(f"\nCreating new sheet '{title}' in target folder...")
    spreadsheet = create_report_sheet(gc, drive_service, target_folder_id, title)
    worksheet = spreadsheet.sheet1
    worksheet.update("A1", rows, value_input_option="USER_ENTERED")

    print(f"\nDone. {len(rows) - 1} row(s) written.")
    print(f"View: https://docs.google.com/spreadsheets/d/{spreadsheet.id}")


if __name__ == "__main__":
    main()