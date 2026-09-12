#!/usr/bin/env python3
"""
sync_contractor_names.py
--------------------------
Compares the unique customer_name / contractor_name values found across:

  - base pipeline data    -> transaction_mapping_base.customer_name in DuckDB
  - credit pipeline data  -> transaction_mapping_credit.customer_name in DuckDB
  - template files        -> contractor_name column, "REPORTED SALES DATA"
                              sheet, header in row 6, read fresh from disk
                              (no pipeline/DuckDB table exists for these yet)

against the master contractors list (Google Sheet, "contractors" tab,
"original_contractor_name" column, first column). Any value not found there
(exact match — case, whitespace, and special characters all matter) gets
appended to the bottom of column A, with a note in column E:
"added YYYY-MM-DD from <source> file data" naming which pipeline(s)
contributed that value (e.g. "base/credit" if it showed up in both).

Base and credit values come straight from DuckDB — this script does NOT
re-read those Excel files or mapping docs. That means it's only as current
as your last run of base_mapping_qc.py / credit_mapping_qc.py; run those
first if you want this month's data included.

Requires oauth_desktop_app.json in the same directory as this script
(same OAuth client used by the other pipeline scripts). A separate
token_contractors.pkl is used since this needs read+write Sheets access,
unlike the read-only scopes those scripts use.

Usage:
    python sync_contractor_names.py
"""

import pickle
from collections import defaultdict
from datetime import date
from pathlib import Path

import duckdb
import gspread
import pandas as pd
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow

# -----------------------------------------------------------------------
# Silence invalid font family values (e.g. 34) that some Template files
# contain — same patch used elsewhere in this pipeline, needed here since
# pandas.read_excel uses openpyxl under the hood for .xlsx files.
# -----------------------------------------------------------------------
from openpyxl.descriptors.base import Min

_original_min_set = Min.__set__


def _patched_min_set(self, instance, value):
    try:
        _original_min_set(self, instance, value)
    except ValueError:
        pass


Min.__set__ = _patched_min_set

# =============================================================================
# CONFIGURATION
# =============================================================================
DUCKDB_PATH = "/Users/lorimartella/Documents/gmatter/charlotte_pipe/charlotte_pipe.duckdb"

TEMPLATES_DIR        = Path("/Users/lorimartella/Documents/gmatter/charlotte_pipe/templates")
TEMPLATE_SHEET_NAME  = "REPORTED SALES DATA"
TEMPLATE_HEADER_ROW  = 6   # 1-indexed, as seen in Excel
TEMPLATE_COLUMN_NAME = "contractor_name"
TEMPLATE_EXTENSIONS  = {".xlsx", ".xls", ".xlsm", ".xlsb"}

CONTRACTORS_SHEET_ID            = "17muso1TB00rbBYH0hiMzaCx0k8y_3u0OKhdKkrvgok4"
CONTRACTORS_TAB_NAME            = "contractors"
CONTRACTORS_NAME_COLUMN_HEADER  = "original_contractor_name"
CONTRACTORS_NAME_COLUMN_LETTER  = "A"
CONTRACTORS_NOTES_COLUMN_LETTER = "E"

BASE_DIR = Path(__file__).parent
SCOPES   = ["https://www.googleapis.com/auth/spreadsheets"]  # read + write
# =============================================================================


# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------

def get_credentials():
    creds = None
    token_path = BASE_DIR / "token_contractors.pkl"
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
# Source 1 & 2: DuckDB (base + credit)
# ---------------------------------------------------------------------------

def get_duckdb_customer_names() -> dict[str, set[str]]:
    """Return {value: {'base', 'credit'}-subset} for every distinct, non-null
    customer_name found in transaction_mapping_base and
    transaction_mapping_credit. Values come through exactly as DuckDB stored
    them (leading/trailing whitespace trimmed at load time; nothing else
    altered)."""
    sources: dict[str, set[str]] = defaultdict(set)

    con = duckdb.connect(DUCKDB_PATH, read_only=True)

    for table, source_label in [
        ("transaction_mapping_base", "base"),
        ("transaction_mapping_credit", "credit"),
    ]:
        rows = con.execute(f"""
            SELECT DISTINCT customer_name FROM {table}
            WHERE customer_name IS NOT NULL
        """).fetchall()
        for (value,) in rows:
            sources[value].add(source_label)

    con.close()
    return sources


# ---------------------------------------------------------------------------
# Source 3: Template files (read fresh — no pipeline/DuckDB table exists yet)
# ---------------------------------------------------------------------------

def get_template_contractor_names() -> dict[str, set[str]]:
    """Return {value: {'template'}} for every non-blank contractor_name value
    found across all Template Excel files. Values are used exactly as read
    (no trimming), since Template files don't go through DuckDB at all."""
    sources: dict[str, set[str]] = defaultdict(set)

    if not TEMPLATES_DIR.exists():
        print(f"  WARNING: templates folder not found: {TEMPLATES_DIR}")
        return sources

    excel_files = sorted(
        f for ext in TEMPLATE_EXTENSIONS
        for f in TEMPLATES_DIR.glob(f"*{ext}")
    )

    header_index = TEMPLATE_HEADER_ROW - 1  # convert to 0-indexed

    for file_path in excel_files:
        try:
            xl = pd.ExcelFile(file_path)
        except Exception as e:
            print(f"  WARNING: could not open '{file_path.name}': {e}")
            continue

        target_sheet = None
        target_lower = TEMPLATE_SHEET_NAME.strip().lower()
        for s in xl.sheet_names:
            if s.strip().lower() == target_lower:
                target_sheet = s
                break

        if target_sheet is None:
            print(f"  (skip) '{file_path.name}': no '{TEMPLATE_SHEET_NAME}' sheet")
            continue

        try:
            df = pd.read_excel(file_path, sheet_name=target_sheet, header=header_index)
        except Exception as e:
            print(f"  WARNING: could not read '{target_sheet}' in '{file_path.name}': {e}")
            continue

        normalised_columns = {str(col).strip().lower(): col for col in df.columns}
        target_col = normalised_columns.get(TEMPLATE_COLUMN_NAME.strip().lower())
        if target_col is None:
            print(f"  (skip) '{file_path.name}': no '{TEMPLATE_COLUMN_NAME}' column")
            continue

        for raw_value in df[target_col]:
            if pd.isna(raw_value):
                continue
            value = str(raw_value)
            if not value.strip():
                continue
            sources[value].add("template")

    return sources


# ---------------------------------------------------------------------------
# Contractors master list
# ---------------------------------------------------------------------------

def get_existing_contractor_names(gc: gspread.Client):
    """Return (worksheet, existing_names_set, next_empty_row_number)."""
    spreadsheet = gc.open_by_key(CONTRACTORS_SHEET_ID)
    worksheet = spreadsheet.worksheet(CONTRACTORS_TAB_NAME)

    all_values = worksheet.get_all_values()
    if not all_values:
        raise ValueError(f"Tab '{CONTRACTORS_TAB_NAME}' appears to be empty.")

    header_row = all_values[0]
    try:
        col_index = header_row.index(CONTRACTORS_NAME_COLUMN_HEADER)
    except ValueError:
        raise ValueError(
            f"Column '{CONTRACTORS_NAME_COLUMN_HEADER}' not found in tab "
            f"'{CONTRACTORS_TAB_NAME}'. Headers found: {header_row}"
        )

    existing = set()
    last_row_with_data = 1  # header row
    for i, row in enumerate(all_values[1:], start=2):
        if col_index < len(row) and row[col_index] != "":
            existing.add(row[col_index])
            last_row_with_data = i

    next_row = last_row_with_data + 1
    return worksheet, existing, next_row


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    print("Gathering customer_name values from DuckDB (base + credit)...")
    combined: dict[str, set[str]] = defaultdict(set)

    for value, srcs in get_duckdb_customer_names().items():
        combined[value] |= srcs
    print(f"  {len(combined)} unique value(s) so far (base + credit)")

    print("\nGathering contractor_name values from Template files...")
    for value, srcs in get_template_contractor_names().items():
        combined[value] |= srcs
    print(f"  {len(combined)} unique value(s) total (base + credit + template)")

    print("\nAuthenticating with Google Sheets...")
    creds = get_credentials()
    gc = gspread.authorize(creds)

    print(f"Reading existing names from '{CONTRACTORS_TAB_NAME}' tab...")
    worksheet, existing_names, next_row = get_existing_contractor_names(gc)
    print(f"  {len(existing_names)} existing name(s) found")

    missing = sorted(value for value in combined if value not in existing_names)
    print(f"\n{len(missing)} new value(s) not found in contractors sheet")

    if not missing:
        print("Nothing to add — every value is already accounted for.")
        return

    today = date.today().isoformat()
    updates = []
    for offset, value in enumerate(missing):
        row_num = next_row + offset
        source_label = "/".join(sorted(combined[value]))
        note = f"added {today} from {source_label} file data"
        updates.append({
            "range": f"{CONTRACTORS_NAME_COLUMN_LETTER}{row_num}",
            "values": [[value]],
        })
        updates.append({
            "range": f"{CONTRACTORS_NOTES_COLUMN_LETTER}{row_num}",
            "values": [[note]],
        })

    print(f"Appending {len(missing)} new row(s) to '{CONTRACTORS_TAB_NAME}' "
          f"(rows {next_row}-{next_row + len(missing) - 1})...")

    # The sheet's grid may not have enough rows allocated yet to write into —
    # expand it first if needed (this is separate from how many rows
    # actually contain data).
    required_rows = next_row + len(missing) - 1
    if worksheet.row_count < required_rows:
        worksheet.add_rows(required_rows - worksheet.row_count)

    worksheet.batch_update(updates, value_input_option="USER_ENTERED")

    print("\nDone. New values added:")
    for value in missing:
        print(f"  + {value!r}  ({'/'.join(sorted(combined[value]))})")


if __name__ == "__main__":
    main()