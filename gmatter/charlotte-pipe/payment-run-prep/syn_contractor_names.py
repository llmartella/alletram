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
                              — PLUS the contractor name embedded in each
                              Template file's own name (position 6, 0-indexed,
                              splitting on "_")

against the master contractors list (Google Sheet, "contractors" tab,
"original_contractor_name" column, first column). Any value not found there
(exact match — case, whitespace, and special characters all matter) gets
appended to the bottom of column A, with a note in column E:
"added YYYY-MM-DD from <source> file data (<file names>)" naming which
pipeline(s) contributed that value (e.g. "base/credit" if it showed up in
both) and every source file it was actually found in.

Base and credit values come straight from DuckDB — this script does NOT
re-read those Excel files or mapping docs. That means it's only as current
as your last run of base_mapping_qc.py / credit_mapping_qc.py; run those
first if you want this month's data included.

Optionally limits base/credit data to files processed in roughly the last
N days (see LOOKBACK_DAYS below) — comment that line out entirely to
include all historical data instead, with no limit.

Requires oauth_desktop_app.json in the same directory as this script
(same OAuth client used by the other pipeline scripts). A separate
token_contractors.pkl is used since this needs read+write Sheets access,
unlike the read-only scopes those scripts use.

Usage:
    python sync_contractor_names.py
"""

import pickle
import re
from collections import defaultdict
from datetime import date, datetime, timedelta
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

# Only include base/credit rows from files processed in roughly the last N
# days (based on the trailing date embedded in each file's name — see
# extract_processed_date() below). To include ALL historical data instead,
# with no date limit, just comment out this whole line.
LOOKBACK_DAYS = 90

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

# Charlotte Pipe file names end with one or more standalone 8-digit
# (YYYYMMDD) tokens — typically a payment period start/end and a processed
# date, e.g. "..._20260401_20260630_20260910_charlotte.xlsx". The lookaround
# pattern below only matches a token that isn't itself part of a longer run
# of digits (so it doesn't false-match inside a 10-digit vendor number, etc).
_DATE_TOKEN_PATTERN = re.compile(r"(?<!\d)(\d{8})(?!\d)")


def extract_processed_date(archive_file_name: str):
    """Best-effort: pull the LAST standalone 8-digit token out of a source
    file's name and parse it as the date it was processed. Returns None if
    no such token is found or it doesn't parse as a real calendar date —
    callers should treat None as "unknown age" rather than "too old"."""
    matches = _DATE_TOKEN_PATTERN.findall(archive_file_name or "")
    if not matches:
        return None
    try:
        return datetime.strptime(matches[-1], "%Y%m%d").date()
    except ValueError:
        return None


def get_duckdb_customer_names(lookback_days) -> dict[str, dict]:
    """Return {value: {"sources": {'base','credit'}-subset, "files": {file
    names}}} for every distinct, non-null customer_name found in
    transaction_mapping_base and transaction_mapping_credit. Values come
    through exactly as DuckDB stored them (leading/trailing whitespace
    trimmed at load time; nothing else altered).

    If lookback_days is not None, only rows from files whose embedded
    processed-date is within the last lookback_days days are included. Rows
    whose date can't be determined are included regardless, rather than
    risking silently dropping current data over a naming quirk."""
    result: dict[str, dict] = defaultdict(lambda: {"sources": set(), "files": set()})
    cutoff = date.today() - timedelta(days=lookback_days) if lookback_days is not None else None

    con = duckdb.connect(DUCKDB_PATH, read_only=True)

    for table, source_label in [
        ("transaction_mapping_base", "base"),
        ("transaction_mapping_credit", "credit"),
    ]:
        rows = con.execute(f"""
            SELECT DISTINCT customer_name, archive_file_name FROM {table}
            WHERE customer_name IS NOT NULL
        """).fetchall()

        skipped_old = 0
        undated = 0
        for value, archive_file_name in rows:
            if cutoff is not None:
                processed_date = extract_processed_date(archive_file_name)
                if processed_date is not None and processed_date < cutoff:
                    skipped_old += 1
                    continue
                if processed_date is None:
                    undated += 1
            result[value]["sources"].add(source_label)
            if archive_file_name:
                result[value]["files"].add(archive_file_name)

        if cutoff is not None:
            note = (f"  {table}: skipped {skipped_old} row(s) older than "
                    f"{lookback_days} day(s)")
            if undated:
                note += (f"; {undated} row(s) had no readable date and were "
                         f"included anyway")
            print(note)

    con.close()
    return dict(result)


# ---------------------------------------------------------------------------
# Source 3: Template files (read fresh — no pipeline/DuckDB table exists yet)
# ---------------------------------------------------------------------------

def get_template_contractor_names() -> dict[str, dict]:
    """Return {value: {"sources": {"template"}, "files": {file names}}} for
    every non-blank contractor_name value found across all Template Excel
    files — both from the contractor_name column's cell data, AND from each
    file name itself (position 6, 0-indexed, splitting on "_" — the same
    slot Template file names embed the contractor name in). Values are used
    exactly as read/extracted (no trimming), since Template files don't go
    through DuckDB at all."""
    result: dict[str, dict] = defaultdict(lambda: {"sources": set(), "files": set()})

    if not TEMPLATES_DIR.exists():
        print(f"  WARNING: templates folder not found: {TEMPLATES_DIR}")
        return dict(result)

    excel_files = sorted(
        f for ext in TEMPLATE_EXTENSIONS
        for f in TEMPLATES_DIR.glob(f"*{ext}")
    )

    header_index = TEMPLATE_HEADER_ROW - 1  # convert to 0-indexed
    FILENAME_CONTRACTOR_POSITION = 6  # 0-indexed, splitting the file name on "_"

    for file_path in excel_files:
        # Filename-embedded contractor name — independent of whether the
        # sheet/column read below succeeds, so a file with a broken sheet
        # still contributes this check.
        name_parts = file_path.name.split("_")
        if len(name_parts) > FILENAME_CONTRACTOR_POSITION:
            filename_value = name_parts[FILENAME_CONTRACTOR_POSITION]
            if filename_value.strip():
                result[filename_value]["sources"].add("template")
                result[filename_value]["files"].add(file_path.name)

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
            result[value]["sources"].add("template")
            result[value]["files"].add(file_path.name)


    return dict(result)


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
    # globals().get(...) rather than referencing LOOKBACK_DAYS directly, so
    # commenting that config line out entirely (to disable the limit) doesn't
    # raise a NameError — it just resolves to None here.
    lookback_days = globals().get("LOOKBACK_DAYS")

    if lookback_days is not None:
        print(f"Gathering customer_name values from DuckDB (base + credit), "
              f"limited to files processed in the last {lookback_days} day(s)...")
    else:
        print("Gathering customer_name values from DuckDB (base + credit), no date limit...")
    combined: dict[str, dict] = defaultdict(lambda: {"sources": set(), "files": set()})

    for value, info in get_duckdb_customer_names(lookback_days).items():
        combined[value]["sources"] |= info["sources"]
        combined[value]["files"] |= info["files"]
    print(f"  {len(combined)} unique value(s) so far (base + credit)")

    print("\nGathering contractor_name values from Template files...")
    for value, info in get_template_contractor_names().items():
        combined[value]["sources"] |= info["sources"]
        combined[value]["files"] |= info["files"]
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
        source_label = "/".join(sorted(combined[value]["sources"]))
        file_list = ", ".join(sorted(combined[value]["files"]))
        note = f"added {today} from {source_label} file data"
        if file_list:
            note += f" ({file_list})"
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
        print(f"  + {value!r}  ({'/'.join(sorted(combined[value]['sources']))})")


if __name__ == "__main__":
    main()