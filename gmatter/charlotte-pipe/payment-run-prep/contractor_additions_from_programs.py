import gspread
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
import os
import pickle
from collections import Counter
from typing import List, Set, Dict

# ========== CONFIGURATION ==========

SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

# OAuth2 credentials file (same one used in your other Sheets scripts)
OAUTH2_CREDENTIALS_FILE = "/Users/lorimartella/Documents/gmatter/client_secret_490677366778-qkpf9k95otitjrneaof6sjtajod97bgc.apps.googleusercontent.com.json"

# Programs sheet (source of contractor name variations)
PROGRAMS_SHEET_ID = "1MXm6sngsqDxAErjZm9odsaS80AS1HPqvhstfCFhdmys"
PROGRAMS_TAB_NAME = "API_Data"
PROGRAMS_CONTRACTOR_COLUMN = "Contractor_Offer_Name__r_Name"
PROGRAMS_VENDOR_RECORD_COLUMN = "Vendor_Record_Number__c"

# Contractors sheet (master list of known variations)
CONTRACTORS_SHEET_ID = "17muso1TB00rbBYH0hiMzaCx0k8y_3u0OKhdKkrvgok4"
CONTRACTORS_TAB_NAME = "contractors"
CONTRACTORS_ORIGINAL_NAME_COLUMN = "original_contractor_name"

# Review tab: created inside the PROGRAMS spreadsheet
REVIEW_TAB_NAME = "Needs Review"
REVIEW_TAB_HEADERS = [
    "original_contractor_name",
    "count_in_programs",
    "vendor_record_numbers",
    "preferred_contractor_name",
    "preferred_contractor_number",
    "contractor_hq",
    "notes",
]

# ========== AUTH ==========

def authenticate_google_sheets() -> gspread.Client:
    """Authenticate with Google Sheets using OAuth2, reusing a cached token if available."""
    creds = None
    token_file = 'token.pickle'

    if os.path.exists(token_file):
        with open(token_file, 'rb') as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            print("Refreshing expired credentials...")
            creds.refresh(Request())
        else:
            print("Starting OAuth2 flow...")
            print("A browser window will open for authentication.")
            flow = InstalledAppFlow.from_client_secrets_file(OAUTH2_CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)

        with open(token_file, 'wb') as token:
            pickle.dump(creds, token)

    print("✅ Authenticated with Google Sheets!")
    return gspread.authorize(creds)


# ========== HELPERS ==========

def get_column_values(client: gspread.Client, sheet_id: str, tab_name: str, column_header: str) -> List[str]:
    """
    Return ALL raw values (in row order, including duplicates) from the column
    with the given header on the given tab. No trimming, no case changes —
    values are returned exactly as stored.
    """
    sheet = client.open_by_key(sheet_id)
    worksheet = sheet.worksheet(tab_name)

    all_values = worksheet.get_all_values()  # list of rows, each a list of cell strings
    if not all_values:
        raise ValueError(f"Tab '{tab_name}' appears to be empty.")

    header_row = all_values[0]
    try:
        col_index = header_row.index(column_header)
    except ValueError:
        raise ValueError(
            f"Column '{column_header}' not found in tab '{tab_name}'. "
            f"Headers found: {header_row}"
        )

    values = []
    for row in all_values[1:]:
        if col_index < len(row):
            value = row[col_index]
        else:
            value = ""  # row shorter than header row
        if value != "":  # skip blank cells, keep everything else exactly as-is
            values.append(value)

    return values


def get_two_column_values(
    client: gspread.Client, sheet_id: str, tab_name: str, name_column: str, extra_column: str
) -> List[tuple]:
    """
    Return (name_value, extra_value) pairs, one per row, aligned by row position.
    Rows with a blank name_value are skipped. extra_value is returned exactly as
    stored (or "" if that cell is blank/missing).
    """
    sheet = client.open_by_key(sheet_id)
    worksheet = sheet.worksheet(tab_name)

    all_values = worksheet.get_all_values()
    if not all_values:
        raise ValueError(f"Tab '{tab_name}' appears to be empty.")

    header_row = all_values[0]
    try:
        name_index = header_row.index(name_column)
    except ValueError:
        raise ValueError(f"Column '{name_column}' not found in tab '{tab_name}'. Headers found: {header_row}")
    try:
        extra_index = header_row.index(extra_column)
    except ValueError:
        raise ValueError(f"Column '{extra_column}' not found in tab '{tab_name}'. Headers found: {header_row}")

    pairs = []
    for row in all_values[1:]:
        name_value = row[name_index] if name_index < len(row) else ""
        if name_value == "":
            continue
        extra_value = row[extra_index] if extra_index < len(row) else ""
        pairs.append((name_value, extra_value))

    return pairs


def get_or_create_review_tab(client: gspread.Client, sheet_id: str) -> gspread.Worksheet:
    """
    Get the Needs Review tab in the Programs spreadsheet, creating it with headers if missing.
    If the tab already exists from a previous run with an older/shorter header row, insert
    any missing headers (e.g. vendor_record_numbers) in the correct position rather than
    overwriting or duplicating existing data.
    """
    sheet = client.open_by_key(sheet_id)
    try:
        worksheet = sheet.worksheet(REVIEW_TAB_NAME)
        print(f"✅ Found existing '{REVIEW_TAB_NAME}' tab")
        _ensure_headers_up_to_date(worksheet)
    except gspread.WorksheetNotFound:
        print(f"📄 Creating new '{REVIEW_TAB_NAME}' tab in Programs spreadsheet")
        worksheet = sheet.add_worksheet(title=REVIEW_TAB_NAME, rows=1000, cols=len(REVIEW_TAB_HEADERS) + 2)
        worksheet.update("A1", [REVIEW_TAB_HEADERS])
    return worksheet


def _ensure_headers_up_to_date(worksheet: gspread.Worksheet) -> None:
    """Insert any REVIEW_TAB_HEADERS columns missing from an existing tab, in the correct position."""
    current_headers = worksheet.row_values(1)

    for expected_index, header in enumerate(REVIEW_TAB_HEADERS):
        if header not in current_headers:
            insert_at = expected_index + 1  # gspread insert_cols is 1-indexed
            print(f"🔧 '{REVIEW_TAB_NAME}' tab is missing '{header}' — inserting it as column {insert_at}")
            worksheet.insert_cols([[header]], insert_at)
            current_headers = worksheet.row_values(1)  # refresh after insert


def get_existing_review_entries(worksheet: gspread.Worksheet) -> Set[str]:
    """Return the set of original_contractor_name values already listed in the Needs Review tab."""
    all_values = worksheet.get_all_values()
    if not all_values:
        return set()

    header_row = all_values[0]
    if REVIEW_TAB_HEADERS[0] not in header_row:
        return set()

    col_index = header_row.index(REVIEW_TAB_HEADERS[0])
    existing = set()
    for row in all_values[1:]:
        if col_index < len(row) and row[col_index] != "":
            existing.add(row[col_index])
    return existing


# ========== MAIN ==========

def main():
    print("🚀 Checking for contractor name variations missing from the Contractors sheet...")
    print("=" * 70)

    client = authenticate_google_sheets()

    # Step 1: Known variations already in the Contractors sheet (exact strings)
    print(f"\n📋 Reading known variations from Contractors sheet ('{CONTRACTORS_TAB_NAME}' tab)...")
    known_variations = set(get_column_values(
        client, CONTRACTORS_SHEET_ID, CONTRACTORS_TAB_NAME, CONTRACTORS_ORIGINAL_NAME_COLUMN
    ))
    print(f"   Found {len(known_variations)} known variation(s)")

    # Step 2: All contractor name + Vendor_Record_Number__c occurrences in the Programs sheet
    print(f"\n📋 Reading contractor names and vendor record numbers from Programs sheet ('{PROGRAMS_TAB_NAME}' tab)...")
    name_vendor_pairs = get_two_column_values(
        client, PROGRAMS_SHEET_ID, PROGRAMS_TAB_NAME,
        PROGRAMS_CONTRACTOR_COLUMN, PROGRAMS_VENDOR_RECORD_COLUMN
    )
    print(f"   Found {len(name_vendor_pairs)} contractor reference(s) across all program rows")

    programs_values = [name for name, _ in name_vendor_pairs]
    counts = Counter(programs_values)
    print(f"   {len(counts)} unique variation(s) among them")

    # Collect the distinct vendor record number(s) seen for each name variation
    vendor_numbers_by_name: Dict[str, List[str]] = {}
    for name, vendor_number in name_vendor_pairs:
        bucket = vendor_numbers_by_name.setdefault(name, [])
        if vendor_number != "" and vendor_number not in bucket:
            bucket.append(vendor_number)

    # Step 3: What's already staged for review, so we don't duplicate or clobber progress
    print(f"\n📋 Checking existing '{REVIEW_TAB_NAME}' tab for already-staged entries...")
    review_worksheet = get_or_create_review_tab(client, PROGRAMS_SHEET_ID)
    already_staged = get_existing_review_entries(review_worksheet)
    print(f"   {len(already_staged)} variation(s) already staged for review")

    # Step 4: Determine what's missing — exact match only, no fuzzy/trim/case logic
    missing = [
        (name, count) for name, count in counts.items()
        if name not in known_variations and name not in already_staged
    ]
    missing.sort(key=lambda x: (-x[1], x[0]))  # most common first, then alphabetical

    print(f"\n🔍 Result: {len(missing)} new variation(s) not found in Contractors sheet and not yet staged")

    if not missing:
        print("✅ Nothing new to add — every variation in Programs is already accounted for.")
        return

    # Step 5: Append new rows to the Needs Review tab
    new_rows = []
    flagged_count = 0
    for name, count in missing:
        vendor_numbers = vendor_numbers_by_name.get(name, [])
        note = ""
        if len(vendor_numbers) > 1:
            note = f"⚠️ {len(vendor_numbers)} different Vendor Record Numbers found for this name — needs review"
            flagged_count += 1
        new_rows.append([name, str(count), "; ".join(vendor_numbers), "", "", "", note])
    review_worksheet.append_rows(new_rows, value_input_option="USER_ENTERED")

    if flagged_count:
        print(f"⚠️ {flagged_count} variation(s) flagged with multiple distinct Vendor Record Numbers — see notes column")

    print(f"✅ Appended {len(new_rows)} row(s) to '{REVIEW_TAB_NAME}' tab in the Programs spreadsheet.")
    print("\nNext steps:")
    print("  1. Open the Needs Review tab and fill in preferred_contractor_name (and number/hq/notes if known).")
    print("  2. Copy finished rows over to the main Contractors sheet's 'contractors' tab.")
    print("  3. Delete/clear finished rows from Needs Review once copied, if you'd like a clean slate.")
    print(f"\n🔗 Programs sheet: https://docs.google.com/spreadsheets/d/{PROGRAMS_SHEET_ID}")


if __name__ == "__main__":
    main()

# ========== SETUP NOTES ==========
"""
- Requires: pip install gspread google-auth google-auth-oauthlib google-auth-httplib2
- Uses the same OAuth2 credentials file / token.pickle pattern as your other Sheets scripts.
  Both spreadsheets must be accessible to whatever Google account you authenticate with.
- Matching is fully exact-string: no trimming, no lowercasing, no fuzzy matching.
  A trailing space or a different capitalization counts as a distinct variation.
- vendor_record_numbers pulls from the Programs sheet's Vendor_Record_Number__c column.
  If a single contractor variation shows up with more than one distinct vendor record
  number across program rows, all distinct values are listed, separated by "; ".
- If a contractor name has more than one distinct Vendor Record Number, the notes column
  is pre-filled with a ⚠️ flag calling that out. This is visibility only — no automatic
  resolution happens. Nothing is written to the Contractors sheet based on this flag.
- Safe to re-run: it never touches rows already sitting in the Needs Review tab, so your
  in-progress manual edits there are preserved. It only appends genuinely new variations.
- Once you've moved a reviewed contractor's row into the main Contractors sheet, you can
  delete its row from Needs Review — the script won't re-add it as long as it's now present
  (exactly) in the Contractors sheet's original_contractor_name column.
"""