import requests
import json
import gspread
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os
import pickle
from datetime import date
from typing import Dict, Any, Optional, List, Set, Tuple

# Google Sheets API scope
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

# Fixed column order for the output sheet to show relevant fields first.
COLUMN_ORDER = [
    "Branch_Name",
    "Branch_Number",
    "Contractor_Offer_Name__r_Name",
    "Name",
    "Offer_Type__c",
    "Description__c",
    "DiscountPercentage",
    "Fittings_text__c",
    "Pipe_text__c",
    "Start_Date__c",
    "Expiration_Date__c",
    "Material_Group__c",
    "Status__c",
    "Vendor_Record_Number__c",
    "Division__c",
    "Offer_Summary__c",
    "CreatedDate",
    "SAP_Acct__c",
    "Discount_Amount__c",
    "Commission_Group__c",
    "Contractor_Key__c",
    "Contractor_Offer_Name__c",
    "Contractor_Offer_Name__r_Id",
    "Contractor_Offer_Name__r_type",
    "Contractor_Program__r",
    "CreatedById",
    "Disount_Type__c",
    "Estimated_Annual_Accrual__c",
    "Id",
    "IsDeleted",
    "Job_Address__City__s",
    "Job_Address__CountryCode__s",
    "Job_Address__PostalCode__s",
    "Job_Address__StateCode__s",
    "Job_Address__Street__s",
    "Job_Project_Name_or_Quote__c",
    "LastModifiedById",
    "LastModifiedDate",
    "Offer_Number__c",
    "OwnerId",
    "Payment_Schedule__c",
    "Payment_Type__c",
    "Percentage_Discount__c",
    "Price_List__c",
    "Program_Material_CreatedById",
    "Program_Material_LastModifiedById",
    "Program_Material_Name",
    "Reason_for_making_the_offer__c",
    "RecordTypeId",
    "Regional_Manager__c",
    "Sales_Rep_Account__c",
    "Sales_Rep_Account__r_Name",
    "Sales_Rep_Account__r_type",
    "Who_will_be_paid_or_credited__c",
    "Year_Active__c",
    "type",
]

def authenticate_google_sheets(credentials_file: str) -> gspread.Client:
    """
    Authenticate with Google Sheets using OAuth2 client credentials.
    
    Args:
        credentials_file: Path to OAuth2 client credentials JSON file
        
    Returns:
        Authenticated gspread client
    """
    creds = None
    token_file = 'token.pickle'
    
    # Check if we have stored credentials
    if os.path.exists(token_file):
        with open(token_file, 'rb') as token:
            creds = pickle.load(token)
    
    # If there are no valid credentials, get them
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            print("Refreshing expired credentials...")
            creds.refresh(Request())
        else:
            print("Starting OAuth2 flow...")
            print("A browser window will open for authentication.")
            flow = InstalledAppFlow.from_client_secrets_file(credentials_file, SCOPES)
            creds = flow.run_local_server(port=0)
        
        # Save credentials for next run
        with open(token_file, 'wb') as token:
            pickle.dump(creds, token)
    
    print("✅ Successfully authenticated with Google Sheets!")
    return gspread.authorize(creds)

def fetch_api_data(endpoint_url: str, params: Optional[Dict] = None) -> Optional[List[Dict]]:
    """
    Fetch data from API endpoint.
    
    Args:
        endpoint_url: API endpoint URL
        params: Optional query parameters
        
    Returns:
        List of dictionaries containing the API data
    """
    try:
        print(f"🌐 Fetching data from: {endpoint_url}")
        
        response = requests.get(endpoint_url, params=params, timeout=30)
        response.raise_for_status()
        
        data = response.json()
        print(f"📊 API Response type: {type(data)}")
        
        # Handle different response structures
        if isinstance(data, list):
            print(f"✅ Found {len(data)} records")
            return data
        elif isinstance(data, dict):
            # Look for common array keys
            for key in ['data', 'results', 'items', 'records', 'list']:
                if key in data and isinstance(data[key], list):
                    print(f"✅ Found {len(data[key])} records in '{key}' field")
                    return data[key]
            
            # If no array found, treat as single record
            print("✅ Single record found, converting to list")
            return [data]
        else:
            print(f"❌ Unexpected data type: {type(data)}")
            return None
            
    except requests.exceptions.RequestException as e:
        print(f"❌ API request failed: {e}")
        return None
    except json.JSONDecodeError as e:
        print(f"❌ Invalid JSON response: {e}")
        return None

def extract_columns_from_data(data: List[Dict]) -> List[str]:
    """
    Return the column order for the sheet: the fixed COLUMN_ORDER first
    (in the specified order), followed by any additional columns found
    in the data that aren't part of COLUMN_ORDER (appended alphabetically
    at the end, so no data gets silently dropped).
    
    Args:
        data: List of dictionaries
        
    Returns:
        List of column names in the desired order
    """
    all_keys = set()
    for item in data:
        if isinstance(item, dict):
            all_keys.update(item.keys())
    
    ordered_columns = [col for col in COLUMN_ORDER if col in all_keys]
    extra_columns = sorted(all_keys - set(COLUMN_ORDER))
    
    if extra_columns:
        print(f"⚠️ Found {len(extra_columns)} columns not in the defined order, appending at end: {extra_columns}")
    
    return ordered_columns + extra_columns

def prepare_sheet_data(data: List[Dict], columns: List[str]) -> List[List]:
    """
    Convert API data to rows for Google Sheets.
    
    Args:
        data: List of dictionaries from API
        columns: List of column names
        
    Returns:
        List of lists ready for Google Sheets
    """
    rows = []
    
    for item in data:
        row = []
        for col in columns:
            value = item.get(col, "")
            
            # Handle different data types
            if value is None:
                row.append("")
            elif isinstance(value, (list, dict)):
                # Convert complex objects to JSON string
                row.append(json.dumps(value))
            elif isinstance(value, bool):
                row.append("TRUE" if value else "FALSE")
            else:
                row.append(str(value))
        
        rows.append(row)
    
    return rows

def write_to_sheet(client: gspread.Client, sheet_id: str, worksheet_name: str, 
                   headers: List[str], data: List[List]) -> bool:
    """
    Write data to Google Sheets.
    
    Args:
        client: Authenticated gspread client
        sheet_id: Google Sheet ID
        worksheet_name: Worksheet name
        headers: Column headers
        data: Data rows
        
    Returns:
        True if successful
    """
    try:
        print(f"📝 Opening Google Sheet...")
        sheet = client.open_by_key(sheet_id)
        
        # Try to get existing worksheet or create new one
        try:
            worksheet = sheet.worksheet(worksheet_name)
            print(f"✅ Found existing worksheet: {worksheet_name}")
        except gspread.WorksheetNotFound:
            print(f"📄 Creating new worksheet: {worksheet_name}")
            # Create worksheet with generous dimensions
            num_cols = max(len(headers) + 10, 50)  # Extra columns for safety
            num_rows = max(len(data) + 100, 5000)  # Extra rows for safety
            worksheet = sheet.add_worksheet(title=worksheet_name, rows=num_rows, cols=num_cols)
        
        # Ensure worksheet has enough rows and columns
        current_rows = worksheet.row_count
        current_cols = worksheet.col_count
        needed_rows = len(data) + 10  # +10 for headers and buffer
        needed_cols = len(headers) + 5  # +5 for buffer
        
        resize_needed = False
        new_rows = current_rows
        new_cols = current_cols
        
        if current_rows < needed_rows:
            new_rows = needed_rows + 100  # Extra buffer
            resize_needed = True
            print(f"📏 Need to expand rows from {current_rows} to {new_rows}")
            
        if current_cols < needed_cols:
            new_cols = needed_cols + 10  # Extra buffer
            resize_needed = True
            print(f"📏 Need to expand columns from {current_cols} to {new_cols}")
        
        if resize_needed:
            print(f"📐 Resizing worksheet to {new_rows} rows × {new_cols} columns...")
            worksheet.resize(rows=new_rows, cols=new_cols)
            print("✅ Worksheet resized successfully")
        
        # Clear existing data
        print("🧹 Clearing existing data...")
        worksheet.clear()
        
        # Prepare all data (headers + rows)
        all_data = [headers] + data
        
        print(f"💾 Writing {len(data)} rows to Google Sheets...")
        
        # Use larger batches but with rate limiting
        batch_size = 100  # Smaller batches to avoid rate limits
        total_batches = (len(all_data) + batch_size - 1) // batch_size
        
        for i in range(0, len(all_data), batch_size):
            batch = all_data[i:i + batch_size]
            batch_num = (i // batch_size) + 1
            start_row = i + 1
            end_row = start_row + len(batch) - 1
            
            # Calculate end column dynamically to handle any number of columns
            num_cols = len(headers)
            if num_cols <= 26:
                end_col = chr(ord('A') + num_cols - 1)
            else:
                # Handle columns beyond Z (AA, AB, etc.)
                first_char = chr(ord('A') + (num_cols - 1) // 26 - 1) if num_cols > 26 else ''
                second_char = chr(ord('A') + (num_cols - 1) % 26)
                end_col = first_char + second_char
            
            range_name = f"A{start_row}:{end_col}{end_row}"
            print(f"📊 Writing batch {batch_num}/{total_batches}: {range_name} ({len(batch)} rows)")
            
            # Retry mechanism for rate limiting
            max_retries = 3
            for attempt in range(max_retries):
                try:
                    worksheet.update(range_name, batch, value_input_option='USER_ENTERED')
                    print(f"✅ Batch {batch_num} completed successfully")
                    break
                    
                except Exception as batch_error:
                    error_str = str(batch_error)
                    print(f"❌ Batch {batch_num} attempt {attempt + 1} failed: {batch_error}")
                    
                    if "Quota exceeded" in error_str or "429" in error_str:
                        # Rate limit hit - wait longer
                        wait_time = (attempt + 1) * 30  # 30, 60, 90 seconds
                        print(f"⏳ Rate limit hit. Waiting {wait_time} seconds before retry...")
                        import time
                        time.sleep(wait_time)
                        
                    elif "exceeds grid limits" in error_str:
                        print(f"❌ Grid size error for batch {batch_num}. Skipping this batch.")
                        print(f"   Range: {range_name}")
                        print(f"   Consider reducing data size or increasing sheet dimensions.")
                        break
                        
                    elif attempt == max_retries - 1:
                        print(f"❌ Batch {batch_num} failed after {max_retries} attempts")
                        # Try row-by-row as last resort
                        print("🔄 Trying row-by-row update for this batch...")
                        for j, row in enumerate(batch):
                            row_num = start_row + j
                            try:
                                worksheet.update(f"A{row_num}:{end_col}{row_num}", [row])
                                if j % 10 == 0:  # Progress indicator
                                    print(f"   📝 Row {row_num} completed")
                                # Small delay to avoid rate limits
                                import time
                                time.sleep(0.5)
                            except Exception as row_error:
                                print(f"❌ Failed to write row {row_num}: {row_error}")
                        break
                    else:
                        # Wait before next attempt
                        import time
                        time.sleep(5)
            
            # Rate limiting: wait between batches
            if batch_num < total_batches:  # Don't wait after the last batch
                import time
                time.sleep(2)  # 2 second delay between batches
        
        print(f"✅ Successfully wrote data to Google Sheets!")
        print(f"🔗 View your sheet: https://docs.google.com/spreadsheets/d/{sheet_id}")
        
        return True
        
    except Exception as e:
        print(f"❌ Error writing to Google Sheets: {e}")
        return False

def get_program_contractor_names(client: gspread.Client, sheet_id: str, tab_name: str,
                                   column_header: str) -> List[str]:
    """
    Return ALL raw values (in row order, including duplicates) from the given
    column on the given tab. No trimming, no case changes — values are
    returned exactly as stored.
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


def get_contractors_names_and_next_row(
    client: gspread.Client, sheet_id: str, tab_name: str, column_header: str
) -> Tuple[gspread.Worksheet, Set[str], int]:
    """Return (worksheet, existing_values_set, next_empty_row_number) for the
    given column, found by header text rather than assuming a fixed column
    letter."""
    sheet = client.open_by_key(sheet_id)
    worksheet = sheet.worksheet(tab_name)

    all_values = worksheet.get_all_values()
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

    existing: Set[str] = set()
    last_row_with_data = 1  # header row
    for i, row in enumerate(all_values[1:], start=2):
        if col_index < len(row) and row[col_index] != "":
            existing.add(row[col_index])
            last_row_with_data = i

    next_row = last_row_with_data + 1
    return worksheet, existing, next_row


def add_new_contractor_names(client: gspread.Client) -> None:
    """
    Find contractor name variations in the Programs sheet's API_Data tab
    (the sheet/tab this script just wrote to) that aren't already in the
    Contractors sheet (exact string match), and append them directly to the
    bottom of the contractors tab's original_contractor_name column, with a
    note identifying the source and date.
    """
    # Programs sheet (source of contractor name variations) — same sheet/tab
    # this script's write_to_sheet() step just populated.
    PROGRAMS_SHEET_ID = "1MXm6sngsqDxAErjZm9odsaS80AS1HPqvhstfCFhdmys"
    PROGRAMS_TAB_NAME = "API_Data"
    PROGRAMS_CONTRACTOR_COLUMN = "Contractor_Offer_Name__r_Name"

    # Contractors sheet (master list of known variations) — new values get
    # appended directly here.
    CONTRACTORS_SHEET_ID            = "17muso1TB00rbBYH0hiMzaCx0k8y_3u0OKhdKkrvgok4"
    CONTRACTORS_TAB_NAME            = "contractors"
    CONTRACTORS_NAME_COLUMN_HEADER  = "original_contractor_name"
    CONTRACTORS_NAME_COLUMN_LETTER  = "A"
    CONTRACTORS_NOTES_COLUMN_LETTER = "E"

    NOTE_SOURCE_LABEL = "program material endpoint"

    print("🚀 Checking for contractor name variations missing from the Contractors sheet...")
    print("=" * 70)

    # Step 1: Contractor name variations from the Programs sheet
    print(f"\n📋 Reading contractor names from Programs sheet ('{PROGRAMS_TAB_NAME}' tab)...")
    program_values = get_program_contractor_names(
        client, PROGRAMS_SHEET_ID, PROGRAMS_TAB_NAME, PROGRAMS_CONTRACTOR_COLUMN
    )
    unique_program_names = set(program_values)
    print(f"   Found {len(program_values)} reference(s), {len(unique_program_names)} unique variation(s)")

    # Step 2: Known variations already in the Contractors sheet (exact strings)
    print(f"\n📋 Reading existing names from Contractors sheet ('{CONTRACTORS_TAB_NAME}' tab)...")
    worksheet, known_variations, next_row = get_contractors_names_and_next_row(
        client, CONTRACTORS_SHEET_ID, CONTRACTORS_TAB_NAME, CONTRACTORS_NAME_COLUMN_HEADER
    )
    print(f"   Found {len(known_variations)} known variation(s)")

    # Step 3: Determine what's missing — exact match only, no fuzzy/trim/case logic
    missing = sorted(name for name in unique_program_names if name not in known_variations)

    print(f"\n🔍 Result: {len(missing)} new variation(s) not found in Contractors sheet")

    if not missing:
        print("✅ Nothing new to add — every variation in Programs is already accounted for.")
        return

    # Step 4: Append new rows directly to the Contractors sheet
    today = date.today().isoformat()
    updates = []
    for offset, name in enumerate(missing):
        row_num = next_row + offset
        note = f"added {today} from {NOTE_SOURCE_LABEL}"
        updates.append({
            "range": f"{CONTRACTORS_NAME_COLUMN_LETTER}{row_num}",
            "values": [[name]],
        })
        updates.append({
            "range": f"{CONTRACTORS_NOTES_COLUMN_LETTER}{row_num}",
            "values": [[note]],
        })

    # The sheet's grid may not have enough rows allocated yet — expand it
    # first if needed (separate from how many rows actually contain data).
    required_rows = next_row + len(missing) - 1
    if worksheet.row_count < required_rows:
        worksheet.add_rows(required_rows - worksheet.row_count)

    print(f"📝 Appending {len(missing)} new row(s) to '{CONTRACTORS_TAB_NAME}' "
          f"(rows {next_row}-{next_row + len(missing) - 1})...")
    worksheet.batch_update(updates, value_input_option="USER_ENTERED")

    print(f"✅ Appended {len(missing)} row(s) to '{CONTRACTORS_TAB_NAME}' tab.")
    for name in missing:
        print(f"  + {name!r}")


def main():
    """Main function to run the API to Google Sheets sync."""
    
    # API Configuration
    API_ENDPOINT = "https://cpf-api.charlottepipe.com/GET_ContractorDB_ContractorProgram_ContractorProgramMaterial?AuthorizationNumber=ClairVoyant_4GV&Passcode=258a1a49-a5b4-48fd-bf65-b549ea36143a"  # Replace with your API endpoint
    API_PARAMS = {
        # Add any query parameters here
        # "limit": 10,
        # "page": 1
    }
    
    # Google Sheets Configuration
    OAUTH2_CREDENTIALS_FILE = "/Users/lorimartella/Documents/gmatter/client_secret_490677366778-qkpf9k95otitjrneaof6sjtajod97bgc.apps.googleusercontent.com.json"  # Path to your OAuth2 client credentials file
    SHEET_ID = "1MXm6sngsqDxAErjZm9odsaS80AS1HPqvhstfCFhdmys"  # Replace with your Google Sheet ID
    WORKSHEET_NAME = "API_Data"  # Name of the worksheet to create/update
    
    # ========== EXECUTION ==========
    
    print("🚀 Starting API to Google Sheets sync...")
    print("=" * 50)
    
    # Step 1: Authenticate with Google Sheets
    try:
        client = authenticate_google_sheets(OAUTH2_CREDENTIALS_FILE)
    except Exception as e:
        print(f"❌ Authentication failed: {e}")
        print("\nTroubleshooting:")
        print("1. Make sure your credentials.json file is in the correct location")
        print("2. Ensure the file is an OAuth2 client credentials file")
        print("3. Check that you've enabled the Google Sheets API in Google Cloud Console")
        return
    
    # Step 2: Fetch data from API
    api_data = fetch_api_data(API_ENDPOINT, API_PARAMS)
    if not api_data:
        print("❌ Failed to fetch data from API")
        return
    
    # Step 3: Extract column names (in the defined order)
    columns = extract_columns_from_data(api_data)
    print(f"📋 Found {len(columns)} columns: {columns[:10]}{'...' if len(columns) > 10 else ''}")
    
    # Debug: Show sample data structure
    if api_data:
        print(f"📊 Sample record keys: {list(api_data[0].keys()) if isinstance(api_data[0], dict) else 'Not a dict'}")
    
    # Step 4: Prepare data for sheets
    sheet_data = prepare_sheet_data(api_data, columns)
    print(f"📊 Prepared {len(sheet_data)} rows for Google Sheets")
    
    # Step 5: Write to Google Sheets
    success = write_to_sheet(client, SHEET_ID, WORKSHEET_NAME, columns, sheet_data)

    if success:
        print("\n🎉 Script completed successfully!")

        # Step 6: Check for new contractor names in the freshly-written
        # API_Data and add any not already in the Contractors sheet.
        print("\n" + "=" * 50)
        try:
            add_new_contractor_names(client)
        except Exception as e:
            print(f"❌ Contractor name check failed: {e}")
    else:
        print("\n❌ Script failed!")

if __name__ == "__main__":
    main()

# ========== SETUP INSTRUCTIONS ==========
"""
SETUP INSTRUCTIONS:

1. Install required packages:
   pip install gspread google-auth google-auth-oauthlib google-auth-httplib2 requests

2. Set up OAuth2 credentials:
   - Go to Google Cloud Console (console.cloud.google.com)
   - Create a project or select existing one
   - Enable Google Sheets API and Google Drive API
   - Go to "Credentials" > "Create Credentials" > "OAuth client ID"
   - Choose "Desktop application"
   - Download the JSON file and rename it to "credentials.json"
   - Place it in the same folder as this script

3. Get your Google Sheet ID:
   - Create a new Google Sheet or use existing one
   - Copy the Sheet ID from the URL:
     https://docs.google.com/spreadsheets/d/SHEET_ID_HERE/edit
   - Replace "your_sheet_id_here" in the script

4. Update configuration:
   - Replace API_ENDPOINT with your actual API URL
   - Add any required API parameters
   - Update OAUTH2_CREDENTIALS_FILE path if needed

5. Run the script:
   python script_name.py
   
   On first run, it will open a browser for authentication.
   Subsequent runs will use saved credentials.

NOTES:
- Columns are written in the fixed order defined in COLUMN_ORDER at the
  top of this file. Any columns present in the API data but not listed
  in COLUMN_ORDER are appended (alphabetically) after the defined ones,
  so no data is silently dropped if the API adds a new field.
- The script automatically detects column structure from your API data
- It handles different API response formats (arrays, objects with data arrays, etc.)
- Large datasets are written in batches for better performance
- Credentials are saved locally for future runs (token.pickle file)
- After a successful write, this script automatically checks the freshly
  written API_Data for contractor name variations not already present in
  the Contractors sheet, and appends any new ones directly to that sheet's
  original_contractor_name column, with a note in column E:
  "added YYYY-MM-DD from program material endpoint". This logic is fully
  self-contained in this file — no separate script or import needed.
"""