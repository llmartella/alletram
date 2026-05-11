import requests
import json
import gspread
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os
import pickle
from typing import Dict, Any, Optional, List

# Google Sheets API scope
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

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
    Extract all unique column names from the data.
    
    Args:
        data: List of dictionaries
        
    Returns:
        List of column names
    """
    columns = set()
    for item in data:
        if isinstance(item, dict):
            columns.update(item.keys())
    
    return sorted(list(columns))

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

def main():
    """Main function to run the API to Google Sheets sync."""
    
    # ========== CONFIGURATION ==========
    
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
    
    # Step 3: Extract column names
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
- The script automatically detects column structure from your API data
- It handles different API response formats (arrays, objects with data arrays, etc.)
- Large datasets are written in batches for better performance
- Credentials are saved locally for future runs (token.pickle file)
"""
