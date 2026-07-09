"""
Consolidated Azure Blob Log Processing Script
Processes logs from multiple services and writes to CSV, Supabase, and Google Drive
"""

import os
import csv
import re
from azure.storage.blob import BlobServiceClient
from datetime import datetime
from dateutil.relativedelta import relativedelta
from supabase import create_client, Client
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# ===== CONFIGURATION =====
# replace this with info from 1Password Secure Notes


#---


# ===== OPTIONAL OVERRIDE =====
# If set, this specific year/month will be used instead of previous-month logic
# Example: SPECIFIC_YEAR_MONTH = "2026/01"
# Leave empty ("") to use previous month automatically
SPECIFIC_YEAR_MONTH = ""   # Format: "YYYY/MM" or leave empty

# Get script directory for output files
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# ===== DATE CALCULATIONS OR SPECIFIC OVERRIDE =====
if SPECIFIC_YEAR_MONTH:
    # Use specific year/month provided
    try:
        year, month = SPECIFIC_YEAR_MONTH.split('/')
        print(f"Using specific year/month: {year}-{month}")
    except ValueError:
        print("Error: SPECIFIC_YEAR_MONTH must be in format 'YYYY/MM' (e.g., '2026/01')")
        exit(1)
else:
    # Use previous month automatically
    today = datetime.today()
    previous_month_date = today - relativedelta(months=1)
    year = previous_month_date.strftime("%Y")
    month = previous_month_date.strftime("%m")
    print(f"Using previous month: {year}-{month}")

# Output CSV filename
output_csv = os.path.join(SCRIPT_DIR, f"v1_usage_{year}_{month}.csv")

# Service configurations
SERVICES = [
    {
        "name": "FreshAddress Email Validation",
        "container": "incindio-prod",
        "directory": f"incindio-app-triggered-FreshAddress-Action/{year}/{month}/",
        "keyword": "Org:".lower(),
        "file_extensions": [".csv"]
    },
    {
        "name": "FreshAddress RT Email Validation",
        "container": "incindio-app-validation-prod",
        "directory": f"incindio-app-validation/{year}/{month}/",
        "keyword": "FreshAddress".lower(),
        "file_extensions": [".csv"]
    },
    {
        "name": "SendGrid Email Validation",
        "container": "incindio-prod",
        "directory": f"incindio-app-triggered-SendGrid-Validation-Action/{year}/{month}/",
        "keyword": "validation",
        "file_extensions": [".csv"]
    }
]

# ===== SUPABASE CONNECTION =====
def get_supabase_client() -> Client:
    """Get Supabase connection"""
    return create_client(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY)

# ===== ORGANIZATION LOOKUP =====
def load_organizations():
    """Load organizations into memory for quick lookup"""
    try:
        supabase = get_supabase_client()
        result = supabase.table('organizations').select('organization_id, name').execute()
        
        org_lookup = {org['organization_id']: org['name'] for org in result.data}
        print(f"Loaded {len(org_lookup)} organizations for lookup")
        return org_lookup
    except Exception as e:
        print(f"Warning: Could not load organizations: {e}")
        return {}

# ===== LOG PARSING FUNCTIONS =====
def parse_freshaddress_log(line):
    """Parse FreshAddress log line"""
    try:
        parts = line.split(',')
        if not parts:
            return None
        
        date = parts[0].strip()
        if 'T' in date:
            date = date.split('T')[0]
        
        config_match = re.search(r'Config:\s*([a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})', line, re.IGNORECASE)
        config = config_match.group(1).strip() if config_match else ""
        
        org_match = re.search(r'Org:\s*([a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})', line, re.IGNORECASE)
        org = org_match.group(1).strip() if org_match else ""
        
        email_match = re.search(r'Validation Request,\s*([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)', line)
        email = email_match.group(1).strip() if email_match else ""
        
        if date and org:
            return (date, config, org, email)
        
        return None
    except Exception as e:
        print(f"Error parsing FreshAddress line: {e}")
        return None

def parse_freshaddress_rt_log(line):
    """Parse FreshAddress RT log line"""
    try:
        parts = line.split(',')
        if not parts:
            return None
        
        date = parts[0].strip()
        if 'T' in date:
            date = date.split('T')[0]
        
        uuid_pattern = r'([a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})\s*"\s*,'
        uuid_match = re.search(uuid_pattern, line, re.IGNORECASE)
        
        if not uuid_match:
            return None
        
        org = uuid_match.group(1).strip()
        config = ""
        email = ""
        
        if date and org:
            return (date, config, org, email)
        
        return None
    except Exception as e:
        print(f"Error parsing FreshAddress RT line: {e}")
        return None

def parse_sendgrid_log(line):
    """Parse SendGrid log line"""
    try:
        timestamp_match = re.search(r'^(\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2})', line)
        
        if timestamp_match:
            raw_timestamp = timestamp_match.group(1)
            date = raw_timestamp.split('T')[0]
        else:
            date = line.split(',')[0].strip().split('T')[0]
        
        uuid_pattern = r'\|([a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})\s*"\s*,?'
        uuid_matches = re.findall(uuid_pattern, line, re.IGNORECASE)
        
        if not uuid_matches:
            return None
        
        org = uuid_matches[0].strip()
        config = ""
        
        email = ""
        if "Validation Request" in line:
            after_validation = line.split("Validation Request", 1)[1]
            email_match = re.search(
                r'([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)',
                after_validation
            )
            if email_match:
                email = email_match.group(1)
        
        if date and org:
            return (date, config, org, email)
        
        return None
    except Exception as e:
        print(f"Error parsing SendGrid line: {e}")
        return None

# ===== MAIN PROCESSING =====
def process_service(service_config, blob_service_client, org_lookup):
    """Process logs for a single service"""
    service_name = service_config["name"]
    container_name = service_config["container"]
    directory_path = service_config["directory"]
    keyword = service_config["keyword"]
    file_extensions = service_config["file_extensions"]
    
    print(f"\n{'='*60}")
    print(f"Processing: {service_name}")
    print(f"Container: {container_name}")
    print(f"Directory: {directory_path}")
    print(f"{'='*60}")
    
    try:
        container_client = blob_service_client.get_container_client(container_name)
    except Exception as e:
        print(f"Error connecting to container {container_name}: {e}")
        return []
    
    try:
        blobs = list(container_client.list_blobs(name_starts_with=directory_path))
        print(f"Found {len(blobs)} blobs")
    except Exception as e:
        print(f"Error listing blobs: {e}")
        return []
    
    if "FreshAddress Email" in service_name:
        parser = parse_freshaddress_log
    elif "FreshAddress RT" in service_name:
        parser = parse_freshaddress_rt_log
    else:
        parser = parse_sendgrid_log
    
    parsed_data = []
    
    for i, blob in enumerate(blobs):
        print(f"Processing blob {i+1}/{len(blobs)}: {blob.name}")
        
        if file_extensions and not any(blob.name.lower().endswith(ext) for ext in file_extensions):
            print(f"  Skipping (wrong extension)")
            continue
        
        try:
            blob_client = container_client.get_blob_client(blob.name)
            text = blob_client.download_blob().readall().decode("utf-8", errors="ignore")
            
            found_count = 0
            for line in text.splitlines():
                if keyword in line.lower():
                    result = parser(line.strip())
                    if result:
                        date, config, org, email = result
                        
                        try:
                            date_parts = date.split('-')
                            year_val = int(date_parts[0])
                            month_val = int(date_parts[1])
                        except (IndexError, ValueError):
                            print(f"  Warning: Could not parse date '{date}', skipping record")
                            continue
                        
                        org_name = org_lookup.get(org, "Unknown")
                        record = [date, year_val, month_val, service_name, org, org_name, config, email]
                        parsed_data.append(record)
                        found_count += 1
            
            if found_count > 0:
                print(f"  Found {found_count} matching lines")
            
        except Exception as e:
            print(f"  Error processing blob: {e}")
            continue
    
    return parsed_data

def aggregate_data(all_data):
    """Aggregate raw data by date, service, organization, and config"""
    from collections import defaultdict
    
    aggregated = defaultdict(int)
    
    for record in all_data:
        date, year_val, month_val, service, org_id, org_name, config, email = record
        key = (date, year_val, month_val, service, org_id, org_name, config)
        aggregated[key] += 1
    
    result = []
    for key, count in aggregated.items():
        date, year_val, month_val, service, org_id, org_name, config = key
        result.append({
            'date': date,
            'year': year_val,
            'month': month_val,
            'service': service,
            'organization_id': org_id,
            'organization_name': org_name,
            'config_id': config if config else None,
            'record_count': count
        })
    
    return result

def upload_to_supabase(aggregated_data):
    """Upload aggregated data to Supabase"""
    if not aggregated_data:
        print("No data to upload to Supabase")
        return
    
    try:
        supabase = get_supabase_client()
        
        # Insert in batches of 1000
        batch_size = 1000
        for i in range(0, len(aggregated_data), batch_size):
            batch = aggregated_data[i:i+batch_size]
            supabase.table('usage_summary').insert(batch).execute()
            print(f"Uploaded batch {i//batch_size + 1}: {len(batch)} records")
        
        print(f"✓ Successfully uploaded {len(aggregated_data)} aggregated records to Supabase")
        
    except Exception as e:
        print(f"Error uploading to Supabase: {e}")

def upload_to_google_drive(file_path):
    """Upload CSV file to Google Drive"""
    try:
        credentials = service_account.Credentials.from_service_account_file(
            GOOGLE_CREDENTIALS_PATH,
            scopes=['https://www.googleapis.com/auth/drive.file']
        )
        
        service = build('drive', 'v3', credentials=credentials)
        
        file_name = os.path.basename(file_path)
        
        # Check if file already exists in the folder
        query = f"name='{file_name}' and '{GOOGLE_DRIVE_FOLDER_ID}' in parents and trashed=false"
        results = service.files().list(q=query, fields="files(id, name)").execute()
        existing_files = results.get('files', [])
        
        if existing_files:
            # Update existing file
            file_id = existing_files[0]['id']
            media = MediaFileUpload(file_path, mimetype='text/csv')
            file = service.files().update(
                fileId=file_id,
                media_body=media,
                fields='id, name, webViewLink'
            ).execute()
            print(f"✓ Successfully updated existing file in Google Drive")
        else:
            # Create new file
            file_metadata = {
                'name': file_name,
                'parents': [GOOGLE_DRIVE_FOLDER_ID]
            }
            
            media = MediaFileUpload(file_path, mimetype='text/csv')
            
            file = service.files().create(
                body=file_metadata,
                media_body=media,
                fields='id, name, webViewLink',
                supportsAllDrives=True
            ).execute()
            print(f"✓ Successfully uploaded to Google Drive")
        
        print(f"  File ID: {file.get('id')}")
        print(f"  Link: {file.get('webViewLink')}")
        
    except Exception as e:
        print(f"Error uploading to Google Drive: {e}")

def main():
    """Main execution function"""
    print("=" * 60)
    print("Azure Blob Log Processing Script")
    print(f"Processing month: {year}-{month}")
    print("=" * 60)
    
    # Load organizations
    org_lookup = load_organizations()
    
    # Initialize Azure client
    print("\nInitializing Azure Blob client...")
    blob_service_client = BlobServiceClient.from_connection_string(AZURE_CONNECTION_STRING)
    
    # Process each service
    all_data = []
    for service_config in SERVICES:
        service_data = process_service(service_config, blob_service_client, org_lookup)
        all_data.extend(service_data)
    
    # Write combined CSV (detailed data)
    if all_data:
        try:
            with open(output_csv, "w", newline='', encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(["date", "year", "month", "service", "organization_id", "organization_name", "config_id", "email"])
                writer.writerows(all_data)
            print(f"\n✓ Successfully wrote {len(all_data)} total records to {output_csv}")
        except Exception as e:
            print(f"\nError writing CSV: {e}")
    else:
        print("\nNo matching data found across all services.")
        return
    
    # Aggregate data
    print("\nAggregating data...")
    aggregated_data = aggregate_data(all_data)
    print(f"Aggregated to {len(aggregated_data)} summary records")
    
    # Upload to Supabase
    print("\nUploading to Supabase...")
    upload_to_supabase(aggregated_data)
    
    # Upload CSV to Google Drive
    print("\nUploading CSV to Google Drive...")
    upload_to_google_drive(output_csv)
    
    # Summary
    print("\n" + "=" * 60)
    print("PROCESSING COMPLETE")
    print(f"Total raw records processed: {len(all_data):,}")
    print(f"Aggregated records in Supabase: {len(aggregated_data):,}")
    print(f"CSV output: {output_csv}")
    print("=" * 60)

if __name__ == "__main__":
    main()
