import os
from twilio.rest import Client
from datetime import datetime, timedelta
from dotenv import load_dotenv
from supabase import create_client, Client as SupabaseClient

# Load environment variables from .env file
load_dotenv()

# Twilio credentials
ACCOUNT_SID = os.environ.get('TWILIO_ACCOUNT_SID')
AUTH_TOKEN = os.environ.get('TWILIO_AUTH_TOKEN')

# Supabase credentials
SUPABASE_URL = os.environ.get('SUPABASE_URL')
SUPABASE_KEY = os.environ.get('SUPABASE_SERVICE_ROLE_KEY')

def get_supabase_client():
    """Initialize Supabase client"""
    print(f"Connecting to Supabase...")
    print(f"  URL: {SUPABASE_URL}")
    print(f"  Key: {'*' * 20}{SUPABASE_KEY[-4:] if SUPABASE_KEY else 'NOT SET'}")
    return create_client(SUPABASE_URL, SUPABASE_KEY)

def get_organization_lookup(supabase):
    """Fetch organizations from Supabase and create a lookup dictionary"""
    try:
        print("  Querying twilio_organizations table...")
        response = supabase.table('twilio_organizations').select('twilio_sid, organization_id').execute()
        
        print(f"  Response data: {response.data}")
        
        # Create lookup dictionary: twilio_sid -> organization_id
        org_lookup = {}
        for org in response.data:
            org_lookup[org['twilio_sid']] = org['organization_id']
        
        print(f"Loaded {len(org_lookup)} Twilio organizations from Supabase")
        return org_lookup
    except Exception as e:
        print(f"Error fetching organizations: {e}")
        import traceback
        traceback.print_exc()
        return {}

def get_organizations_data(supabase):
    """Fetch full organization data from Supabase"""
    try:
        response = supabase.table('organizations').select('organization_id, name, platform').execute()
        
        # Create lookup dictionary: organization_id -> {name, platform}
        orgs_data = {}
        for org in response.data:
            orgs_data[org['organization_id']] = {
                'name': org['name'],
                'platform': org['platform']
            }
        
        print(f"Loaded {len(orgs_data)} organizations data from Supabase")
        return orgs_data
    except Exception as e:
        print(f"Error fetching organizations data: {e}")
        return {}

def get_service_lookup(supabase):
    """Fetch Twilio services from Supabase and create a lookup dictionary"""
    try:
        response = supabase.table('twilio_service').select('twilio_category, service_name').execute()
        
        # Create lookup dictionary: twilio_category -> service_name
        service_lookup = {}
        for service in response.data:
            service_lookup[service['twilio_category']] = service['service_name']
        
        print(f"Loaded {len(service_lookup)} Twilio services from Supabase")
        return service_lookup
    except Exception as e:
        print(f"Error fetching Twilio services: {e}")
        return {}

def get_subaccounts(client):
    """Retrieve only subaccounts under the main account"""
    subaccounts = []
    try:
        print(f"  Main Account SID: {ACCOUNT_SID}")
        print(f"  Checking for subaccounts...\n")
        
        # Fetch accounts
        accounts = client.api.accounts.list()
        
        for account in accounts:
            # Skip the main account itself
            if account.sid == ACCOUNT_SID:
                print(f"  ⊙ Main Account: {account.friendly_name} ({account.sid})")
                continue
            
            # Check owner_account_sid
            owner_sid = getattr(account, 'owner_account_sid', None)
            
            print(f"  Account: {account.friendly_name}")
            print(f"    SID: {account.sid}")
            print(f"    Owner SID: {owner_sid}")
            print(f"    Status: {account.status}")
            
            # Only add if it's actually owned by our main account AND active
            if owner_sid == ACCOUNT_SID and account.status.lower() == 'active':
                subaccounts.append({
                    'sid': account.sid,
                    'friendly_name': account.friendly_name,
                    'status': account.status
                })
                print(f"    ✓ ADDED as subaccount\n")
            elif owner_sid == ACCOUNT_SID and account.status.lower() != 'active':
                print(f"    ✗ SKIPPED (status: {account.status} - only active accounts included)\n")
            else:
                print(f"    ✗ SKIPPED (owner {owner_sid} != main account {ACCOUNT_SID})\n")
                
    except Exception as e:
        print(f"Error fetching subaccounts: {e}")
    return subaccounts

def get_usage_records(client, account_sid, start_date=None, end_date=None):
    """Get usage records for an account"""
    usage_data = []
    
    # Categories to include
    ALLOWED_CATEGORIES = [
        'sms-inbound',
        'sms-outbound',
        'mms-inbound',
        'mms-outbound',
        'calls-inbound',
        'calls-outbound',
        'number-format-lookups'
        'carrier-lookups'
    ]
    
    try:
        # If no dates provided, get previous month
        if not start_date:
            today = datetime.now()
            # Get first day of current month, then go back one day to get last month
            first_of_current_month = today.replace(day=1)
            last_day_prev_month = first_of_current_month.replace(day=1) - timedelta(days=1)
            first_day_prev_month = last_day_prev_month.replace(day=1)
            
            start_date = first_day_prev_month.strftime('%Y-%m-%d')
        if not end_date:
            today = datetime.now()
            first_of_current_month = today.replace(day=1)
            last_day_prev_month = first_of_current_month - timedelta(days=1)
            end_date = last_day_prev_month.strftime('%Y-%m-%d')
        
        # Get usage records
        records = client.usage.records.list(
            start_date=start_date,
            end_date=end_date
        )
        
        for record in records:
            # Filter: only include specified categories and non-zero usage
            if record.category in ALLOWED_CATEGORIES and float(record.usage) != 0:
                usage_data.append({
                    'account_sid': account_sid,
                    'category': record.category,
                    'description': record.description,
                    'count': record.count,
                    'count_unit': record.count_unit,
                    'usage': record.usage,
                    'usage_unit': record.usage_unit,
                    'price': record.price,
                    'price_unit': record.price_unit,
                    'start_date': str(record.start_date),
                    'end_date': str(record.end_date)
                })
    except Exception as e:
        print(f"Error fetching usage for {account_sid}: {e}")
    
    return usage_data

def transform_for_supabase(usage_data, org_lookup, orgs_data, service_lookup):
    """Transform usage data to match Supabase schema"""
    transformed_data = []
    current_date = datetime.now().date().isoformat()
    current_time = datetime.now().isoformat()
    
    for record in usage_data:
        account_sid = record['account_sid']
        category = record['category']
        
        # Lookup organization_id from twilio_organizations
        organization_id = org_lookup.get(account_sid)
        
        if not organization_id:
            print(f"Warning: No organization found for account_sid {account_sid}, skipping record")
            continue
        
        # Get organization name and platform from organizations table
        org_data = orgs_data.get(organization_id)
        
        if not org_data:
            print(f"Warning: No organization data found for organization_id {organization_id}, skipping record")
            continue
        
        # Lookup service name using category
        service_name = service_lookup.get(category)
        
        if not service_name:
            print(f"Warning: No service found for category '{category}', skipping record")
            continue
        
        # Parse start_date to extract year and month
        start_date = datetime.strptime(record['start_date'], '%Y-%m-%d')
        
        transformed_record = {
            'date': current_date,
            'year': start_date.strftime('%Y'),
            'month': start_date.strftime('%m'),
            'service': service_name,
            'organization_id': organization_id,
            'organization_name': org_data['name'],
            'config_id': None,
            'record_count': int(float(record['usage'])),
            'created_at': current_time,
            'platform': org_data['platform']
        }
        
        transformed_data.append(transformed_record)
    
    return transformed_data

def insert_to_supabase(supabase, data):
    """Insert data into Supabase usage_summary table"""
    if not data:
        print("No data to insert into Supabase")
        return
    
    try:
        response = supabase.table('usage_summary').insert(data).execute()
        print(f"✓ Successfully inserted {len(data)} records into Supabase")
        return response
    except Exception as e:
        print(f"✗ Error inserting data into Supabase: {e}")
        return None

def main():
    # Initialize clients
    twilio_client = Client(ACCOUNT_SID, AUTH_TOKEN)
    supabase = get_supabase_client()
    
    # Calculate previous month dates
    today = datetime.now()
    first_of_current_month = today.replace(day=1)
    last_day_prev_month = first_of_current_month - timedelta(days=1)
    first_day_prev_month = last_day_prev_month.replace(day=1)
    
    month_year = last_day_prev_month.strftime('%B %Y')
    print(f"Fetching billing data for {month_year}...")
    print(f"Date range: {first_day_prev_month.strftime('%Y-%m-%d')} to {last_day_prev_month.strftime('%Y-%m-%d')}")
    
    # Get organization lookup from Supabase
    print("\nFetching Twilio organizations from Supabase...")
    org_lookup = get_organization_lookup(supabase)
    
    if not org_lookup:
        print("Error: Could not load Twilio organizations. Exiting.")
        return
    
    # Get organizations data from Supabase
    print("Fetching organizations data from Supabase...")
    orgs_data = get_organizations_data(supabase)
    
    if not orgs_data:
        print("Error: Could not load organizations data. Exiting.")
        return
    
    # Get service lookup from Supabase
    print("Fetching Twilio services from Supabase...")
    service_lookup = get_service_lookup(supabase)
    
    if not service_lookup:
        print("Error: Could not load Twilio services. Exiting.")
        return
    
    print("\nFetching data for main account...")
    
    # Get main account data
    main_usage = get_usage_records(twilio_client, ACCOUNT_SID)
    
    all_usage_data = main_usage
    
    # Get subaccounts
    print("Fetching subaccounts...")
    subaccounts = get_subaccounts(twilio_client)
    print(f"Found {len(subaccounts)} subaccounts")
    
    # Categories to include
    ALLOWED_CATEGORIES = [
        'sms-inbound',
        'sms-outbound',
        'mms-inbound',
        'mms-outbound',
        'calls-inbound',
        'calls-outbound',
        'number-format-lookups',
        'carrier-lookups'
    ]
    
    # Get data for each subaccount
    for subaccount in subaccounts:
        print(f"Fetching data for subaccount: {subaccount['friendly_name']}")
        
        try:
            sub_usage = []
            records = twilio_client.api.v2010.accounts.get(subaccount['sid']).usage.records.list(
                start_date=first_day_prev_month.strftime('%Y-%m-%d'),
                end_date=last_day_prev_month.strftime('%Y-%m-%d')
            )
            
            for record in records:
                # Filter: only include specified categories and non-zero usage
                if record.category in ALLOWED_CATEGORIES and float(record.usage) != 0:
                    sub_usage.append({
                        'account_sid': subaccount['sid'],
                        'category': record.category,
                        'description': record.description,
                        'count': record.count,
                        'count_unit': record.count_unit,
                        'usage': record.usage,
                        'usage_unit': record.usage_unit,
                        'price': record.price,
                        'price_unit': record.price_unit,
                        'start_date': str(record.start_date),
                        'end_date': str(record.end_date)
                    })
            
            all_usage_data.extend(sub_usage)
            
        except Exception as e:
            print(f"Error fetching data for subaccount {subaccount['friendly_name']}: {e}")
    
    print(f"\nTotal usage records fetched from Twilio: {len(all_usage_data)}")
    
    # Transform data for Supabase
    print("\nTransforming data for Supabase...")
    transformed_data = transform_for_supabase(all_usage_data, org_lookup, orgs_data, service_lookup)
    
    # Insert into Supabase
    print("\nInserting data into Supabase...")
    insert_to_supabase(supabase, transformed_data)
    
    print("\nBilling data extraction and upload complete!")
    print(f"Records processed: {len(transformed_data)}")

if __name__ == "__main__":
    if not ACCOUNT_SID or not AUTH_TOKEN:
        print("Error: Please set TWILIO_ACCOUNT_SID and TWILIO_AUTH_TOKEN environment variables")
    elif not SUPABASE_URL or not SUPABASE_KEY:
        print("Error: Please set SUPABASE_URL and SUPABASE_KEY environment variables")
    else:
        main()
