#!/usr/bin/env python3
"""
Zoho CRM Name Updater - Windows 11 Version with Zoho Sheets Integration
Updates contact names in Zoho CRM based on Zoho Sheets data (fetched via URL)

WINDOWS 11 REQUIREMENTS:
1. Python 3.8+ installed
2. Required packages: pip install requests
3. Zoho API credentials in %USERPROFILE%\.zoho_env file

USAGE:
python update_crm_names_windows.py --email test@example.com
python update_crm_names_windows.py --email test@example.com --name "Custom Name"
python update_crm_names_windows.py --all
python update_crm_names_windows.py --url "https://sheet.zoho.com/..." --all
"""

import os
import sys
import json
import time
import csv
import io
import requests
from typing import Dict, List, Optional
from pathlib import Path
from datetime import datetime
from urllib.parse import urlparse

# Windows-specific imports
import platform
if platform.system() == 'Windows':
    try:
        import winreg
    except ImportError:
        pass


# =============================================================================
# CONFIGURATION - Google Sheet URL (auto-fetches latest data when script runs)
# =============================================================================
DEFAULT_SHEET_URL = "https://docs.google.com/spreadsheets/d/1OpOrPoMq4NP0T09iSj8g5ZMZl7DfHufCruRLeN-Z6YU/export?format=csv"
# =============================================================================


class ZohoTokenManager:
    """Windows-compatible Zoho token manager with auto-refresh"""
    def __init__(self, env_file: str = '.zoho_env'):
        self.env_file = os.path.expanduser(f'~/{env_file}')
        self.load_credentials()
    
    def load_credentials(self):
        """Load credentials from environment file"""
        if not os.path.exists(self.env_file):
            self._show_setup_instructions()
            return
        
        print(f"📄 Loading credentials from: {self.env_file}")
        with open(self.env_file, 'r') as f:
            for line in f:
                if '=' in line and not line.strip().startswith('#'):
                    key, value = line.strip().split('=', 1)
                    os.environ[key] = value
        
        self.client_id = os.getenv('ZOHO_CLIENT_ID')
        self.client_secret = os.getenv('ZOHO_CLIENT_SECRET')
        self.refresh_token = os.getenv('ZOHO_REFRESH_TOKEN')
        self.api_domain = os.getenv('ZOHO_API_DOMAIN', 'https://www.zohoapis.com')
        self.access_token = os.getenv('ZOHO_ACCESS_TOKEN')
        self.token_expires_at = int(os.getenv('ZOHO_TOKEN_EXPIRES_AT', 0))
        
        if not all([self.client_id, self.client_secret, self.refresh_token]):
            self._show_setup_instructions()
    
    def _show_setup_instructions(self):
        """Show setup instructions"""
        print("="*70)
        print("🔧 SETUP REQUIRED")
        print("="*70)
        print("Missing Zoho API credentials. Please create a .zoho_env file:")
        print(f"📁 Location: {os.path.expanduser('~')}/.zoho_env")
        print()
        print("📝 Required contents:")
        print("ZOHO_CLIENT_ID=your_client_id")
        print("ZOHO_CLIENT_SECRET=your_client_secret")
        print("ZOHO_REFRESH_TOKEN=your_refresh_token")
        print("ZOHO_API_DOMAIN=https://www.zohoapis.com")
        print("="*70)
        sys.exit(1)
    
    def is_token_expired(self) -> bool:
        """Check if current access token is expired"""
        current_time = int(time.time())
        return current_time >= self.token_expires_at
    
    def refresh_access_token(self):
        """Refresh the access token using refresh token"""
        print("🔄 Refreshing access token...")
        url = "https://accounts.zoho.com/oauth/v2/token"
        data = {
            'refresh_token': self.refresh_token,
            'client_id': self.client_id,
            'client_secret': self.client_secret,
            'grant_type': 'refresh_token'
        }
        
        response = requests.post(url, data=data)
        
        if response.status_code == 200:
            token_data = response.json()
            self.access_token = token_data['access_token']
            
            # Calculate expiration time (current time + expires_in - 60 seconds buffer)
            expires_in = token_data.get('expires_in', 3600)
            self.token_expires_at = int(time.time()) + expires_in - 60
            
            # Update environment file
            self.update_env_file()
            
            print(f"✅ Token refreshed! Expires at: {datetime.fromtimestamp(self.token_expires_at)}")
            return token_data
        else:
            print(f"❌ Error refreshing token: {response.text}")
            raise Exception(f"Failed to refresh token: {response.text}")
    
    def update_env_file(self):
        """Update environment file with new token"""
        lines = []
        with open(self.env_file, 'r') as f:
            lines = f.readlines()
        
        token_updated = False
        expires_updated = False
        
        for i, line in enumerate(lines):
            if line.startswith('ZOHO_ACCESS_TOKEN='):
                lines[i] = f"ZOHO_ACCESS_TOKEN={self.access_token}\n"
                token_updated = True
            elif line.startswith('ZOHO_TOKEN_EXPIRES_AT='):
                lines[i] = f"ZOHO_TOKEN_EXPIRES_AT={self.token_expires_at}\n"
                expires_updated = True
        
        if not token_updated:
            lines.append(f"ZOHO_ACCESS_TOKEN={self.access_token}\n")
        if not expires_updated:
            lines.append(f"ZOHO_TOKEN_EXPIRES_AT={self.token_expires_at}\n")
        
        with open(self.env_file, 'w') as f:
            f.writelines(lines)
    
    def get_valid_token(self) -> str:
        """Get a valid access token, refresh if necessary"""
        if self.is_token_expired():
            print("⏰ Token expired, refreshing...")
            self.refresh_access_token()
        return self.access_token

    def make_api_call(self, endpoint, method='GET', data=None):
        """Make an API call to Zoho CRM with automatic token management"""
        token = self.get_valid_token()
        
        if not endpoint.startswith('/'):
            endpoint = '/' + endpoint
        
        url = f"{self.api_domain}{endpoint}"
        headers = {
            'Authorization': f'Zoho-oauthtoken {token}',
            'Content-Type': 'application/json'
        }
        
        try:
            if method.upper() == 'GET':
                response = requests.get(url, headers=headers)
            elif method.upper() == 'POST':
                response = requests.post(url, headers=headers, json=data)
            elif method.upper() == 'PUT':
                response = requests.put(url, headers=headers, json=data)
            elif method.upper() == 'DELETE':
                response = requests.delete(url, headers=headers)
            else:
                raise ValueError(f"Unsupported HTTP method: {method}")
            
            # Handle 401 - token might be invalid, try refreshing
            if response.status_code == 401:
                print("🔄 Received 401, refreshing token and retrying...")
                self.refresh_access_token()
                headers['Authorization'] = f'Zoho-oauthtoken {self.access_token}'
                
                if method.upper() == 'GET':
                    response = requests.get(url, headers=headers)
                elif method.upper() == 'POST':
                    response = requests.post(url, headers=headers, json=data)
                elif method.upper() == 'PUT':
                    response = requests.put(url, headers=headers, json=data)
                elif method.upper() == 'DELETE':
                    response = requests.delete(url, headers=headers)
            
            response.raise_for_status()
            
            if response.text:
                try:
                    return response.json()
                except json.JSONDecodeError:
                    if '/search' in endpoint:
                        return {'data': []}
                    raise Exception(f"Invalid JSON response from {method} {endpoint}")
            else:
                return {'data': []}
                
        except requests.exceptions.RequestException as e:
            raise Exception(f"API call failed for {method} {endpoint}: {str(e)}")


class ZohoSheetFetcher:
    """Fetches CSV data from Zoho Sheets"""
    
    def __init__(self, token_manager: ZohoTokenManager = None):
        self.token_manager = token_manager
    
    def fetch_from_url(self, url: str) -> str:
        """
        Fetch CSV data from a Google Sheets or Zoho Sheet URL
        
        Supports:
        - Google Sheets export URLs
        - Published/shared Zoho Sheet URLs
        """
        print(f"🌐 Fetching data from spreadsheet...")
        
        # Handle Google Sheets URLs
        if 'docs.google.com/spreadsheets' in url:
            # Convert edit URL to export URL if needed
            if '/edit' in url:
                sheet_id = url.split('/d/')[1].split('/')[0]
                url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
        # Handle Zoho Sheet URLs
        elif 'zoho' in url and 'output=csv' not in url and 'format=csv' not in url:
            separator = '&' if '?' in url else '?'
            url = f"{url}{separator}output=csv"
        
        print(f"   URL: {url[:70]}..." if len(url) > 70 else f"   URL: {url}")
        
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        }
        
        # If we have a token manager, add authorization for private sheets
        if self.token_manager and self.token_manager.access_token:
            headers['Authorization'] = f'Zoho-oauthtoken {self.token_manager.access_token}'
        
        try:
            response = requests.get(url, headers=headers, timeout=30)
            response.raise_for_status()
            
            # Check if we got CSV data
            content_type = response.headers.get('Content-Type', '')
            
            if 'text/csv' in content_type or 'text/plain' in content_type or response.text.strip().startswith('"') or ',' in response.text.split('\n')[0]:
                print(f"   ✅ Successfully fetched CSV data ({len(response.text)} bytes)")
                return response.text
            else:
                # Might be HTML page - try to extract CSV link or handle error
                if 'html' in content_type.lower():
                    raise Exception("Received HTML instead of CSV. Make sure the sheet is published with CSV format.")
                
                return response.text
                
        except requests.exceptions.Timeout:
            raise Exception("Request timed out. Check your internet connection.")
        except requests.exceptions.RequestException as e:
            raise Exception(f"Failed to fetch Zoho Sheet: {str(e)}")
    
    def parse_csv_string(self, csv_string: str) -> Dict[str, Dict]:
        """Parse CSV string into vendor data dictionary"""
        vendor_data = {}
        
        # Use StringIO to read CSV from string
        csv_file = io.StringIO(csv_string)
        
        # Handle BOM if present
        content = csv_file.read()
        if content.startswith('\ufeff'):
            content = content[1:]
        csv_file = io.StringIO(content)
        
        reader = csv.DictReader(csv_file)
        
        # Validate required columns
        required_columns = ['Contact email', 'Name']
        if reader.fieldnames:
            missing = [col for col in required_columns if col not in reader.fieldnames]
            if missing:
                raise Exception(f"Missing required columns: {missing}. Found: {reader.fieldnames}")
        
        row_count = 0
        for row in reader:
            # Use contact email as key for lookup
            email = row.get('Contact email', '').strip().lower()
            if email:
                vendor_data[email] = {
                    'name': row.get('Name', '').strip(),
                    'nickname': row.get('Nickname', '').strip(),
                    'contact_email': row.get('Contact email', '').strip(),
                    'payment_emails': row.get('Emails for payment receipts', '').strip()
                }
                row_count += 1
        
        print(f"   📊 Parsed {row_count} vendor records")
        return vendor_data


class CRMNameUpdater:
    def __init__(self, data_source: str, is_url: bool = False):
        """
        Initialize the CRM Name Updater
        
        Args:
            data_source: Either a file path or a Zoho Sheet URL
            is_url: True if data_source is a URL, False if it's a file path
        """
        self.manager = ZohoTokenManager()
        self.data_source = data_source
        self.is_url = is_url
        self.module = "Contacts"
        
        # Load vendor data from CSV (file or URL)
        self.vendor_data = self._load_data()
        
        # Store results
        self.results = {
            'total_vendors': len(self.vendor_data),
            'contacts_found': 0,
            'names_updated': 0,
            'already_correct': 0,
            'errors': [],
            'processed_contacts': []
        }
        
        print("🚀 Zoho CRM Name Updater Initialized (Windows 11)")
        print(f"   📄 Data source: {'URL' if is_url else 'Local file'}")
        print(f"   📋 Module: {self.module}")
        print(f"   👥 Vendors loaded: {len(self.vendor_data)}")
        print(f"   💻 Platform: {platform.system()} {platform.release()}")
        print("")
    
    def _load_data(self) -> Dict[str, Dict]:
        """Load vendor data from CSV file or Zoho Sheet URL"""
        if self.is_url:
            return self._load_from_url()
        else:
            return self._load_from_file()
    
    def _load_from_url(self) -> Dict[str, Dict]:
        """Load vendor data from Zoho Sheet URL"""
        fetcher = ZohoSheetFetcher(self.manager)
        csv_string = fetcher.fetch_from_url(self.data_source)
        return fetcher.parse_csv_string(csv_string)
    
    def _load_from_file(self) -> Dict[str, Dict]:
        """Load vendor data from local CSV file"""
        vendor_data = {}
        
        # Convert to absolute path for Windows compatibility
        csv_path = os.path.abspath(self.data_source)
        
        try:
            with open(csv_path, 'r', encoding='utf-8-sig') as csvfile:
                reader = csv.DictReader(csvfile)
                for row in reader:
                    # Use contact email as key for lookup
                    email = row['Contact email'].strip().lower()
                    vendor_data[email] = {
                        'name': row['Name'].strip(),
                        'nickname': row.get('Nickname', '').strip(),
                        'contact_email': row['Contact email'].strip(),
                        'payment_emails': row.get('Emails for payment receipts', '').strip()
                    }
        except FileNotFoundError:
            print(f"❌ CSV file not found: {csv_path}")
            print("💡 Make sure the CSV file is in the same directory as this script")
            print("💡 Or use --url to fetch from Zoho Sheets directly")
            sys.exit(1)
        except Exception as e:
            print(f"❌ Error reading CSV file: {str(e)}")
            sys.exit(1)
        
        return vendor_data
    
    def search_contact_by_email(self, email: str) -> Optional[Dict]:
        """Search for a contact by email address"""
        print(f"🔍 Searching for contact with email: {email}")
        
        try:
            # Use search API with email criteria
            search_criteria = f"(Email:equals:{email})"
            endpoint = f'/crm/v2/{self.module}/search?criteria={search_criteria}'
            
            response = self.manager.make_api_call(endpoint)
            
            if response.get('data') and len(response['data']) > 0:
                contact = response['data'][0]
                print(f"   ✅ Found contact: {contact.get('Full_Name', 'Unknown')} (ID: {contact['id']})")
                return contact
            else:
                print(f"   ❌ No contact found with email: {email}")
                return None
                
        except Exception as e:
            error_msg = f"Error searching for email {email}: {str(e)}"
            print(f"   ❌ {error_msg}")
            self.results['errors'].append(error_msg)
            return None
    
    def update_contact_name(self, contact_id: str, current_name: str, new_name: str) -> bool:
        """Update a contact's name in Zoho CRM"""
        try:
            if current_name == new_name:
                print(f"   ⚠️  Name is already correct: {current_name}")
                self.results['already_correct'] += 1
                return True
            
            # Split new name into first and last name
            name_parts = new_name.strip().split()
            if len(name_parts) == 1:
                first_name = name_parts[0]
                last_name = ""
            else:
                first_name = name_parts[0]
                last_name = " ".join(name_parts[1:])
            
            print(f"   📝 Updating: First='{first_name}', Last='{last_name}'")
            
            data = {
                "data": [{
                    "id": contact_id,
                    "First_Name": first_name,
                    "Last_Name": last_name
                }]
            }
            
            endpoint = f'/crm/v2/{self.module}'
            response = self.manager.make_api_call(endpoint, 'PUT', data)
            
            if response.get('data') and len(response['data']) > 0:
                if response['data'][0].get('code') == 'SUCCESS':
                    print(f"   ✅ Successfully updated name: '{current_name}' → '{new_name}'")
                    self.results['names_updated'] += 1
                    return True
            
            print(f"   ❌ Failed to update name for contact {contact_id}")
            return False
            
        except Exception as e:
            error_msg = f"Error updating name for contact {contact_id}: {str(e)}"
            print(f"   ❌ {error_msg}")
            self.results['errors'].append(error_msg)
            return False
    
    def process_email(self, email: str, vendor_info: Dict) -> Dict:
        """Process a single email address and update the contact name"""
        result = {
            'email': email,
            'vendor_name': vendor_info['name'],
            'status': 'not_found',
            'current_name': None,
            'updated_name': None,
            'contact_id': None
        }
        
        contact = self.search_contact_by_email(email)
        
        if contact:
            current_name = contact.get('Full_Name', 'Unknown')
            result['current_name'] = current_name
            result['contact_id'] = contact['id']
            self.results['contacts_found'] += 1
            
            new_name = vendor_info['name']
            result['updated_name'] = new_name
            
            if self.update_contact_name(contact['id'], current_name, new_name):
                result['status'] = 'updated' if current_name != new_name else 'already_correct'
            else:
                result['status'] = 'error'
        
        return result
    
    def process_single_email(self, email: str, custom_name: str = None):
        """Process a single email address for testing"""
        email_key = email.strip().lower()
        
        if custom_name:
            vendor_info = {
                'name': custom_name,
                'nickname': custom_name,
                'contact_email': email,
                'payment_emails': email
            }
            print(f"📊 Processing email: {email}")
            print(f"   📝 Custom name: {custom_name}")
        else:
            if email_key not in self.vendor_data:
                print(f"❌ Email {email} not found in vendor data")
                return
            
            vendor_info = self.vendor_data[email_key]
            print(f"📊 Processing email: {email}")
            print(f"   📝 Vendor name: {vendor_info['name']}")
            print(f"   🏷️  Vendor nickname: {vendor_info['nickname']}")
        
        print("-" * 50)
        
        result = self.process_email(email, vendor_info)
        self.results['processed_contacts'].append(result)
        
        print(f"\n📋 Result for {email}:")
        print(f"   Status: {result['status']}")
        if result['current_name']:
            print(f"   Current name: {result['current_name']}")
        if result['updated_name']:
            print(f"   Updated name: {result['updated_name']}")
        
        self.print_summary()
        self.save_csv_with_status()
    
    def process_all_vendors(self):
        """Process all vendors from the data source"""
        print(f"\n📊 Processing {len(self.vendor_data)} vendor(s)...\n")
        
        for i, (email, vendor_info) in enumerate(self.vendor_data.items(), 1):
            print(f"\n[{i}/{len(self.vendor_data)}] Processing: {email}")
            print(f"   📝 Vendor: {vendor_info['name']}")
            print("-" * 50)
            
            result = self.process_email(email, vendor_info)
            self.results['processed_contacts'].append(result)
            
            # Add delay to avoid rate limiting
            if i < len(self.vendor_data):
                time.sleep(0.5)
            
            # Save progress every 10 contacts
            if i % 10 == 0:
                print(f"\n💾 Saving progress after {i} contacts...")
                self.save_csv_with_status()
        
        self.print_summary()
        self.save_csv_with_status()
    
    def print_summary(self):
        """Print a summary of the processing results"""
        print("\n" + "="*70)
        print("📊 PROCESSING SUMMARY")
        print("="*70)
        print(f"Total vendors:             {self.results['total_vendors']}")
        print(f"Contacts found:            {self.results['contacts_found']}")
        print(f"Names updated:             {self.results['names_updated']}")
        print(f"Already correct:           {self.results['already_correct']}")
        print(f"Errors:                    {len(self.results['errors'])}")
        
        if self.results['errors']:
            print("\n❌ ERRORS:")
            for error in self.results['errors']:
                print(f"   - {error}")
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        results_file = f'name_update_results_{timestamp}.json'
        results_path = os.path.abspath(results_file)
        
        try:
            with open(results_path, 'w', encoding='utf-8') as f:
                json.dump(self.results, f, indent=2, ensure_ascii=False)
            print(f"\n📄 Detailed results saved to: {results_path}")
        except Exception as e:
            print(f"\n❌ Error saving results file: {e}")
        
        print("="*70)
    
    def save_csv_with_status(self):
        """Save an updated CSV file with status column"""
        if not self.results['processed_contacts']:
            return
        
        status_lookup = {}
        for contact in self.results['processed_contacts']:
            email = contact['email'].lower()
            status_lookup[email] = {
                'status': contact['status'],
                'current_name': contact.get('current_name', ''),
                'updated_name': contact.get('updated_name', '')
            }
        
        # Create output CSV from processed data
        output_rows = []
        fieldnames = ['Contact email', 'Name', 'Nickname', 'Update_Status', 'Previous_Name', 'Update_Timestamp']
        
        for email, vendor_info in self.vendor_data.items():
            row = {
                'Contact email': vendor_info['contact_email'],
                'Name': vendor_info['name'],
                'Nickname': vendor_info.get('nickname', '')
            }
            
            if email in status_lookup:
                contact_info = status_lookup[email]
                if contact_info['status'] == 'updated':
                    row['Update_Status'] = 'UPDATED'
                    row['Previous_Name'] = contact_info['current_name']
                elif contact_info['status'] == 'already_correct':
                    row['Update_Status'] = 'NO_CHANGE_NEEDED'
                    row['Previous_Name'] = ''
                elif contact_info['status'] == 'not_found':
                    row['Update_Status'] = 'NOT_FOUND_IN_CRM'
                    row['Previous_Name'] = ''
                else:
                    row['Update_Status'] = 'ERROR'
                    row['Previous_Name'] = contact_info['current_name']
                row['Update_Timestamp'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            else:
                row['Update_Status'] = 'NOT_PROCESSED'
                row['Previous_Name'] = ''
                row['Update_Timestamp'] = ''
            
            output_rows.append(row)
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_file = f'vendors_with_status_{timestamp}.csv'
        output_path = os.path.abspath(output_file)
        
        try:
            with open(output_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(output_rows)
            
            print(f"📊 Updated CSV saved to: {output_path}")
        except Exception as e:
            print(f"❌ Error saving updated CSV: {str(e)}")


def main():
    """Main entry point"""
    import argparse
    
    if sys.version_info < (3, 8):
        print("❌ This script requires Python 3.8 or higher")
        print(f"   Current version: {sys.version}")
        sys.exit(1)
    
    try:
        import requests
    except ImportError:
        print("❌ Required package 'requests' not found")
        print("💡 Install it with: pip install requests")
        sys.exit(1)
    
    parser = argparse.ArgumentParser(
        description='Update Zoho CRM contact names from Zoho Sheets or CSV file',
        epilog="""
Examples:
  # Fetch from Zoho Sheet URL (always gets latest data):
  python update_crm_names_windows.py --url "https://sheet.zoho.com/..." --all
  python update_crm_names_windows.py --url "https://sheet.zoho.com/..." --email test@example.com
  
  # Use local CSV file:
  python update_crm_names_windows.py --csv "vendors.csv" --all
  python update_crm_names_windows.py --email test@example.com --name "Custom Name"
        """,
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    
    # Data source options
    source_group = parser.add_mutually_exclusive_group()
    source_group.add_argument('--url', help='Zoho Sheet URL (CSV export format)')
    source_group.add_argument('--csv', default='Vendors paid in 2025.csv', 
                              help='Local CSV file (default: Vendors paid in 2025.csv)')
    
    # Action options
    parser.add_argument('--email', help='Single email address to process')
    parser.add_argument('--name', help='Custom name to set (use with --email)')
    parser.add_argument('--all', action='store_true', help='Process all vendors')
    
    args = parser.parse_args()
    
    print("="*70)
    print("🪟 ZOHO CRM NAME UPDATER - WINDOWS 11 EDITION")
    print("="*70)
    print(f"🐍 Python: {sys.version.split()[0]}")
    print(f"💻 Platform: {platform.system()} {platform.release()}")
    print(f"📁 Working Directory: {os.getcwd()}")
    print("="*70)
    
    # Determine data source
    if args.url:
        data_source = args.url
        is_url = True
        print(f"🌐 Data source: Zoho Sheet URL")
    elif DEFAULT_SHEET_URL:
        data_source = DEFAULT_SHEET_URL
        is_url = True
        print(f"🌐 Data source: Default Sheet URL")
    else:
        data_source = args.csv
        is_url = False
        print(f"📄 Data source: Local file ({data_source})")
    
    print("="*70)
    
    try:
        updater = CRMNameUpdater(data_source, is_url=is_url)
    except Exception as e:
        print(f"❌ Failed to initialize: {e}")
        print("\n💡 Common solutions:")
        if is_url:
            print("- Check that the Zoho Sheet URL is correct and accessible")
            print("- Make sure the sheet is published/shared with CSV format")
        else:
            print("- Check that the CSV file exists")
        print("- Verify the .zoho_env file has correct API credentials")
        print("- Make sure you have internet access")
        sys.exit(1)
    
    try:
        if args.email:
            updater.process_single_email(args.email, args.name)
        elif args.all:
            confirm = input("\n⚠️  Process ALL vendors? This may take time. Type 'YES' to continue: ")
            if confirm.upper() == 'YES':
                updater.process_all_vendors()
            else:
                print("❌ Operation cancelled")
        else:
            print("❌ No action specified. Use --help for usage information.")
            print("\n📖 Quick Examples:")
            print('  python update_crm_names_windows.py --url "https://sheet.zoho.com/..." --all')
            print("  python update_crm_names_windows.py --email test@example.com")
    except KeyboardInterrupt:
        print("\n⚠️  Process interrupted by user")
    except Exception as e:
        print(f"\n❌ Error during processing: {e}")


if __name__ == "__main__":
    main()
