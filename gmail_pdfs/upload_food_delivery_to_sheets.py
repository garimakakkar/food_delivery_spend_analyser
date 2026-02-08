#!/usr/bin/env python3
"""
Upload Zomato order data to Google Sheets
"""

import gspread
from google.oauth2.service_account import Credentials
import csv
import os
import sys

def upload_to_google_sheets(csv_path: str, sheet_name: str = "Zomato Orders", credentials_path: str = None):
    """Upload CSV data to Google Sheets"""
    
    # Check for credentials file
    if credentials_path is None:
        # Look for common credential file names
        possible_paths = [
            "credentials.json",
            "service_account.json",
            "google_credentials.json",
            os.path.expanduser("~/.config/gspread/service_account.json")
        ]
        
        for path in possible_paths:
            if os.path.exists(path):
                credentials_path = path
                break
    
    if not credentials_path or not os.path.exists(credentials_path):
        print("❌ Google Sheets credentials not found!")
        print("\nTo upload to Google Sheets, you need to:")
        print("1. Go to https://console.cloud.google.com/")
        print("2. Create a new project or select an existing one")
        print("3. Enable Google Sheets API")
        print("4. Create a Service Account")
        print("5. Download the JSON credentials file")
        print("6. Save it as 'credentials.json' in this directory")
        print("\nAlternatively, you can manually import the CSV file:")
        print(f"   CSV file location: {csv_path}")
        print("   1. Open Google Sheets")
        print("   2. File > Import > Upload")
        print("   3. Select the CSV file")
        return False
    
    try:
        # Authenticate
        scope = ['https://spreadsheets.google.com/feeds',
                 'https://www.googleapis.com/auth/drive']
        creds = Credentials.from_service_account_file(credentials_path, scopes=scope)
        client = gspread.authorize(creds)
        
        # Read CSV data
        data = []
        with open(csv_path, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            headers = reader.fieldnames
            for row in reader:
                data.append([row.get(h, '') for h in headers])
        
        # Create or open spreadsheet
        try:
            spreadsheet = client.create(sheet_name)
            print(f"✅ Created new spreadsheet: {sheet_name}")
        except Exception as e:
            # Try to open existing spreadsheet
            try:
                spreadsheet = client.open(sheet_name)
                print(f"✅ Opened existing spreadsheet: {sheet_name}")
            except:
                print(f"❌ Error accessing spreadsheet: {str(e)}")
                return False
        
        # Get the first worksheet
        worksheet = spreadsheet.sheet1
        
        # Clear existing data
        worksheet.clear()
        
        # Add headers
        worksheet.append_row(headers)
        
        # Add data rows
        if data:
            worksheet.append_rows(data)
        
        # Format header row
        worksheet.format('1:1', {
            'textFormat': {'bold': True},
            'backgroundColor': {'red': 0.2, 'green': 0.6, 'blue': 0.9}
        })
        
        # Auto-resize columns
        try:
            worksheet.columns_auto_resize(0, len(headers))
        except:
            pass
        
        # Get the URL
        url = spreadsheet.url
        print(f"\n✅ Successfully uploaded data to Google Sheets!")
        print(f"📊 Spreadsheet URL: {url}")
        print(f"📝 Sheet name: {sheet_name}")
        print(f"📈 Rows uploaded: {len(data)}")
        
        return True
        
    except Exception as e:
        print(f"❌ Error uploading to Google Sheets: {str(e)}")
        print(f"\nYou can manually import the CSV file:")
        print(f"   CSV file location: {csv_path}")
        return False

if __name__ == "__main__":
    csv_path = "/Users/garima/Desktop/zomato/gmail_pdfs/zomato_orders.csv"
    
    if not os.path.exists(csv_path):
        print(f"❌ CSV file not found: {csv_path}")
        sys.exit(1)
    
    # Check for credentials file as command line argument
    credentials_path = sys.argv[1] if len(sys.argv) > 1 else None
    
    upload_to_google_sheets(csv_path, credentials_path=credentials_path)
