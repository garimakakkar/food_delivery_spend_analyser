import os
import re
import pdfplumber
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# Scopes for Google Sheets
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def authenticate_sheets():
    """Authenticate and return Google Sheets service"""
    creds = None
    
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    
    return build('sheets', 'v4', credentials=creds)

def extract_invoice_data(pdf_path):
    """Extract data from Zomato invoice PDF"""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = ""
            for page in pdf.pages:
                text += page.extract_text() or ""
        
        invoice_data = {
            'filename': os.path.basename(pdf_path),
            'date': '',
            'restaurant': '',
            'amount': '',
            'items': []
        }
        
        # Extract date - looking for "Order Time: 28 January 2026"
        date_patterns = [
            r'Order Time:\s*(\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{4})',
            r'(\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{4})',
            r'(\d{1,2}[/-]\d{1,2}[/-]\d{4})',
        ]
        for pattern in date_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                invoice_data['date'] = match.group(1).strip()
                break
        
        # Extract restaurant name - looking for "Restaurant Name: P.F. Chang's"
        restaurant_patterns = [
            r'Restaurant Name:\s*([^\n]+)',
            r'Ordered from[:\s]+([^\n]+)',
            r'Delivery from[:\s]+([^\n]+)',
        ]
        for pattern in restaurant_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                restaurant = match.group(1).strip()
                restaurant = re.sub(r'\s*\(.*?\)\s*', '', restaurant)
                invoice_data['restaurant'] = restaurant[:100]
                break
        
        # Extract total amount - FIXED to match "Total ₹1,262.53" format exactly
        # The key is to match the word "Total" (not "subtotal") followed by the rupee amount
        total_patterns = [
            # Match "Total" at the start of a line or after whitespace, followed by rupee amount
            r'(?:^|\n)\s*Total\s+₹?([\d,]+\.?\d{0,2})',
            # Backup: Match "Total" with colon
            r'Total:\s*₹?([\d,]+\.?\d{0,2})',
        ]
        
        for pattern in total_patterns:
            match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
            if match:
                # Remove commas from the amount
                amount_str = match.group(1).replace(',', '')
                invoice_data['amount'] = f"₹{amount_str}"
                break
        
        # Extract items - looking for lines like "Burnt Chilli Hakka Noodles 1 ₹495 ₹495"
        item_patterns = [
            # Pattern: Item name, quantity, unit price, total price
            r'^([A-Z][^\n\d]+?)\s+(\d+)\s+₹([\d,]+)\s+₹([\d,]+)',
        ]
        
        items = []
        lines = text.split('\n')
        for line in lines:
            for pattern in item_patterns:
                match = re.match(pattern, line.strip(), re.IGNORECASE)
                if match:
                    try:
                        item_name = match.group(1).strip()
                        qty = match.group(2)
                        price = match.group(4).replace(',', '')
                        
                        # Skip lines that look like headers or subtotals
                        if len(item_name) > 3 and item_name not in ['Taxes', 'Total', 'Item']:
                            items.append(f"{qty}x {item_name} (₹{price})")
                    except:
                        continue
        
        invoice_data['items'] = items[:10]
        
        return invoice_data
        
    except Exception as e:
        print(f"  ⚠ Error parsing {pdf_path}: {e}")
        return None

def create_spreadsheet(service, sheet_name='Zomato Invoices'):
    """Create new Google Sheet"""
    spreadsheet = {
        'properties': {
            'title': sheet_name
        }
    }
    
    result = service.spreadsheets().create(body=spreadsheet).execute()
    spreadsheet_id = result['spreadsheetId']
    
    print(f"✓ Created new sheet: {sheet_name}")
    print(f"✓ Sheet URL: https://docs.google.com/spreadsheets/d/{spreadsheet_id}")
    
    # Add headers
    headers = [['Date', 'Restaurant', 'Amount', 'Items', 'Filename']]
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range='A1:E1',
        valueInputOption='RAW',
        body={'values': headers}
    ).execute()
    
    return spreadsheet_id

def parse_invoices_to_sheet(folder_path='gmail_pdfs', sheet_name='Zomato Invoices'):
    """Parse all PDFs in folder and add to Google Sheet"""
    
    if not os.path.exists(folder_path):
        print(f"✗ Folder '{folder_path}' not found!")
        return
    
    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]
    
    if not pdf_files:
        print(f"✗ No PDF files found in '{folder_path}'")
        return
    
    print(f"Found {len(pdf_files)} PDF files in '{folder_path}'")
    
    # Authenticate
    print("\nAuthenticating with Google Sheets...")
    service = authenticate_sheets()
    
    # Create spreadsheet
    print(f"\nCreating Google Sheet: '{sheet_name}'")
    spreadsheet_id = create_spreadsheet(service, sheet_name=sheet_name)
    
    # Parse invoices
    print("\n" + "="*60)
    print("PARSING INVOICES")
    print("="*60)
    
    rows_to_add = []
    successful = 0
    failed = 0
    
    for idx, filename in enumerate(pdf_files, 1):
        pdf_path = os.path.join(folder_path, filename)
        print(f"\n[{idx}/{len(pdf_files)}] Parsing: {filename}")
        
        invoice_data = extract_invoice_data(pdf_path)
        
        if invoice_data:
            items_str = ", ".join(invoice_data['items']) if invoice_data['items'] else "N/A"
            
            row = [
                invoice_data['date'] or 'N/A',
                invoice_data['restaurant'] or 'N/A',
                invoice_data['amount'] or 'N/A',
                items_str,
                invoice_data['filename']
            ]
            rows_to_add.append(row)
            
            print(f"  ✓ Date: {invoice_data['date']}")
            print(f"  ✓ Restaurant: {invoice_data['restaurant']}")
            print(f"  ✓ Amount: {invoice_data['amount']}")
            print(f"  ✓ Items: {len(invoice_data['items'])} found")
            
            successful += 1
        else:
            failed += 1
    
    # Add rows to sheet
    if rows_to_add:
        print("\n" + "="*60)
        print("WRITING TO GOOGLE SHEET")
        print("="*60)
        
        service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id,
            range='A2',
            valueInputOption='RAW',
            body={'values': rows_to_add}
        ).execute()
        
        print(f"\n✓ Successfully parsed: {successful} invoices")
        if failed > 0:
            print(f"⚠ Failed to parse: {failed} invoices")
        print(f"\n✓ View your sheet: https://docs.google.com/spreadsheets/d/{spreadsheet_id}")
    else:
        print("\n⚠ No invoice data could be extracted from any PDF")

if __name__ == '__main__':
    parse_invoices_to_sheet(folder_path='gmail_pdfs', sheet_name='Zomato Invoices')