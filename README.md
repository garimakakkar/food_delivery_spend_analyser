# Food Delivery Spend Analyser

Automatically download PDF invoices from Gmail and parse them into a Google Sheet.

## Features
- Downloads PDF attachments from Gmail using search queries
- Parses invoices to extract date, restaurant, amount, and items
- Creates a Google Sheet with organized invoice data

## Setup
1. Install dependencies: `pip install pdfplumber google-api-python-client google-auth-httplib2 google-auth-oauthlib`
2. Set up Google Cloud project and enable Gmail API + Google Sheets API
3. Download `credentials.json` from Google Cloud Console
4. Run the scripts

## Usage
```bash
# Download invoices
python3 download_gmail_pdfs.py

# Parse existing invoices
python3 parse_invoices.py
```

## Security Note
Never commit `credentials.json` or `token.json` to the repository.
