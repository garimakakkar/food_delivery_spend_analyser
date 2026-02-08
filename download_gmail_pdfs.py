import os
import base64
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from email.mime.text import MIMEText
import re

# If modifying these scopes, delete the file token.json
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def authenticate_gmail():
    """Authenticate and return Gmail API service"""
    creds = None
    
    # token.json stores the user's access and refresh tokens
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    
    # If there are no (valid) credentials available, let the user log in
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    
    return build('gmail', 'v1', credentials=creds)

def sanitize_filename(filename):
    """Remove invalid characters from filename"""
    return re.sub(r'[<>:"/\\|?*]', '_', filename)

def download_pdf_attachments(service, query='has:attachment filename:pdf', 
                            download_folder='gmail_pdfs', max_results=100):
    """
    Download PDF attachments from Gmail
    
    Args:
        service: Gmail API service instance
        query: Gmail search query (default: all PDFs)
        download_folder: Folder to save PDFs
        max_results: Maximum number of emails to process
    """
    
    # Create download folder if it doesn't exist
    if not os.path.exists(download_folder):
        os.makedirs(download_folder)
    
    print(f"Searching for emails with query: {query}")
    
    try:
        # Search for messages
        results = service.users().messages().list(
            userId='me', 
            q=query,
            maxResults=max_results
        ).execute()
        
        messages = results.get('messages', [])
        
        if not messages:
            print('No messages found.')
            return
        
        print(f"Found {len(messages)} emails. Processing...")
        
        pdf_count = 0
        
        for idx, message in enumerate(messages, 1):
            msg_id = message['id']
            
            # Get the full message
            msg = service.users().messages().get(
                userId='me', 
                id=msg_id
            ).execute()
            
            # Get email subject for organizing
            headers = msg['payload'].get('headers', [])
            subject = next((h['value'] for h in headers if h['name'] == 'Subject'), 'No Subject')
            sender = next((h['value'] for h in headers if h['name'] == 'From'), 'Unknown')
            
            print(f"\n[{idx}/{len(messages)}] Processing: {subject[:50]}...")
            
            # Process parts recursively
            parts = msg['payload'].get('parts', [])
            if not parts:
                # Check if the whole payload is an attachment
                parts = [msg['payload']]
            
            pdf_count = process_parts(service, parts, msg_id, download_folder, pdf_count, subject)
        
        print(f"\n✓ Complete! Downloaded {pdf_count} PDF attachments to '{download_folder}/'")
        
    except Exception as e:
        print(f'An error occurred: {e}')

def process_parts(service, parts, msg_id, download_folder, pdf_count, subject):
    """Recursively process message parts to find PDF attachments"""
    
    for part in parts:
        # Check nested parts
        if 'parts' in part:
            pdf_count = process_parts(service, part['parts'], msg_id, 
                                     download_folder, pdf_count, subject)
        
        filename = part.get('filename', '')
        
        # Check if it's a PDF attachment
        if filename and filename.lower().endswith('.pdf'):
            attachment_id = part['body'].get('attachmentId')
            
            if attachment_id:
                # Download the attachment
                attachment = service.users().messages().attachments().get(
                    userId='me',
                    messageId=msg_id,
                    id=attachment_id
                ).execute()
                
                data = attachment['data']
                file_data = base64.urlsafe_b64decode(data)
                
                # Create unique filename
                safe_filename = sanitize_filename(filename)
                filepath = os.path.join(download_folder, f"{pdf_count+1:04d}_{safe_filename}")
                
                # Save the file
                with open(filepath, 'wb') as f:
                    f.write(file_data)
                
                pdf_count += 1
                print(f"  ✓ Downloaded: {filename} ({len(file_data)/1024:.1f} KB)")
    
    return pdf_count

if __name__ == '__main__':
    # Authenticate
    service = authenticate_gmail()
    
    # Example queries - customize as needed:
    
    # All PDFs
    download_pdf_attachments(service, query=('(((has:attachment OR has:drive) has:pdf zomato) has:attachment after:2025/8/6 before:2026/8/7'))
    
    # PDFs from specific sender
    # download_pdf_attachments(service, query='from:sender@example.com has:attachment filename:pdf')
    
    # PDFs after a specific date
    # download_pdf_attachments(service, query='has:attachment filename:pdf after:2024/01/01')
    
    # PDFs with specific subject keywords
    # download_pdf_attachments(service, query='subject:invoice has:attachment filename:pdf')