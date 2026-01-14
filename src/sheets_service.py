"""
Google Sheets Service - Handles Google Sheets API operations
"""
import os
import pickle
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
TOKEN_PATH = 'credentials/sheets_token.pickle'
CREDENTIALS_PATH = os.path.join(
    BASE_DIR,
    "credentials",
    "credentials.json"
)
class SheetsService:
    def __init__(self):
        self.service = self._authenticate()
    
    def _authenticate(self):
        """Authenticate and return Sheets service"""
        creds = None
        
        # Load existing token
        if os.path.exists(TOKEN_PATH):
            with open(TOKEN_PATH, 'rb') as token:
                creds = pickle.load(token)
        
        # Refresh or get new credentials
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    CREDENTIALS_PATH, SCOPES)
                creds = flow.run_local_server(port=0)
            
            # Save credentials
            os.makedirs('credentials', exist_ok=True)
            with open(TOKEN_PATH, 'wb') as token:
                pickle.dump(creds, token)
        
        return build('sheets', 'v4', credentials=creds)
    
    def _get_existing_emails(self, spreadsheet_id, sheet_name):
        """Get existing email IDs from state file"""
        state_file = 'email_state.txt'
        existing = set()
        
        if os.path.exists(state_file):
            with open(state_file, 'r') as f:
                existing = set(line.strip() for line in f)
        
        return existing
    
    def _save_email_ids(self, email_ids):
        """Save processed email IDs to state file"""
        state_file = 'email_state.txt'
        
        with open(state_file, 'a') as f:
            for email_id in email_ids:
                f.write(f"{email_id}\n")
    
    def append_emails(self, spreadsheet_id, sheet_name, emails):
        """Append emails to sheet, avoiding duplicates"""
        try:
            # Get existing email IDs
            existing_ids = self._get_existing_emails(spreadsheet_id, sheet_name)
            
            # Filter out duplicates
            new_emails = [e for e in emails if e['id'] not in existing_ids]
            
            if not new_emails:
                print("No new emails to add (all are duplicates)")
                return 0
            
            # Prepare rows with content length validation
            rows = []
            new_ids = []
            skipped = 0
            
            for email in new_emails:
                # Double-check content length (safety check)
                content = email['content']
                if len(content) > 50000:
                    content = content[:45000] + "\n\n[Content truncated - exceeded character limit]"
                    print(f"Warning: Truncated email from {email['from']}")
                
                # Also check subject length
                subject = email['subject']
                if len(subject) > 50000:
                    subject = subject[:1000] + "... [truncated]"
                
                rows.append([
                    email['from'][:50000],  # Safety check all fields
                    subject,
                    email['date'][:50000],
                    content
                ])
                new_ids.append(email['id'])
            
            # Append to sheet
            body = {'values': rows}
            self.service.spreadsheets().values().append(
                spreadsheetId=spreadsheet_id,
                range=f'{sheet_name}!A:D',
                valueInputOption='RAW',
                body=body
            ).execute()
            
            # Save new IDs to state
            self._save_email_ids(new_ids)
            
            if skipped > 0:
                print(f"Skipped {skipped} email(s) due to size issues")
            
            return len(new_emails)
        
        except HttpError as error:
            print(f'An error occurred: {error}')
            # If still getting character limit error, provide helpful message
            if '50000 characters' in str(error):
                print("\nTroubleshooting: Some email content is still too long.")
                print("This has been logged. Check email_sync.log for details.")
            return 0
    
    def initialize_sheet(self, spreadsheet_id, sheet_name):
        """Initialize sheet with headers if empty"""
        try:
            headers = [['From', 'Subject', 'Date', 'Content']]
            body = {'values': headers}
            
            self.service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=f'{sheet_name}!A1:D1',
                valueInputOption='RAW',
                body=body
            ).execute()
            
            print("Sheet initialized with headers")
        
        except HttpError as error:
            print(f'Error initializing sheet: {error}')