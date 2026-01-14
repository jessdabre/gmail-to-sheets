import os
import pickle
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

SCOPES = ['https://www.googleapis.com/auth/gmail.modify']
TOKEN_PATH = 'credentials/token.pickle'
CREDENTIALS_PATH = os.path.join(
    BASE_DIR,
    "credentials",
    "credentials.json"
)

class GmailService:
    def __init__(self):
        self.service = self._authenticate()
    
    def _authenticate(self):
        """Authenticate and return Gmail service"""
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
        
        return build('gmail', 'v1', credentials=creds)
    
    def get_unread_emails(self):
        """Fetch all unread emails from inbox"""
        try:
            results = self.service.users().messages().list(
                userId='me',
                q='is:unread in:inbox',
                maxResults=100
            ).execute()
            
            messages = results.get('messages', [])
            
            emails = []
            for message in messages:
                email_data = self.service.users().messages().get(
                    userId='me',
                    id=message['id'],
                    format='full'
                ).execute()
                emails.append(email_data)
            
            return emails
        
        except HttpError as error:
            print(f'An error occurred: {error}')
            return []
    
    def mark_as_read(self, message_ids):
        """Mark emails as read"""
        try:
            for msg_id in message_ids:
                self.service.users().messages().modify(
                    userId='me',
                    id=msg_id,
                    body={'removeLabelIds': ['UNREAD']}
                ).execute()
        except HttpError as error:
            print(f'Error marking as read: {error}')