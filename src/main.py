import os
import sys
from gmail_service import GmailService
from sheets_service import SheetsService
from email_parser import EmailParser

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))


from config import SPREADSHEET_ID, SHEET_NAME
import logging

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('email_sync.log'),
        logging.StreamHandler()
    ]
)

def main():
    """Main function to sync Gmail to Google Sheets"""
    logger = logging.getLogger(__name__)
    
    try:
        # Initialize services
        logger.info("Initializing Gmail service...")
        gmail_service = GmailService()
        
        logger.info("Initializing Sheets service...")
        sheets_service = SheetsService()
        
        # Get unread emails
        logger.info("Fetching unread emails...")
        emails = gmail_service.get_unread_emails()
        
        if not emails:
            logger.info("No unread emails found.")
            return
        
        logger.info(f"Found {len(emails)} unread email(s)")
        
        # Parse emails
        parser = EmailParser()
        parsed_emails = []
        
        for email_data in emails:
            parsed = parser.parse_email(email_data)
            if parsed:
                # Log email details for debugging
                content_len = len(parsed['content'])
                logger.debug(f"Parsed email from {parsed['from']}: {content_len} chars")
                
                if content_len > 40000:
                    logger.warning(f"Large email detected: {content_len} chars from {parsed['from']}")
                
                parsed_emails.append(parsed)
            else:
                logger.warning("Failed to parse an email")
        
        # Check for duplicates and append to sheet
        if parsed_emails:
            logger.info(f"Processing {len(parsed_emails)} email(s)...")
            new_count = sheets_service.append_emails(
                SPREADSHEET_ID, 
                SHEET_NAME, 
                parsed_emails
            )
            logger.info(f"Added {new_count} new email(s) to sheet")
            
            # Mark emails as read
            email_ids = [email['id'] for email in emails]
            gmail_service.mark_as_read(email_ids)
            logger.info(f"Marked {len(email_ids)} email(s) as read")
        
        logger.info("Sync completed successfully!")
        
    except Exception as e:
        logger.error(f"Error during sync: {str(e)}", exc_info=True)
        raise

if __name__ == "__main__":
    main()