Gmail to Google Sheets Automation

High-Level Architecture


┌─────────────────┐
│   Gmail Inbox   │
│  (Unread only)  │
└────────┬────────┘
         │
         │ Gmail API (OAuth 2.0)
         ▼
┌─────────────────┐
│  gmail_service  │◄─── Fetches unread emails
│      .py        │     Marks as read
└────────┬────────┘
         │
         │ Raw email data
         ▼
┌─────────────────┐
│  email_parser   │◄─── Extracts: From, Subject,
│      .py        │     Date, Content (plain text)
└────────┬────────┘
         │
         │ Parsed email objects
         ▼
┌─────────────────┐
│ sheets_service  │◄─── Checks duplicates (state file)
│      .py        │     Appends new rows
└────────┬────────┘
         │
         │ Google Sheets API (OAuth 2.0)
         ▼
┌─────────────────┐
│  Google Sheet   │
│  [From|Subject  │
│   |Date|Content]│
└─────────────────┘

State Persistence:
┌─────────────────┐
│email_state.txt  │◄─── Stores processed email IDs
└─────────────────┘     (one ID per line)

Setup Instructions
1. Prerequisites
- Python 3.8 or higher
- A Google account
- Gmail and Google Sheets access

2. Clone the Repository
git clone <repo-url>
cd gmail-to-sheets

3. Install Dependencies
pip install -r requirements.txt

4. Enable Google APIs

Gmail API:
1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select existing one
3. Enable **Gmail API**
4. Go to "Credentials" → "Create Credentials" → "OAuth 2.0 Client ID"
5. Choose "Desktop app"
6. Download the credentials JSON file
7. Rename it to `credentials.json` and place in `credentials/` folder

Google Sheets API:
1. In the same project, enable **Google Sheets API**
2. Use the same `credentials.json` file

5. Create Google Sheet
1. Create a new Google Sheet
2. Copy the Spreadsheet ID from the URL:
   ```
   https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit
   ```
3. Open `config.py` and update `SPREADSHEET_ID`

6. Project Structure Setup
mkdir -p src credentials proof

7. Run the Script
python main.py


First run:
- Browser will open for OAuth consent
- Grant access to Gmail and Sheets
- Token will be saved for future runs


OAuth Flow Explanation

This project uses OAuth 2.0 with User Consent.

Flow:
1. First Run: Script checks for `token.pickle` and `sheets_token.pickle`
2. If not found, initiates OAuth flow via browser
3. User logs in and grants permissions
4. Access token + refresh token are saved locally
5. Subsequent Runs: Tokens are reused and auto-refreshed if expired

Why two tokens?
- `token.pickle`: Gmail API access
- `sheets_token.pickle`: Google Sheets API access
- Separate tokens for security and scope isolation

Duplicate Prevention Logic

Method:ile-based state persistence (`email_state.txt`)

How it works:
1. Each Gmail message has a unique ID
2. After processing an email, its ID is appended to `email_state.txt`
3. Before adding to sheet, script checks if email ID exists in state file
4. Only new emails (not in state file) are appended

Why this approach?
- ✅ Simple and reliable
- ✅ Persists across script runs
- ✅ No API calls needed to check duplicates
- ✅ Fast O(n) lookup using sets
- ⚠️ State file grows over time (can be archived periodically)

Alternative approaches considered:
- Querying Google Sheets for existing emails (slower, uses API quota)
- Using last processed timestamp (misses emails if received simultaneously)

## 📊 State Persistence Method

File:`email_state.txt`

Format:
```
1a2b3c4d5e6f7g8h
2b3c4d5e6f7g8h9i
3c4d5e6f7g8h9i0j
...
```

Implementation:
1. On startup, read all IDs into a Python set
2. Filter incoming emails against this set
3. Append new IDs after successful sheet update
4. Set-based lookup ensures O(1) duplicate checking

hallenge Faced & Solution

Challenge 1: HTML-only emails breaking the parser

Problem: Some emails only have HTML content (no plain text part). Initial implementation failed to extract body text, leaving the Content column empty.

Solution:
1. Modified `email_parser.py` to check for both `text/plain` and `text/html` MIME types
2. Implemented `_html_to_text()` method using regex to strip HTML tags
3. Falls back to HTML conversion if plain text unavailable
4. Added error handling for malformed HTML

Challenge 2: Google Sheets 50,000 character limit error

Problem: Some emails (especially newsletters or automated reports) have very long content that exceeds Google Sheets' 50,000 character per cell limit, causing the API to reject the entire batch.

Solution:
1. Added `MAX_CONTENT_LENGTH = 45000` constant (with buffer)
2. Implemented `_truncate_content()` method in email parser
3. Automatically truncates long content with clear indicator
4. Added safety checks in sheets service for all fields
5. Created `debug_emails.py` utility to identify large emails before processing

Code snippet:

def _truncate_content(self, content):
    if len(content) <= MAX_CONTENT_LENGTH:
        return content
    truncated = content[:MAX_CONTENT_LENGTH]
    truncated += f"\n\n[... Content truncated. Original length: {len(content)} characters]"
    return truncated


Limitations

1. Scalability: File-based state doesn't scale for millions of emails
   - Solution: Migrate to SQLite or database
   
2. API Quota: Gmail API has daily quotas (1 billion quota units/day)
   - Solution: Implement exponential backoff and batch processing

3. Concurrency: Not safe for parallel execution
   - Solution: Add file locking or use database with transactions

4. Email Size: Very large emails (>45KB content) are truncated
   - Solution: Content is automatically truncated with clear indicator
   - Use `debug_emails.py` to identify large emails before processing

5. State Recovery: If `email_state.txt` is deleted, all emails reprocess
   - Solution: Backup state file or use sheet-based verification

6. HTML Conversion: Complex HTML formatting is lost
   - Solution: Store HTML separately or use better parsing library (BeautifulSoup)
