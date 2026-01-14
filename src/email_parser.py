"""
Email Parser - Extracts relevant fields from Gmail API response
"""
import base64
from email.utils import parsedate_to_datetime
from html import unescape
import re

# Google Sheets has a 50,000 character limit per cell
MAX_CONTENT_LENGTH = 45000  # Leave some buffer

class EmailParser:
    def parse_email(self, email_data):
        """Parse email data from Gmail API"""
        try:
            headers = email_data['payload']['headers']
            
            # Extract headers
            from_email = self._get_header(headers, 'From')
            subject = self._get_header(headers, 'Subject')
            date_str = self._get_header(headers, 'Date')
            
            # Parse date
            date = self._parse_date(date_str)
            
            # Extract body
            content = self._get_body(email_data['payload'])
            
            # Truncate content if too long
            content = self._truncate_content(content)
            
            return {
                'id': email_data['id'],
                'from': from_email,
                'subject': subject,
                'date': date,
                'content': content
            }
        
        except Exception as e:
            print(f"Error parsing email: {e}")
            return None
    
    def _get_header(self, headers, name):
        """Get specific header value"""
        for header in headers:
            if header['name'].lower() == name.lower():
                return header['value']
        return ''
    
    def _parse_date(self, date_str):
        """Parse email date to readable format"""
        try:
            dt = parsedate_to_datetime(date_str)
            return dt.strftime('%Y-%m-%d %H:%M:%S')
        except:
            return date_str
    
    def _get_body(self, payload):
        """Extract email body (plain text preferred)"""
        body = ''
        
        if 'parts' in payload:
            for part in payload['parts']:
                if part['mimeType'] == 'text/plain':
                    body = self._decode_body(part['body'])
                    break
                elif part['mimeType'] == 'text/html' and not body:
                    html_body = self._decode_body(part['body'])
                    body = self._html_to_text(html_body)
        else:
            if 'body' in payload and 'data' in payload['body']:
                body = self._decode_body(payload['body'])
        
        return body.strip()
    
    def _decode_body(self, body_data):
        """Decode base64 email body"""
        if 'data' in body_data:
            try:
                decoded = base64.urlsafe_b64decode(
                    body_data['data'].encode('ASCII')
                )
                return decoded.decode('utf-8')
            except:
                return ''
        return ''
    
    def _html_to_text(self, html):
        """Convert HTML to plain text"""
        # Remove HTML tags
        text = re.sub('<[^<]+?>', '', html)
        # Decode HTML entities
        text = unescape(text)
        # Clean up whitespace
        text = re.sub(r'\s+', ' ', text)
        return text.strip()
    
    def _truncate_content(self, content):
        """Truncate content to fit Google Sheets cell limit"""
        if len(content) <= MAX_CONTENT_LENGTH:
            return content
        
        # Truncate and add indicator
        truncated = content[:MAX_CONTENT_LENGTH]
        truncated += f"\n\n[... Content truncated. Original length: {len(content)} characters]"
        return truncated