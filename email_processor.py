import imaplib
import email
import os
import requests
from email.header import decode_header
import json
import sys
import urllib.parse

def check_environment():
    """Check and validate all environment variables"""
    required_vars = ['EMAIL', 'EMAIL_PASSWORD', 'WEBHOOK_URL', 'IMAP_SERVER']
    missing_vars = [var for var in required_vars if not os.getenv(var)]

    if missing_vars:
        print(f"ERROR: Missing required environment variables: {', '.join(missing_vars)}")
        sys.exit(1)

def decode_str(string):
    """Decode email subject or sender name"""
    decoded_string, charset = decode_header(string)[0]
    if isinstance(decoded_string, bytes):
        return decoded_string.decode(charset or 'utf-8')
    return decoded_string

def get_email_body(msg):
    """Extract email body, preferring HTML content"""
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/html":
                return part.get_payload(decode=True).decode()
            elif part.get_content_type() == "text/plain":
                return part.get_payload(decode=True).decode()
    return msg.get_payload(decode=True).decode()

def process_name(full_name):
    """Split full name into first and last name"""
    parts = full_name.split()
    first_name = parts[0] if parts else ""
    last_name = " ".join(parts[1:]) if len(parts) > 1 else ""
    return first_name, last_name

def ensure_https_url(url):
    """Ensure the URL uses HTTPS"""
    parsed = urllib.parse.urlparse(url)
    if parsed.scheme != 'https':
        return url.replace('http://', 'https://', 1)
    return url

def send_webhook_request(url, data):
    """Send webhook request with proper headers and HTTPS"""
    url = ensure_https_url(url)

    headers = {
        'Content-Type': 'application/json',
        'Accept': 'application/json',
        'Origin': urllib.parse.urlparse(url).netloc,
        'Access-Control-Request-Method': 'POST',
        'Access-Control-Request-Headers': 'Content-Type, Accept',
        'User-Agent': 'FluentSupport-EmailProcessor/1.0'
    }

    # First, send preflight request
    preflight_response = requests.options(
        url,
        headers=headers,
        timeout=30,
        verify=True  # Enforce SSL verification
    )

    print(f"Preflight response status: {preflight_response.status_code}")
    print(f"Preflight headers: {dict(preflight_response.headers)}")

    # Send actual request
    response = requests.post(
        url,
        json=data,
        headers=headers,
        timeout=30,
        verify=True  # Enforce SSL verification
    )

    return response

def main():
    try:
        check_environment()

        EMAIL = os.environ['EMAIL']
        PASSWORD = os.environ['EMAIL_PASSWORD']
        WEBHOOK_URL = os.environ['WEBHOOK_URL']
        IMAP_SERVER = os.environ['IMAP_SERVER']

        print("Connecting to email server...")
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)

        try:
            mail.login(EMAIL, PASSWORD)
            print("Successfully logged in")
        except imaplib.IMAP4.error as e:
            print(f"Failed to login: {str(e)}")
            sys.exit(1)

        mail.select('INBOX')

        # Search for unread emails
        _, messages = mail.search(None, 'UNSEEN')

        if not messages[0]:
            print("No unread messages found. Exiting.")
            mail.logout()
            return

        message_nums = messages[0].split()
        print(f"Found {len(message_nums)} unread messages")

        for msg_num in message_nums:
            try:
                print(f"\nProcessing message {msg_num.decode()}...")
                _, msg_data = mail.fetch(msg_num, '(RFC822)')
                email_body = msg_data[0][1]
                msg = email.message_from_bytes(email_body)

                # Extract sender information
                from_header = msg['from']
                if '<' in from_header:
                    sender_name = from_header.split('<')[0].strip().strip('"')
                    sender_email = from_header.split('<')[1].strip('>')
                else:
                    sender_name = ''
                    sender_email = from_header.strip()

                first_name, last_name = process_name(sender_name)

                # Prepare ticket data
                ticket_data = {
                    'title': decode_str(msg['subject']),
                    'content': get_email_body(msg),
                    'priority': 'Normal',
                    'sender': {
                        'first_name': first_name,
                        'last_name': last_name,
                        'email': sender_email
                    }
                }

                # Send request with CORS handling
                response = send_webhook_request(WEBHOOK_URL, ticket_data)

                if response.status_code == 200:
                    print(f"Created ticket: {ticket_data['title']}")
                    mail.store(msg_num, '+FLAGS', '\\Seen')
                    print("Marked email as read")
                else:
                    print(f"Failed to create ticket. Status: {response.status_code}")
                    print(f"Response content: {response.text}")

            except Exception as e:
                print(f"Error processing message: {str(e)}")
                continue

        mail.logout()
        print("Finished processing messages")

    except Exception as e:
        print(f"Unexpected error: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()