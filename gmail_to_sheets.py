import os.path
import base64
import re

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build



SCOPES = ['https://www.googleapis.com/auth/gmail.readonly',
          'https://www.googleapis.com/auth/spreadsheets']

# Google Sheet ID and target range
SPREADSHEET_ID = '1VL8YicNwk9ti_84ieRFk9GanoZVhaf_BbUOxQJdas38'
RANGE_NAME = 'Sheet1!A2'  

def gmail_authenticate():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)

        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds

def get_emails(service):
    results = service.users().messages().list(userId='me', maxResults=5).execute()
    messages = results.get('messages', [])

    email_data = []
    for msg in messages:
        msg_data = service.users().messages().get(userId='me', id=msg['id']).execute()
        headers = msg_data['payload']['headers']

        subject = next((h['value'] for h in headers if h['name'] == 'Subject'), 'No Subject')
        snippet = msg_data.get('snippet', '')

        email_data.append([subject, snippet])
    
    return email_data

def update_sheet(data):
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()
    request = sheet.values().append(
        spreadsheetId=SPREADSHEET_ID,
        range=RANGE_NAME,
        valueInputOption="RAW",
        body={"values": data}
    )
    request.execute()

def main():
    creds = gmail_authenticate()
    gmail_service = build('gmail', 'v1', credentials=creds)

    print(" Fetching emails...")
    emails = get_emails(gmail_service)

    if emails:
        print(f"📋 Writing {len(emails)} emails to Google Sheet...")
        update_sheet(emails)
        print("✅ Done!")
    else:
        print("❌ No emails found.")

if __name__ == '__main__':
    main()
