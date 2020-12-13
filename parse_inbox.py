import pickle
import os.path
import base64
import re
import email
import requests
import time

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# Constants

SCOPES = ['https://mail.google.com/', 'https://www.googleapis.com/auth/spreadsheets']
FILTERS = ['office@triumphpm.com']

PMCONTACTS = ['office@triumphpm.com', 'leasingagents@triumphpm.com', 'Office@triumphpm.com',
                'guestcards@appfolio.com', 'contact@triumphpm.com', '(702) 367-2323', '(702) 816-8090',
                '(702) 250-5958', '(888) 658-7368']

SHEET_ID = '1Yfl1stekowIhbVqM5ofSzb4We7Z69n7_eoELsEQlOY4'

# Global Regex Patterns
ADDRESS_PATTERN = re.compile("in (.*)|for (.*)|about (.*)")
PHONE_PATTERN = re.compile(r".*?(\(\d{3}\D{0,3}\d{3}\D{0,3}\d{4}).*?", re.S)
EMAIL_PATTERN = re.compile(r'mailto:(\S+@\S+)"')
CONTACT_URL_PATTERN = re.compile(r'(https://www.zillow.com/rental-manager/inquiry-contact.*)>')
EXTERNAL_URL_NAME_PATTERN = re.compile(r'Text-h1">(.*?)</div>')
EXTERNAL_URL_NUMBER_PATTERN = re.compile(r'stacked-md">(.*?)</div>')
NAME_PATTERN_SUBJECT = re.compile("New Lead: (.*) interested")
NAME_PATTERN_EMAIL = re.compile(r'target="_blank">([^>]+)</a><br>')

def authenticate():
    """ Authorizes user if credentials exist, otherwise
        creates new token file and prompts user for
        authorization
    Args:
        None
    Returns:
        Gmail API Credential Object
    """
    creds = None

    #Check if token.pickle file exists
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If credentials are invalid or not available
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        # Save token for future use
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    return creds

def get_messages(gmail_service, amount):
    """ Creates service object using credentials and
        sends request to Gmail servers to retrieve a 
        specified amount of messages
    Args:
        gmail_service   -> Gmail API service object needed to make request
        amount          -> number of messages to be retrieved
    Return:
        messages        -> list containing all messages retrieved
    """
    result = gmail_service.users().messages().list(maxResults=amount, userId='me').execute()
    messages = result.get('messages')

    return messages

def check_headers(sender, date, subject):
    """ Takes in message data and runs checks to determine
        if message is a rental lead
    Args:
        sender      -> email address of person who sent the email
        date        -> email timestamp
        subject     -> email subject string
    Returns:
        boolean     -> True if email is lead, False otherwise
    """
    if sender != 'johngriffith6@gmail.com':
        return False
    matches = ADDRESS_PATTERN.findall(subject)
    if not matches:
        return False
    
    return True

def parse_data(body, subject):
    """ This function takes in cleartext html data, and email subject
        and parses it to find appropriate info

        There are two kinds of common formats for a lead email.
        One kind has the name and phone number in cleartext, 
        while the other requires you to click a hyperlink to
        access that data
    Args:
        body        -> html data in cleartext
    Returns:
        info        -> tuple of format (name, number, address)
    """
    info = ['', '', '']
    # First try for phone number in plain text
    matches = PHONE_PATTERN.findall(body)
    # Remove instances of PM phone number
    for number in matches:
        if number not in PMCONTACTS:
            info[1] = number
            break
    # Next, try for plain text email address if no phone number
    if info[1] == '':
        matches = EMAIL_PATTERN.findall(body)
        for email in matches:
            if email not in PMCONTACTS:
                info[1] = email
                break
    # If none of the above, contact info is embedded in a hyperlink
    if info[0] == '':
        matches = CONTACT_URL_PATTERN.findall(body)
        if matches:
            contact_url = matches[0]
            r = requests.get(contact_url)
            matches = EXTERNAL_URL_NAME_PATTERN.findall(r.text)
            if matches:
                info[0] = matches[0]
            # Now grab phone number from external link if possible
            if info[1] == '':
                matches = EXTERNAL_URL_NUMBER_PATTERN.findall(r.text)
                if matches:
                    info[1] = matches[0]
    # Parse subject line for name if not grabbed by previous operation
    """ Different subject line formats:
        ...New Lead: X interested... (used by Triumph PM)
    """
    if info[0] == '':
        matches = NAME_PATTERN_SUBJECT.findall(subject)
        if len(matches) > 0:
            info[0] = matches[0]

    # Parse email body for contact name if not in subject line or external url
    if info[0] == '':
        matches = NAME_PATTERN_EMAIL.findall(body)
        if len(matches) > 0:
            info[0] = ''.join(matches[0].splitlines())
        else:
            info[0] = 'No Name Provided'
    
    # Parse subject line for address of inquiry
    matches = ADDRESS_PATTERN.findall(subject)
    for item in matches[0]:
        if item != '':
            info[2] = item

    return tuple(info)

def decode(text):
    """ Takes raw email data text and decodes it
    Args:
        text        -> raw email text
    Returns:
        message     -> decoded text
    """
    if len(text) > 0:
        message = base64.urlsafe_b64decode(text)
        message = str(message, 'utf-8')
        #message = email.message_from_string(message)
        return message
    return None

def read_message(payload):
    """ Extracts email body data from payload
    Args:
        payload     -> payload of email
    Returns:
        message     -> decoded message body
    """
    message = None
    # JSON message data can be found in several different places
    if "data" in payload['body']:
        message = payload['body']['data']
        message = decode(message)
    elif "data" in payload['parts'][0]['body']:
        message = payload['parts'][0]['body']['data']
        message = decode(message)
    elif "parts" in payload['parts'][0]:
        biggest = 0
        for part in payload['parts'][0]['parts']:
            if "body" in part:
                if part['body']['size'] > biggest:
                    biggest = part['body']['size']
                    message = part['body']['data']
        message = decode(message)
    else:       # should never get here, failsafe
        print('body has no data.')
    return message

def write_to_sheet(sheets_service, info):
    """ Takes tuple of contact info and writes to specified Google Sheet
    Args:
        sheets_service  -> Google Sheets service object
        info            -> Tuple containing contact info
    Returns:
        Boolean         -> True if enough info was found for contact
    """
    # Perform checks on info tuple
    if info[1] == '' or info[2] == '':
        return False
    if info[0] == '':
        info[0] = 'No Name Provided'
    
    # Iterate through rows to find an empty one
    row = 1
    while(1):
        cells = 'A' + str(row) + ':C' + str(row)
        # Try except to avoid error on rate limit
        try:
            result = sheets_service.spreadsheets().values().get(
                spreadsheetId=SHEET_ID, range=cells).execute()
        except:
            print('Rate limit hit. Will wait 5 seconds and try again')
            time.sleep(5)
        # If no 'values' key in response, write info tuple to that row
        if 'values' not in result:
            values = [[info[0], info[1], info[2]]]
            body = {'values':values}
            result = sheets_service.spreadsheets().values().update(
                spreadsheetId=SHEET_ID, range=cells,
                valueInputOption='RAW', body=body).execute()
            break
        row += 1
    print("\nSuccessfully added one lead to spreadsheet\n")
    return True

def filter_messages(messages, gmail_service, sheets_service):
    """ Takes list of message objects and performs appropriate operations
        on each
    Args:
        messages        -> list of message objects
        gmail_service   -> Gmail API service object
        sheets_service  -> Sheets API service object
    Returns:
        None
    """
    # Iterate through messages queue and parse content
    for msg in messages:
        # Get message using id 
        print('Email Id: ' + msg['id'])
        text = gmail_service.users().messages().get(userId='me', id=msg['id']).execute()
        # Get header from payload
        payload = text['payload']
        headers = payload['headers']
        # Extract header information
        for x in headers:
            if x['name'] == 'From':
                sender = x['value']
                sender = re.sub('^[^<]*<', '', sender)
                sender = re.sub('>', '', sender)
            if x['name'] == 'Date':
                date = x['value']
            if x['name'] == 'Subject':
                subject = x['value']
        # If headers are appropriate, process data
        if check_headers(sender, date, subject):
            print('\n')
            print('Sender: ' + sender)
            print('Date: ' + date)
            print('Subject: ' + subject + '\n')
            message_body = read_message(payload)
            # Call parse_data function to extract contact information, and address of inquiry
            info = parse_data(message_body, subject)
            # Write to Google Sheets and delete email if true
            if write_to_sheet(sheets_service, info):
                # gmail_service.users().messages().delete(userId='me', id=msg['id']).execute()
                pass



def main():
    number_to_fetch = input("How many recent emails should be processed: ")
    # Authenticate with user and get credentials
    creds = authenticate()
    # Connect to Gmail API and use service object to retrieve messages
    gmail_service = build('gmail', 'v1', credentials=creds)
    sheets_service = build('sheets', 'v4', credentials=creds)
    messages = get_messages(gmail_service, number_to_fetch)
    
    filter_messages(messages, gmail_service, sheets_service)
    print("Finished parsing " + str(number_to_fetch) + " emails. Exiting program successfully...")

if __name__ == '__main__':
    main()