import pickle
import os
import base64
import re
import email
import requests
import time
import datetime
import json

import wave
# import speech_recognition as sr

from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# Constants
SCOPES = ['https://mail.google.com/', 'https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
FILTERS = ['office@triumphpm.com', 'tpmassistant@triumphpm.com', 'wasim@faraneshlv.com', 'lorraine@iresvegas.com',
            'leasing@total-re.com', 'juli@market1realty.com', 'yourrealtorbrian@gmail.com', 'jackie.akester@gmail.com',
            'peacerealtylv@gmail.com', 'rentvegasnow@gmail.com', 'JDutt@mcgareycampagroup.com', 'morrisamortensen@gmail.com',
            'leasing@faraneshlv.com', 'hello@faraneshlv.com', 'lorraine@peacerealtylv.com', 'stacy.bandgrealty@gmail.com']

PMCONTACTS = ['office@triumphpm.com', 'leasingagents@triumphpm.com', 'Office@triumphpm.com',
                'guestcards@appfolio.com', 'contact@triumphpm.com', '(702) 367-2323', '(702) 816-8090',
                '(702) 250-5958', '(888) 658-7368', 'Lorraine@iresvegas.com', 'lorraine@iresvegas.com',
                '702-902-4304', 'RentalFeedConnect@zillowgroup.com', 'rentalapplications@zillow.com',
                '(702) 265-5604', '(702)%2047823', '(702)%2050146', '702-478-2360', '702-501-4649',
                '702-550-2222', '702-367-2323', '725-222-3015', 'VirtualOfficeVoiceMails@8x8.com']

SHEET_ID = '1Yfl1stekowIhbVqM5ofSzb4We7Z69n7_eoELsEQlOY4'
DRIVE_ID = '1ZrK2q5hz0zbbNpbhsHELmVOlisaVACx4'

# Global Regex Patterns
CONTACT_PATTERN_1 = re.compile(r'%26phone=(.*)%26date')
CONTACT_PATTERN_2 = re.compile(r".*?(\(\d{3}\D{0,3}\d{3}\D{0,3}\d{4}).*?", re.S)
CONTACT_PATTERN_3 = re.compile(r'phone=(.*?)&date=', re.DOTALL)
CONTACT_PATTERN_4 = re.compile(r'\r\n(\d{3}-\d{3}-\d{4})\r\n')
CONTACT_PATTERN_5 = re.compile(r'&amp;phone=(.*?)&amp')
CONTACT_PATTERN_6 = re.compile(r'\r\n> PHONE\r\n> (.*)\r\n')
EMAIL_PATTERN_1 = re.compile(r'\r\n(.*)<mailto:.*?>')
EMAIL_PATTERN_2 = re.compile(r'<a href="mailto:(.*?)">')
EMAIL_PATTERN_3 = re.compile(r'\r\n(.*?@.*?)\r\nCOMMENTS')
EMAIL_PATTERN_4 = re.compile(r'\r\n> (.*?@.*?)\r\n> COMMENTS')
EMAIL_PATTERN_5 = re.compile(r'3D>\r\n(.*)\r\n\r\nCOMMENTS')
EMAIL_PATTERN_6 = re.compile(r'3D>\r\n(.*)\r\n')
ADDRESS_PATTERN_1 = re.compile(r' about (.*)')
ADDRESS_PATTERN_2 = re.compile(r' for (.*)')
ADDRESS_PATTERN_3 = re.compile(r' in (.*)')
ADDRESS_PATTERN_4 = re.compile(r' Lead for (.*) from ')
ADDRESS_PATTERN_5 = re.compile(r' to tour (.*)')
ADDRESS_PATTERN_6 = re.compile(r' \((.*)\)')
SUBJECT_NAME_1 = re.compile(r'FW: (.*) is requesting')
SUBJECT_NAME_2 = re.compile(r'Fwd: (.*) is requesting')
SUBJECT_NAME_3 = re.compile(r'New Lead: (.*) interested')
SUBJECT_NAME_4 = re.compile(r'Re: (.*) is requesting')
SUBJECT_NAME_5 = re.compile(r'Lead from (.*) \(')
SUBJECT_NAME_6 = re.compile(r'from (.*) - ')
SUBJECT_NAME_7 = re.compile(r'AG Lead from (.*) \(')
BODY_NAME_1 = re.compile(r'\r\n\r\n\r\n(.*)\r\n\(?\d\d\d\)?')
BODY_NAME_2 = re.compile(r'CONTACT INFO\r\n\r\n(.*?)<https://link.edgepilot')
BODY_NAME_3 = re.compile(r'</b> (.*) &lt;guestcards')
BODY_NAME_4 = re.compile(r'Lead (.*) found you through RentPath')
BODY_NAME_5 = re.compile(r'&amp;name=(.*?)&amp')
BODY_NAME_6 = re.compile(r'From: (.*?) <guestcards@appfolio.com>')

# Global Variables
SHEET_ROW = 1
NEXT_PAGE_TOKEN = ''
MESSAGE_COUNT = 0
DATE = datetime.datetime.now().strftime('%x')

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
    global NEXT_PAGE_TOKEN
    if NEXT_PAGE_TOKEN == '':
        result = gmail_service.users().messages().list(maxResults=amount, userId='me', includeSpamTrash=False).execute()
    else:
        result = gmail_service.users().messages().list(maxResults=amount, userId='me', includeSpamTrash=False, pageToken=NEXT_PAGE_TOKEN).execute()
    messages = result.get('messages')
    NEXT_PAGE_TOKEN = result.get('nextPageToken')

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
    if sender in FILTERS:
        address = get_address('', subject)
        if address != '':
            return True
    
    return False

def get_name(body, subject):
    """ Wrapper function for name parsing"""
    # Try subject line first
    matches = SUBJECT_NAME_1.findall(subject)
    if matches:
        return matches[0]

    matches = SUBJECT_NAME_2.findall(subject)
    if matches:
        return matches[0]

    matches = SUBJECT_NAME_3.findall(subject)
    if matches:
        return matches[0]

    matches = SUBJECT_NAME_4.findall(subject)
    if matches:
        return matches[0]

    matches = SUBJECT_NAME_5.findall(subject)
    if matches:
        return matches[0]
    
    matches = SUBJECT_NAME_6.findall(subject)
    if matches:
        return matches[0]

    matches = SUBJECT_NAME_7.findall(subject)
    if matches:
        return matches[0]

    # Then try body
    matches = BODY_NAME_1.findall(body)
    if matches:
        return matches[0]

    matches = BODY_NAME_2.findall(body)
    if matches:
        return matches[0]

    matches = BODY_NAME_3.findall(body)
    if matches:
        return matches[0]

    matches = BODY_NAME_4.findall(body)
    if matches:
        return matches[0]

    matches = BODY_NAME_5.findall(body)
    if matches and '%20' in matches[0]:
        matches[0] = re.sub('%20', ' ', matches[0])
        return matches[0]

    matches = BODY_NAME_6.findall(body)
    if matches:
        return matches[0]

    print('Could not get name')
    return ''

def get_contact(body):
    """ Wrapper function for contact parsing"""
    # Try getting phone number
    matches = CONTACT_PATTERN_1.findall(body)
    if matches:
        for match in matches:
            if match not in PMCONTACTS:
                return match
    
    matches = CONTACT_PATTERN_2.findall(body)
    if matches:
        for match in matches:
            if match not in PMCONTACTS:
                return match

    matches = CONTACT_PATTERN_3.findall(body)
    if matches:
        for match in matches:
            if match not in PMCONTACTS:
                return match

    matches = CONTACT_PATTERN_4.findall(body)
    if matches:
        for match in matches:
            if match not in PMCONTACTS:
                return match

    matches = CONTACT_PATTERN_5.findall(body)
    if matches:
        for match in matches:
            if match not in PMCONTACTS:
                return match

    matches = CONTACT_PATTERN_6.findall(body)
    if matches:
        for match in matches:
            if match not in PMCONTACTS:
                return match

    # Try getting email contact
    
    matches = EMAIL_PATTERN_1.findall(body)
    if matches:
        for match in matches:
            if match not in PMCONTACTS:
                return match

    matches = EMAIL_PATTERN_2.findall(body)
    if matches:
        for match in matches:
            if match not in PMCONTACTS:
                return match

    matches = EMAIL_PATTERN_3.findall(body)
    if matches:
        for match in matches:
            if match not in PMCONTACTS:
                return match

    matches = EMAIL_PATTERN_4.findall(body)
    if matches:
        for match in matches:
            if match not in PMCONTACTS:
                return match

    matches = EMAIL_PATTERN_5.findall(body)
    if matches:
        for match in matches:
            if match not in PMCONTACTS:
                return match

    matches = EMAIL_PATTERN_6.findall(body)
    if matches:
        for match in matches:
            if match not in PMCONTACTS and '@' in match:
                return match

    print("Could not get phone number")
    return ''

def get_address(body, subject):
    """ Wrapper function for address parsing"""
    matches = ADDRESS_PATTERN_1.findall(subject)
    if matches:
        return matches[0]
    
    matches = ADDRESS_PATTERN_2.findall(subject)
    if matches:
        return matches[0]

    matches = ADDRESS_PATTERN_3.findall(subject)
    if matches:
        return matches[0]
    
    matches = ADDRESS_PATTERN_4.findall(subject)
    if matches:
        return matches[0]

    matches = ADDRESS_PATTERN_5.findall(subject)
    if matches:
        return matches[0]

    matches = ADDRESS_PATTERN_6.findall(subject)
    if matches:
        return matches[0]

    return ''

def parse_data(body, subject):
    """ This function takes in cleartext html data, and email subject
        and parses it to find appropriate info
    Args:
        body        -> html data in cleartext
        subject     -> subject line
    Returns:
        info        -> list of format [name, number, address]
    """
    info = ['', '', '']

    info[0] = get_name(body, subject)
    info[1] = get_contact(body)
    info[2] = get_address(body, subject)

    return info

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

def post_processing(info):
    """ This function takes info list and performs formatting changes
    Args:
        info            -> list of lead contact information
    Returns:
        info            -> formatted list of lead contact information
    """
    if ' - gmail' in info[0]:
        info[0] = info[0].replace(' - gmail', '')
    if ' - icloud' in info[0]:
        info[0] = info[0].replace(' - icloud', '')
    if ' - yahoo' in info[0]:
        info[0] = info[0].replace(' - yahoo', '')
    if 'Automatic reply:' in info[0]:
        info[0] = info[0].replace('Automatic reply: ', '')
    if '@' in info[0]:
        info[0] = re.sub(r'(@.*)', '', info[0])

    if info[0] in PMCONTACTS:
        info[0] = 'No name provided'

    if ':' in info[2]:
        info[2] = re.sub(r':.*', '', info[2])
    if '*' in info[2]:
        info[2] = re.sub(r'\*.*', '', info[2])
    if 'Aparment List' in info[2]:
        info[2] = re.sub(r'from .* - Apartment List', '', info[2])

    if info[3]:
        info[3] = re.sub(r' [\+-]\d\d\d\d', '', info[3])
        info[3] = re.sub(r'\w\w\w, ', '', info[3])

    return info

def write_to_sheet(sheets_service, info):
    """ Takes tuple of contact info and writes to specified Google Sheet
    Args:
        sheets_service  -> Google Sheets service object
        info            -> Tuple containing contact info
    Returns:
        Boolean         -> True if enough info was found for contact
    """
    global SHEET_ROW
    info = list(info)
    # Perform checks on info tuple
    if info[1] == '' or info[2] == '':
        print('Lead is missing either a contact, or address\n')
        return False
    if len(info[1]) > 50:
        print('Something went wrong during parsing, contact info not valid\n')
        return False
    if info[0] == '':
        info[0] = 'No Name Provided'
    # Post Processing for info strings
    post_processing(info)
    # Find first empty row (quickly)
    result = sheets_service.spreadsheets().values().batchGet(
        spreadsheetId=SHEET_ID, ranges='A1:A2500').execute()
    SHEET_ROW = len(result['valueRanges'][0]['values'])
    # Iterate through sheet to fill rows
    while(1):
        SHEET_ROW += 1
        cells = 'A' + str(SHEET_ROW) + ':D' + str(SHEET_ROW)
        # Try except to avoid error on rate limit
        try:
            result = sheets_service.spreadsheets().values().get(
                spreadsheetId=SHEET_ID, range=cells).execute()
        except:
            print('Rate limit hit. Will wait 20 seconds and try again')
            time.sleep(20)
            continue
        # If no 'values' key in response, write info tuple to that row
        if 'values' not in result:
            values = [[info[0], info[1], info[2], info[3]]]
            body = {'values':values}
            result = sheets_service.spreadsheets().values().update(
                spreadsheetId=SHEET_ID, range=cells,
                valueInputOption='RAW', body=body).execute()
            break
    return True

def filter_messages(messages, gmail_service, sheets_service, drive_service):
    """ Takes list of message objects and performs appropriate operations
        on each
    Args:
        messages        -> list of message objects
        gmail_service   -> Gmail API service object
        sheets_service  -> Sheets API service object
        drive_service   -> Drive API service object
    Returns:
        None
    """
    global MESSAGE_COUNT
    # Iterate through messages queue and parse content
    for msg in messages:
        # Get message using id 
        # print('Processing email #' + str(MESSAGE_COUNT))
        MESSAGE_COUNT += 1
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
            print('------Found Lead in Message #' + str(MESSAGE_COUNT) + '------')
            print('Sender: ' + sender)
            print('Date: ' + date)
            print('Subject: ' + subject)
            # Call parse_data function to extract contact information, and address of inquiry
            # Try-except to handle errors that may terminate program
            try:
                message_body = read_message(payload)
                info = parse_data(message_body, subject)
                if date:
                    info.append(date)
                # Attempts to write to Google Sheets and puts email in trash if successful
                print(info)
                if write_to_sheet(sheets_service, info):
                    gmail_service.users().messages().trash(userId='me', id=msg['id']).execute()
                    print("Successfully added contact to spreadsheet\n")
            except:
                print('Error processing message #' + str(MESSAGE_COUNT) + '\n')
        # If message is a voicemail
        if sender == 'VirtualOfficeVoiceMails@8x8.com':
            print('------Voicemail Lead Found in Message #' + str(MESSAGE_COUNT) + '------')
            # Get ID of attachment
            for part in payload['parts']:
                if part['mimeType'] == 'audio/x-wav':
                    att_id = part['body']['attachmentId']
            # Get audio binary and save to file
            audio_binary = gmail_service.users().messages().attachments().get(userId='me', messageId=msg['id'], id=att_id).execute()
            audio_data = base64.urlsafe_b64decode(audio_binary['data'].encode('UTF-8'))
            w = wave.open('temp_audio.wav', 'w')
            w.setnchannels(1)
            w.setsampwidth(2)
            w.setframerate(8000)
            w.writeframesraw(audio_data)
            w.close()
            # Upload file to shared Google Drive folder
            metadata = {'name':DATE + '.wav', 'parents':[DRIVE_ID]}
            media = MediaFileUpload('temp_audio.wav', mimetype='audio/wav')
            try:
                drive_service.files().create(body=metadata, media_body=media, fields='id').execute()
                gmail_service.users().messages().trash(userId='me', id=msg['id']).execute()
                print('Added voicemail Lead to Google Drive\n')
            except:
                print('There was a problem adding the voicemail to Google Drive\n')
            """
            # Use voice recognition to get voicemail string (POOR ACCURACY FOR VOICEMAILS)
            r = sr.Recognizer()
            r.pause_threshold = 10
            r.non_speaking_duration = 10
            with sr.AudioFile('temp_audio.wav') as source:
                audio = r.record(source)
                try:
                    transcript = r.recognize_google(audio)
                    print(transcript)
                except:
                    print('Could not transcribe audio')
                print('')
            """

def main():
    number_to_fetch = input("How many recent emails should be processed: ")
    # Authenticate with user and get credentials
    try:
        creds = authenticate()
    except:
        os.remove('token.pickle')
        print("Credentials have expired, please run the program again")
    # Connect to Gmail API and use service object to retrieve messages
    gmail_service = build('gmail', 'v1', credentials=creds)
    sheets_service = build('sheets', 'v4', credentials=creds)
    drive_service = build('drive', 'v3', credentials=creds)
    # If less than 500 messages requested
    remaining = int(number_to_fetch)
    if remaining <= 500 and remaining > 0:
        messages = get_messages(gmail_service, str(remaining))
        filter_messages(messages, gmail_service, sheets_service, drive_service)
    # Gmail API messages GET call returns maximum of 500 messages
    # In this case we need multiple iterations
    elif remaining > 500:
        for _ in range(int(remaining/500)):
            messages = get_messages(gmail_service, str(remaining))
            filter_messages(messages, gmail_service, sheets_service, drive_service)
            remaining -= 500
    else:
        print('Invalid input. Exiting program...')
        exit
    print("Finished parsing " + str(number_to_fetch) + " emails. Exiting program...")

if __name__ == '__main__':
    main()