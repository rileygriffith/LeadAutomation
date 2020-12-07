import pickle
import os.path
import base64
import re

from bs4 import BeautifulSoup

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

SCOPES = ['https://mail.google.com/']
FILTERS = ['office@triumphpm.com']
PHONE_REGEX = re.compile('^(\+\d{1,2}\s)?\(?\d{3}\)?[\s.-]\d{3}[\s.-]\d{4}$')

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

def get_messages(service, amount):
    """ Creates service object using credentials and
        sends request to Gmail servers to retrieve a 
        specified amount of messages
    Args:
        service     -> service object needed to make request
        amount      -> number of messages to be retrieved
    Return:
        messages    -> list containing all messages retrieved
    """
    result = service.users().messages().list(maxResults=amount, userId='me').execute()
    messages = result.get('messages')

    return messages

def main():
    # Authenticate with user and get credentials
    creds = authenticate()
    # Connect to Gmail API and use service object to retrieve messages
    service = build('gmail', 'v1', credentials=creds)
    messages = get_messages(service, 200)
    
    # Iterate through messages queue and parse content
    for msg in messages:
        # Get message using id 
        text = service.users().messages().get(userId='me', id=msg['id']).execute()
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
                
        # If previously extracted values are appropriate, decode body of email (base64)
        if sender == 'johngriffith6@gmail.com':
            print('Sender: ' + sender)
            print('Date: ' + date)
            print('Subject: ' + subject)
            data = payload['parts'][0]['body']['data']
            data = data.replace("-","+").replace("_","/")
            decoded_data = base64.b64decode(data)
            body = str(decoded_data)
            #TODO: Match phone numbers and append to output file
            """contact = PHONE_REGEX.findall(body)
            print(contact)
            f = open('output.txt', 'a')
            f.write(contact)
            f.close()
            """

if __name__ == '__main__':
    main()