This tool serves to extract contact information from specific contacts using Gmail API, and appends data appropriately to a document using Google Sheets API

To use, program will prompt user to input the amount of emails they want to parse, and get to work. Program has built in list of senders to parse data from, but this could easily be changed hard coding the 'FILTERS' list in leads.py

Files:
    
    README.md   ->  README file
    
    leads.py    ->  Main file, takes in user input on number of messages to parse and proceeds appropriately

Note: Must already have a Google Sheets document and ID hard coded into the program (variable called "SHEETS_ID")

Depedencies:

    - pip install requests
    
    - pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib

To Use:

    python leads.py
