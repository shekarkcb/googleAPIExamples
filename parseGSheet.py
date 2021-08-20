#!/usr/bin/python3
"""
Script : parseGSheet.py
Desc   : Sample script to parse the gsheet contents.
         1. Create a project in google api console.
         2. Allow spreadsheets API permissions to that project
         3. Create oauth tokens for that project
         4. download the token informaton as token.json and keep in the same directory as this file
         7. On the first run, it will try to authorize using the token.json plus google authentication page.
         8. Enter the same creds as spreadsheet owner/permissioned user to authenticate/authorize and click allow.
         9. Flow is complete, then it downloads token.pickle, both toekn.json & token.pickle is required to parse gsheets
         without hitting authorization page again.
         10. SPREADSHEET_ID is your google spreadsheet id, which is part of URL
Modified : @shekarkcb         
"""
from __future__ import  print_function
import pickle
import os.path
import os
import datetime
import smtplib
from socket import IPV6_RTHDR_TYPE_0
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

picklePath = 'tokens.pickle'
credsPath =  'token.json'
dt = datetime.datetime.now() - datetime.timedelta(days =1)
SCOPES = [ 'https://www.googleapis.com/auth/spreadsheets.readonly' ]

SPREADSHEET_ID = 'abcdefghijklmnopqrstuvwxyz_123456'

def main():
    creds = None
    if os.path.exists(picklePath):
        with open(picklePath, 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(credsPath, SCOPES)
            creds = flow.run_local_server(port = 0)
        #Save the creds for next run
        with open(picklePath, 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    #Calling the sheet API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId = SPREADSHEET_ID, range="Sheet1!A:E").execute()
    if not result:
        print("ERROR: Spreadsheet parsing failed","Unable to Parse Spreadsheet !!!.")

    vals = result.get('values',[])
    if not vals:
        print("NO DATA FOUND")
    else:
        try:
            for r in (vals):
                # Printing first row's 5 column values, keep adding r{INDEX}
                print(f'{r[0]}|{r[1]}|{r[2]}|{r[3]}|{r[4]}')
        except (IOError, ValueError, EOFError) as e:
                print(f'Unable to parse? {e}')

if __name__ == '__main__':
    main()
