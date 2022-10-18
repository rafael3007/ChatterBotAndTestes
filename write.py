
from __future__ import print_function

import os.path
import google.auth

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '1onYLUGKteVoHdXDQfRGpp4TKDLTp9p3oROO9jsWyDKo'
SAMPLE_RANGE_NAME = 'Página2!A2:C2'

def append_values(spreadsheet_id, range_name, value_input_option,
                  _values):
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    try:
        service = build('sheets', 'v4', credentials=creds)

        values = [
            [
                # Cell values ...
                'F', 'B'
            ],[
                'vyuyu','fcytcyt'
            ]
            # Additional rows ...
        ]
        body = {
            'values': values
        }
        result = service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id, range=range_name,
            valueInputOption=value_input_option, body=body).execute()
        print(f"{(result.get('updates').get('updatedCells'))} cells appended.")
        return result

    except HttpError as error:
        print(f"An error occurred: {error}")
        return error


if __name__ == '__main__':
    # Pass: spreadsheet_id, range_name value_input_option and _values)
    append_values("1onYLUGKteVoHdXDQfRGpp4TKDLTp9p3oROO9jsWyDKo",
                  "A1:C2", "USER_ENTERED",
                  [
                      ['F', 'B'],
                      ['C', 'D']
                  ])