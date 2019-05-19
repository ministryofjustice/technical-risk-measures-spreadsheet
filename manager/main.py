import pickle
import os
import os.path
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from google.oauth2 import service_account

import batch_requests


# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

SERVICE_ACCOUNT_FILE = os.environ['SERVICE_ACCOUNT_FILE']

SPREADSHEET_ID = os.environ['SPREADSHEET_ID']
SHEET_ID = os.environ['SHEET_ID']


def get_creds():
    creds = None

    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    token_filename = 'token.pickle'
    if os.path.exists(token_filename):
        with open(token_filename, 'rb') as token:
            creds = pickle.load(token)

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            creds = service_account.Credentials.from_service_account_file(
                        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        # Save the credentials for the next run
        with open(token_filename, 'wb') as token:
            pickle.dump(creds, token)

    return creds


def try_reading(service):
    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                range='Sheet1!A1:A2').execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')
    else:
        for row in values:
            print(row)


def apply_batch(service, requests):
    """
    Apply multiple updates in a single batch.

    `requests` should be a list of Request objects representing batch update
    operations, as described here:

        https://developers.google.com/sheets/api/guides/batchupdate
    """
    body = {
        'requests': requests
    }
    service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body=body).execute()

    print('Spreadsheet updated')


def try_writing(service):
    requests = [batch_requests.test_write_request(SHEET_ID)]
    apply_batch(service, requests)


def main():
    """Apply all the specified changes to a sheet.
    """
    service = build('sheets', 'v4', credentials=get_creds())

    try_reading(service)
    try_writing(service)

    requests = batch_requests.all_requests_in_order(SHEET_ID)
    apply_batch(service, requests)


if __name__ == '__main__':
    main()
