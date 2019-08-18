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


def delete_all_existing_conditional_formatting_rules(service):
    """
    Find out how many existing rules are on the sheet, then delete them all.

    There's no API support for passing the complete list of rules to set, sadly.
    Modifying existing rules in place would be really fiddly, mainly because
    the API only supports adding and updating operations on individual rules.
    So delete them all each time before setting the current rules.
    """
    data = service.spreadsheets().get(
        spreadsheetId=SPREADSHEET_ID,
        includeGridData=False
    ).execute()

    sheet = [s for s in data['sheets'] if s['properties']['sheetId'] == int(SHEET_ID)][0]
    number_of_rules_to_delete = len(sheet.get('conditionalFormats', []))

    if number_of_rules_to_delete > 0:
        print(f"Deleting {number_of_rules_to_delete} existing conditional formatting rules")
        requests = [
            {
                "deleteConditionalFormatRule": {
                    "sheetId": SHEET_ID,
                    "index": 0
                }
            } for n in range(number_of_rules_to_delete)
        ]
        apply_batch(service, requests)


def main():
    """Apply all the specified changes to a sheet.
    """
    service = build('sheets', 'v4', credentials=get_creds())

    delete_all_existing_conditional_formatting_rules(service)

    requests = batch_requests.all_requests_in_order(SHEET_ID)
    apply_batch(service, requests)
    print('Spreadsheet updated')


if __name__ == '__main__':
    main()
