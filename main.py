import pickle
import os
import os.path
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from google.oauth2 import service_account


# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

SERVICE_ACCOUNT_FILE = os.environ['SERVICE_ACCOUNT_FILE']

SPREADSHEET_ID = os.environ['SPREADSHEET_ID']
RANGE_NAME_READ = os.environ['RANGE_NAME_READ']


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
                                range=RANGE_NAME_READ).execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')
    else:
        for row in values:
            print(row)


def try_writing(service):
    requests = []

    range = {
        "sheetId": 0,
        "startColumnIndex": 2,
        "startRowIndex": 2,
        "endColumnIndex": 4,
        "endRowIndex": 4
    }
    rows = [
        {'values': [
            {'userEnteredValue': {'stringValue': 'Can'}},
            {'userEnteredValue': {'stringValue': 'I'}}
        ]},
        {'values': [
            {'userEnteredValue': {'stringValue': 'write'}},
            {'userEnteredValue': {'stringValue': 'here?'}}
        ]}
    ]

    requests.append({
        'updateCells': {
            'fields': '*',
            'range': range,
            'rows': rows
        }
    })

    body = {
        'requests': requests
    }
    service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body=body).execute()

    print('Spreadsheet cell values updated')


def main():
    """Shows basic usage of the Sheets API.
    """
    service = build('sheets', 'v4', credentials=get_creds())

    try_reading(service)
    try_writing(service)


if __name__ == '__main__':
    main()
