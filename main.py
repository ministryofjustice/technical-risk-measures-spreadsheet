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
RANGE_NAME_WRITE = os.environ['RANGE_NAME_WRITE']


def get_creds():
    creds = None

    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            creds = service_account.Credentials.from_service_account_file(
                        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    return creds


def try_writing(service):
    values = [
        ['Can', 'I'],
        ['write', 'here?'],
    ]
    body = {
        'values': values
    }
    result = service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME_WRITE,
        valueInputOption='RAW', body=body).execute()
    print('{0} cells updated.'.format(result.get('updatedCells')))


def main():
    """Shows basic usage of the Sheets API.
    """
    service = build('sheets', 'v4', credentials=get_creds())

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

    try_writing(service)


if __name__ == '__main__':
    main()
