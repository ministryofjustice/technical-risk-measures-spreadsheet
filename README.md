Python scripts to set up spreadsheets for tracking [MOJ's technical risk
measures](https://github.com/ministryofjustice/technical-risk-measures), using the [Google Sheets API](https://developers.google.com/sheets/api/).

## Development setup

This project uses [Pipenv](https://docs.pipenv.org/en/latest/basics/) to manage
its dependencies.

Prequisites:

- Python 3.7 installed and available on your `PATH`
- Pipenv installed for that Python version

Run these commands to set up a virtualenv for this project and to install the
dependencies specified in the `Pipfile`:

```
$ pipenv --python 3.7
$ pipenv install --dev
```

To run the tests:

```
$ pipenv run pytest
```

## Setting up a technical risk measures spreadsheet

- Create a new spreadsheet in Google Sheets
- Add 3 columns to it (so the columns go up to AC)
- Copy the `.env.example` file to a new file called `.env`
- Copy the spreadsheet's id (from its URL) into `.env` as `SPREADSHEET_ID`
- From the [Google APIs Console](https://console.developers.google.com/apis/dashboard),
create a new project and:
    - Enable the Google Sheets API in your new project
    - Under "Credentials", create a service account key for your project. When
prompted, create a new service account without any roles for it.
- Move the json file that was downloaded when you created the service account
to the root of the repo. This file contains the private key for authenticating
with the API as this service account - `*.json` is in the `.gitignore` so it
should be safe to keep it here as long as you keep the file extension on it.
- Add the name of that key file into `.env` as `SERVICE_ACCOUNT_FILE`
- In your spreadsheet, give edit access to the email address of your service
account - you can find this as `"client_email"` in the json key file

Now you're ready to run this script!

```
pipenv run python manager/main.py
```
