def test_write_request(sheet_id):
    """
    Test writing to a sheet, mainly to check we have permission to write.
    """
    range = {
        "sheetId": sheet_id,
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

    return {
        'updateCells': {
            'fields': '*',
            'range': range,
            'rows': rows
        }
    }


def merge_header_row_cells(sheet_id):
    ranges = [
        {
            "sheetId": sheet_id,
            "startColumnIndex": 0,
            "startRowIndex": 0,
            "endColumnIndex": 1,
            "endRowIndex": 2
        },
        {
            "sheetId": sheet_id,
            "startColumnIndex": 1,
            "startRowIndex": 0,
            "endColumnIndex": 4,
            "endRowIndex": 1
        },
        {
            "sheetId": sheet_id,
            "startColumnIndex": 6,
            "startRowIndex": 0,
            "endColumnIndex": 9,
            "endRowIndex": 1
        },
    ]

    requests = [
        {
            'mergeCells': {
                'range': range,
                'mergeType': 'MERGE_ALL'
            }
        }
        for range in ranges
    ]
    return requests


def write_second_header_row_request(sheet_id):
    """
    Write the second header row values.
    """
    range = {
        "sheetId": sheet_id,
        "startColumnIndex": 0,
        "startRowIndex": 1,
        "endColumnIndex": 29,
        "endRowIndex": 2
    }
    values = [
        'Service',
        'People summary',
        'Tech summary',
        'Atrophy summary',
        'Service area',
        'Notes',
        'Number of civil servants',
        'Number of contractors',
        'Managed service?',
        'Team can deploy in working hours?',
        'Application has automated tests?',
        'Application has build promotion through test environments?',
        'Team who owns the app own the complete deployment?',
        'Can deploy multiple times a day?',
        'Alerts for an incident are well understood and sent to a known place?',
        'Logs are aggregated and searchable?',
        'Application has a backup/recovery strategy?',
        'Code base can be easily changed?',
        'Has a risk register?',
        'When was the risk register last updated?',
        'Number of medium security risks',
        'Number of high security risks',
        'When do licences expire?',
        'When will major dependencies go out of support?',
        'When were general dependency updates last applied?',
        'When were the oldest unapplied security patches released?',
        'When does the support contract expire?',
        'When do the next relevant legislation changes come into effect?',
        'Are we actively preventing degradation over time?'
    ]
    rows = [
        {'values': [
            {'userEnteredValue': {'stringValue': value}} for value in values
        ]}
    ]

    return {
        'updateCells': {
            'fields': '*',
            'range': range,
            'rows': rows
        }
    }


def all_requests_in_order(sheet_id):
    """
    Return all the real requests, in the right order for applying as a batch.

    Uses extend and append to construct a flat list of requests, since some
    functions return multiple similar request objects.
    """
    requests = []

    requests.extend(merge_header_row_cells(sheet_id))
    requests.append(write_second_header_row_request(sheet_id))

    return requests
