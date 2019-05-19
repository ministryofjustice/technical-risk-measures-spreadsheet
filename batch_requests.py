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


all_requests_in_order = [
    write_second_header_row_request,
]
