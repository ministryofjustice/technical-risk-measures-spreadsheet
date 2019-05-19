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
