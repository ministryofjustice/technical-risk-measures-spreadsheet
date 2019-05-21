from helpers import a1_to_range, red_background, amber_background, \
    green_background, add_conditional_formatting_request


def test_write_request(sheet_id):
    """
    Test writing to a sheet, mainly to check we have permission to write.
    """
    range = a1_to_range('C3:D4', sheet_id)
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
        a1_to_range('B1:D1', sheet_id),
        a1_to_range('G1:I1', sheet_id),
        a1_to_range('J1:U1', sheet_id),
        a1_to_range('V1:AB1', sheet_id),
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


def write_first_header_rows(sheet_id):
    """
    Write the first header row values.

    These are isolated cells rather than a contiguous range, so specify them as
    separate requests.
    """
    requests = [
        {
            'updateCells': {
                'fields': '*',
                'range': a1_to_range('B1', sheet_id),
                'rows': [
                    {'values': [
                        {'userEnteredValue': {'stringValue': 'Summaries'}}
                    ]}
                ]
            }
        },
        {
            'updateCells': {
                'fields': '*',
                'range': a1_to_range('G1', sheet_id),
                'rows': [
                    {'values': [
                        {'userEnteredValue': {'stringValue': 'People criteria'}}
                    ]}
                ]
            }
        },
        {
            'updateCells': {
                'fields': '*',
                'range': a1_to_range('J1', sheet_id),
                'rows': [
                    {'values': [
                        {'userEnteredValue': {'stringValue': 'Tech criteria'}}
                    ]}
                ]
            }
        },
        {
            'updateCells': {
                'fields': '*',
                'range': a1_to_range('V1', sheet_id),
                'rows': [
                    {'values': [
                        {'userEnteredValue': {'stringValue': 'Atrophy criteria'}}
                    ]}
                ]
            }
        }
    ]

    return requests


def write_second_header_row_request(sheet_id):
    """
    Write the second header row values.
    """
    range = a1_to_range('A2:AB2', sheet_id)
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
        'Team has a good understanding of application\'s security?',
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


def set_data_validation_number_of_people_request(sheet_id):
    range = a1_to_range('G3:H1000', sheet_id)
    request = {
        "setDataValidation": {
            "range": range,
            "rule": {
                "showCustomUi": False,
                "strict": False,
                "condition": {
                    "values": [
                        {"userEnteredValue": "0"},
                        {"userEnteredValue": "100"},
                    ],
                    "type": "NUMBER_BETWEEN",
                },
            },
        }
    }

    return request


def set_data_validation_managed_service_request(sheet_id):
    range = a1_to_range('I3:I1000', sheet_id)
    request = {
        "setDataValidation": {
            "range": range,
            "rule": {
                "showCustomUi": False,
                "strict": False,
                "condition": {
                    "type": "BOOLEAN",
                },
            },
        }
    }

    return request


def set_data_validation_tech_boolean_criteria_request(sheet_id):
    range = a1_to_range('J3:S1000', sheet_id)
    request = {
        "setDataValidation": {
            "range": range,
            "rule": {
                "showCustomUi": False,
                "strict": False,
                "condition": {
                    "type": "BOOLEAN",
                },
            },
        }
    }

    return request


def set_data_validation_number_of_security_risks_request(sheet_id):
    range = a1_to_range('T3:U1000', sheet_id)
    request = {
        "setDataValidation": {
            "range": range,
            "rule": {
                "showCustomUi": False,
                "strict": False,
                "condition": {
                    "values": [
                        {"userEnteredValue": "0"},
                        {"userEnteredValue": "500"},
                    ],
                    "type": "NUMBER_BETWEEN",
                },
            },
        }
    }

    return request


def set_data_validation_atrophy_dates_request(sheet_id):
    range = a1_to_range('V3:AA1000', sheet_id)
    request = {
        "setDataValidation": {
            "range": range,
            "rule": {
                "showCustomUi": False,
                "strict": False,
                "condition": {
                    "type": "DATE_IS_VALID",
                },
            },
        }
    }

    return request


def set_data_validation_preventing_degradation_request(sheet_id):
    range = a1_to_range('AB3:AB1000', sheet_id)
    request = {
        "setDataValidation": {
            "range": range,
            "rule": {
                "showCustomUi": False,
                "strict": False,
                "condition": {
                    "type": "BOOLEAN",
                },
            },
        }
    }

    return request


def set_borders(sheet_id):
    ranges = [
        a1_to_range('B1:B1000', sheet_id),
        a1_to_range('E1:E1000', sheet_id),
        a1_to_range('G1:G1000', sheet_id),
        a1_to_range('J1:J1000', sheet_id),
        a1_to_range('V1:V1000', sheet_id),
    ]
    requests = [
        {
            "updateBorders": {
                "range": range,
                "left": {
                    "color": {
                        "red": 0.0,
                        "green": 0.0,
                        "blue": 0.0
                    },
                    "style": "SOLID_MEDIUM"
                },
            }
        } for range in ranges
    ]

    return requests


def add_conditional_formatting_people_red_request(sheet_id):
    index = 0
    range = a1_to_range('B3:B1000', sheet_id)
    formula = "=AND(EQ(G3,0), EQ(H3,0), EQ(I3,FALSE))"
    values = [{"userEnteredValue": formula}]

    return add_conditional_formatting_request(index, range, values, red_background)


def add_conditional_formatting_people_amber_request(sheet_id):
    index = 1
    range = a1_to_range('B3:B1000', sheet_id)
    formula = "=OR(AND(EQ(G3,1), EQ(H3,0), EQ(I3,FALSE)), AND(EQ(G3,0), (H3>=1), EQ(I3,FALSE)), AND(EQ(G3,0), EQ(H3,0), EQ(I3,TRUE)), AND(EQ(G3,1), EQ(H3,1), EQ(I3,FALSE)))"
    values = [{"userEnteredValue": formula}]

    return add_conditional_formatting_request(index, range, values, amber_background)


def add_conditional_formatting_people_green_request(sheet_id):
    index = 2
    range = a1_to_range('B3:B1000', sheet_id)
    formula = "=OR((G3>=2), AND(EQ(G3,1), H3>=2), AND(EQ(G3,1), EQ(I3,TRUE)))"
    values = [{"userEnteredValue": formula}]

    return add_conditional_formatting_request(index, range, values, green_background)


def add_conditional_formatting_tech_red(sheet_id):
    requests = []

    # Tech binary criteria - ease and risk of making changes
    index = 3
    range = a1_to_range('C3:C1000', sheet_id)
    formula = "=COUNTIF(J3:R3,\"FALSE\") >= 2"
    values = [{"userEnteredValue": formula}]
    requests.append(add_conditional_formatting_request(index, range, values, red_background))

    # Understanding of security
    index = 4
    range = a1_to_range('C3:C1000', sheet_id)
    formula = "=EQ(S3, FALSE)"
    values = [{"userEnteredValue": formula}]
    requests.append(add_conditional_formatting_request(index, range, values, red_background))

    # Number of medium risks
    index = 5
    range = a1_to_range('C3:C1000', sheet_id)
    formula = "=T3 > 5"
    values = [{"userEnteredValue": formula}]
    requests.append(add_conditional_formatting_request(index, range, values, red_background))

    # Number of high risks
    index = 6
    range = a1_to_range('C3:C1000', sheet_id)
    formula = "=U3 >= 1"
    values = [{"userEnteredValue": formula}]
    requests.append(add_conditional_formatting_request(index, range, values, red_background))

    return requests


def add_conditional_formatting_tech_amber_request(sheet_id):
    # Number of medium risks
    index = 7
    range = a1_to_range('C3:C1000', sheet_id)
    formula = "=AND(T3 >= 2, T3 <= 5)"
    values = [{"userEnteredValue": formula}]

    return add_conditional_formatting_request(index, range, values, amber_background)


def all_requests_in_order(sheet_id):
    """
    Return all the real requests, in the right order for applying as a batch.

    Uses extend and append to construct a flat list of requests, since some
    functions return multiple similar request objects.
    """
    requests = []

    requests.extend(merge_header_row_cells(sheet_id))
    requests.extend(write_first_header_rows(sheet_id))
    requests.append(write_second_header_row_request(sheet_id))
    requests.append(set_data_validation_number_of_people_request(sheet_id))
    requests.append(set_data_validation_managed_service_request(sheet_id))
    requests.append(set_data_validation_tech_boolean_criteria_request(sheet_id))
    requests.append(set_data_validation_number_of_security_risks_request(sheet_id))
    requests.append(set_data_validation_atrophy_dates_request(sheet_id))
    requests.append(set_data_validation_preventing_degradation_request(sheet_id))
    requests.extend(set_borders(sheet_id))
    requests.append(add_conditional_formatting_people_red_request(sheet_id))
    requests.append(add_conditional_formatting_people_amber_request(sheet_id))
    requests.append(add_conditional_formatting_people_green_request(sheet_id))
    requests.extend(add_conditional_formatting_tech_red(sheet_id))
    requests.append(add_conditional_formatting_tech_amber_request(sheet_id))

    return requests
