from helpers import a1_to_range


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
        a1_to_range('J1:V1', sheet_id),
        a1_to_range('W1:AC1', sheet_id),
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
                'range': a1_to_range('W1', sheet_id),
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
    range = a1_to_range('A2:AC2', sheet_id)
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


def set_data_validation_risk_register_updated_request(sheet_id):
    range = a1_to_range('T3:T1000', sheet_id)
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


def set_data_validation_number_of_security_risks_request(sheet_id):
    range = a1_to_range('U3:V1000', sheet_id)
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
    range = a1_to_range('W3:AB1000', sheet_id)
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
    range = a1_to_range('AC3:AC1000', sheet_id)
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
        a1_to_range('W1:W1000', sheet_id),
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
    request = {
        "addConditionalFormatRule": {
            "index": 0,
            "rule": {
                "ranges": [
                    a1_to_range('B3:B1000', sheet_id)
                ],
                "booleanRule": {
                    "condition": {
                        "values": [
                            {"userEnteredValue": "=AND(EQ(G3,0), EQ(H3,0), EQ(I3,FALSE))"},
                        ],
                        "type": "CUSTOM_FORMULA",
                    },
                    "format": {
                        "backgroundColor": {
                            "blue": 0,
                            "green": 0,
                            "red": 1,
                        }
                    }
                }
            }
        }
    }

    return request


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
    requests.append(set_data_validation_risk_register_updated_request(sheet_id))
    requests.append(set_data_validation_number_of_security_risks_request(sheet_id))
    requests.append(set_data_validation_atrophy_dates_request(sheet_id))
    requests.append(set_data_validation_preventing_degradation_request(sheet_id))
    requests.extend(set_borders(sheet_id))
    requests.append(add_conditional_formatting_people_red_request(sheet_id))

    return requests
