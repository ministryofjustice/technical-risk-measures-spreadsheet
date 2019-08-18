from conditional_formatting import date_in_past_condition, \
    date_equal_or_earlier_than_condition, date_later_than_condition
from helpers import a1_to_range, red_background, amber_background, \
    green_background, add_conditional_formatting_request, light_red_background, \
    light_amber_background, light_green_background


def remove_all_user_entered_formatting_request(sheet_id):
    """
    Do this at the start, to clear any formatting that's been copied+pasted.

    This is a single request that can be done as part of the main batch, unlike
    deleting the conditional formatting rules.
    """
    return {
        "repeatCell": {
            "range": a1_to_range('A1:AD1000', sheet_id),
            "cell": {
                "userEnteredFormat": {},
            },
            "fields": "userEnteredFormat"
        }
    }


def merge_header_row_cells(sheet_id):
    cell_ranges = [
        a1_to_range('B1:D1', sheet_id),
        a1_to_range('G1:I1', sheet_id),
        a1_to_range('J1:U1', sheet_id),
        a1_to_range('V1:AB1', sheet_id),
        a1_to_range('AC1:AD1', sheet_id),
    ]

    requests = [
        {
            'mergeCells': {
                'range': cell_range,
                'mergeType': 'MERGE_ALL'
            }
        }
        for cell_range in cell_ranges
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
        },
        {
            'updateCells': {
                'fields': '*',
                'range': a1_to_range('AC1', sheet_id),
                'rows': [
                    {'values': [
                        {'userEnteredValue': {'stringValue': 'Update details'}}
                    ]}
                ]
            }
        },
    ]

    return requests


def write_second_header_row_request(sheet_id):
    """
    Write the second header row values.
    """
    cell_range = a1_to_range('A2:AD2', sheet_id)
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
        'Application has a backup & recovery strategy?',
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
        'Are we actively preventing degradation over time?',
        'Updated by',
        'Updated on',
    ]
    rows = [
        {'values': [
            {'userEnteredValue': {'stringValue': value}} for value in values
        ]}
    ]

    return {
        'updateCells': {
            'fields': '*',
            'range': cell_range,
            'rows': rows
        }
    }


def set_data_validation_number_of_people_request(sheet_id):
    cell_range = a1_to_range('G3:H1000', sheet_id)
    request = {
        "setDataValidation": {
            "range": cell_range,
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
    cell_range = a1_to_range('I3:I1000', sheet_id)
    request = {
        "setDataValidation": {
            "range": cell_range,
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
    cell_range = a1_to_range('J3:S1000', sheet_id)
    request = {
        "setDataValidation": {
            "range": cell_range,
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
    cell_range = a1_to_range('T3:U1000', sheet_id)
    request = {
        "setDataValidation": {
            "range": cell_range,
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
    cell_range = a1_to_range('V3:AA1000', sheet_id)
    request = {
        "setDataValidation": {
            "range": cell_range,
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
    cell_range = a1_to_range('AB3:AB1000', sheet_id)
    request = {
        "setDataValidation": {
            "range": cell_range,
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


def set_data_validation_updated_on_date_request(sheet_id):
    cell_range = a1_to_range('AD3:AD1000', sheet_id)
    request = {
        "setDataValidation": {
            "range": cell_range,
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


def format_dates(sheet_id):
    cell_ranges = [
        a1_to_range('V3:AA1000', sheet_id),
        a1_to_range('AD3:AD1000', sheet_id),
    ]
    requests = [
        {
            "repeatCell": {
                "range": cell_range,
                "cell": {
                    "userEnteredFormat": {
                        "numberFormat": {
                            "type": "DATE",
                        }
                    },
                },
                "fields": "userEnteredFormat(numberFormat)"
            }
        } for cell_range in cell_ranges
    ]

    return requests


def set_borders(sheet_id):
    cell_ranges = [
        a1_to_range('B1:B1000', sheet_id),
        a1_to_range('E1:E1000', sheet_id),
        a1_to_range('G1:G1000', sheet_id),
        a1_to_range('J1:J1000', sheet_id),
        a1_to_range('V1:V1000', sheet_id),
        a1_to_range('AC1:AC1000', sheet_id),
    ]
    requests = [
        {
            "updateBorders": {
                "range": cell_range,
                "left": {
                    "color": {
                        "red": 0.0,
                        "green": 0.0,
                        "blue": 0.0
                    },
                    "style": "SOLID_MEDIUM"
                },
            }
        } for cell_range in cell_ranges
    ]

    return requests


def set_default_green_background_tech_atrophy_summaries_request(sheet_id):
    """
    The tech and atrophy summary columns have a green background by default.

    Conditional formatting rules override this for red and amber.

    We need to pass values for every cell in the range we're updating, hence the
    two list comprehensions.
    """
    cell_range = a1_to_range('C3:D1000', sheet_id)
    rows = [
        {
            'values': [
                {'userEnteredFormat':
                    {
                        "backgroundColor": green_background['backgroundColor'],
                    }
                } for n in range(2)
            ]
        } for n in range(997)
    ]

    return {
        'updateCells': {
            'fields': '*',
            'range': cell_range,
            'rows': rows
        }
    }


def add_conditional_formatting_people_red_summary_request(sheet_id):
    index = 0
    cell_range = a1_to_range('B3:B1000', sheet_id)
    formula = "=AND(EQ(G3,0), EQ(H3,0), EQ(I3,FALSE))"
    values = [{"userEnteredValue": formula}]

    return add_conditional_formatting_request(index, cell_range, values, red_background)


def add_conditional_formatting_people_amber_summary_request(sheet_id):
    index = 1
    cell_range = a1_to_range('B3:B1000', sheet_id)
    formula = "=OR(AND(EQ(G3,1), EQ(H3,0), EQ(I3,FALSE)), AND(EQ(G3,0), (H3>=1), EQ(I3,FALSE)), AND(EQ(G3,0), EQ(H3,0), EQ(I3,TRUE)), AND(EQ(G3,0), (H3>=1), EQ(I3,TRUE)), AND(EQ(G3,1), EQ(H3,1), EQ(I3,FALSE)))"
    values = [{"userEnteredValue": formula}]

    return add_conditional_formatting_request(index, cell_range, values, amber_background)


def add_conditional_formatting_people_green_summary_request(sheet_id):
    index = 2
    cell_range = a1_to_range('B3:B1000', sheet_id)
    formula = "=OR((G3>=2), AND(EQ(G3,1), H3>=2), AND(EQ(G3,1), EQ(I3,TRUE)))"
    values = [{"userEnteredValue": formula}]

    return add_conditional_formatting_request(index, cell_range, values, green_background)


def add_conditional_formatting_tech_summary_red(sheet_id):
    requests = []

    # Tech binary criteria - ease and risk of making changes
    index = 3
    cell_range = a1_to_range('C3:C1000', sheet_id)
    formula = "=COUNTIF(J3:R3,\"FALSE\") >= 2"
    values = [{"userEnteredValue": formula}]
    requests.append(add_conditional_formatting_request(index, cell_range, values, red_background))

    # Understanding of security
    index = 4
    cell_range = a1_to_range('C3:C1000', sheet_id)
    formula = "=EQ(S3, FALSE)"
    values = [{"userEnteredValue": formula}]
    requests.append(add_conditional_formatting_request(index, cell_range, values, red_background))

    # Number of medium risks
    index = 5
    cell_range = a1_to_range('C3:C1000', sheet_id)
    formula = "=T3 > 5"
    values = [{"userEnteredValue": formula}]
    requests.append(add_conditional_formatting_request(index, cell_range, values, red_background))

    # Number of high risks
    index = 6
    cell_range = a1_to_range('C3:C1000', sheet_id)
    formula = "=U3 >= 1"
    values = [{"userEnteredValue": formula}]
    requests.append(add_conditional_formatting_request(index, cell_range, values, red_background))

    return requests


def add_conditional_formatting_tech_amber_summary_request(sheet_id):
    # Number of medium risks
    index = 7
    cell_range = a1_to_range('C3:C1000', sheet_id)
    formula = "=AND(T3 >= 2, T3 <= 5)"
    values = [{"userEnteredValue": formula}]

    return add_conditional_formatting_request(index, cell_range, values, amber_background)


def add_conditional_formatting_atrophy_summary_red(sheet_id):
    requests = []

    # Licences expire - red if date in past
    index = 8
    cell_range = a1_to_range('D3:D1000', sheet_id)

    column = 'V'
    formula = date_in_past_condition(column)
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, red_background))

    # Dependencies go out of support - red if date in past
    index = 9
    cell_range = a1_to_range('D3:D1000', sheet_id)
    column = 'W'
    formula = date_in_past_condition(column)
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, red_background))

    # General dependency updates last applied - red if date more than a year ago
    index = 10
    cell_range = a1_to_range('D3:D1000', sheet_id)
    formula = date_equal_or_earlier_than_condition('X', days_in_future=-365)
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, red_background))

    # Oldest unapplied security patches released - red if date more than 6 months ago
    index = 11
    cell_range = a1_to_range('D3:D1000', sheet_id)
    formula = date_equal_or_earlier_than_condition('Y', days_in_future=(6 * 30 * -1))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, red_background))

    # Support contract expire - red if date in past
    index = 12
    cell_range = a1_to_range('D3:D1000', sheet_id)
    column = 'Z'
    formula = date_in_past_condition(column)
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, red_background))

    # Relevant legislation changes come into effect - red if date less than
    # 3 months in the future (or already past)
    index = 13
    cell_range = a1_to_range('D3:D1000', sheet_id)
    formula = date_equal_or_earlier_than_condition('AA', days_in_future=(3 * 30))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, red_background))

    return requests


def add_conditional_formatting_atrophy_summary_amber(sheet_id):
    requests = []

    # Licences expire - amber if date less than 6 months in the future (or already past)
    index = 14
    cell_range = a1_to_range('D3:D1000', sheet_id)
    formula = date_equal_or_earlier_than_condition('V', days_in_future=(6 * 30))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, amber_background))

    # Dependencies go out of support - amber if date less than 6 months in the future (or already past)
    index = 15
    cell_range = a1_to_range('D3:D1000', sheet_id)
    formula = date_equal_or_earlier_than_condition('W', days_in_future=(6 * 30))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, amber_background))

    # General dependency updates last applied - amber if date more than 6 months ago
    index = 16
    cell_range = a1_to_range('D3:D1000', sheet_id)
    formula = date_equal_or_earlier_than_condition('X', days_in_future=(6 * 30 * -1))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, amber_background))

    # Oldest unapplied security patches released - amber if date more than 3 months ago
    index = 17
    cell_range = a1_to_range('D3:D1000', sheet_id)
    formula = date_equal_or_earlier_than_condition('Y', days_in_future=(3 * 30 * -1))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, amber_background))

    # Support contract expire - amber if date less than 9 months in the future (or already past)
    index = 18
    cell_range = a1_to_range('D3:D1000', sheet_id)
    formula = date_equal_or_earlier_than_condition('Z', days_in_future=(9 * 30))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, amber_background))

    # Relevant legislation changes come into effect - amber if date less than 6 months in the future (or already past)
    index = 19
    cell_range = a1_to_range('D3:D1000', sheet_id)
    formula = date_equal_or_earlier_than_condition('AA', days_in_future=(6 * 30))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, amber_background))

    # Actively preventing degradation over time? - amber if false
    index = 20
    cell_range = a1_to_range('D3:D1000', sheet_id)
    formula = "=EQ(AB3,FALSE)"
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, amber_background))

    return requests


def add_conditional_formatting_tech_individual_criteria_red(sheet_id):
    requests = []

    # Binary tech criteria - red if false
    index = 21
    cell_range = a1_to_range('J3:S1000', sheet_id)
    formula = "=EQ(J3,FALSE)"
    values = [{"userEnteredValue": formula}]
    requests.append(add_conditional_formatting_request(index, cell_range, values, light_red_background))

    # Number of medium risks - red if more than 5
    index = 22
    cell_range = a1_to_range('T3:T1000', sheet_id)
    formula = "=T3 > 5"
    values = [{"userEnteredValue": formula}]
    requests.append(add_conditional_formatting_request(index, cell_range, values, light_red_background))

    # Number of high risks - red if one or more
    index = 23
    cell_range = a1_to_range('U3:U1000', sheet_id)
    formula = "=U3 >= 1"
    values = [{"userEnteredValue": formula}]
    requests.append(add_conditional_formatting_request(index, cell_range, values, light_red_background))

    return requests


def add_conditional_formatting_tech_individual_criteria_amber_request(sheet_id):
    # Number of medium risks - amber if between 2 and 5 inclusive
    index = 24
    cell_range = a1_to_range('T3:T1000', sheet_id)
    formula = "=AND(T3 >= 2, T3 <= 5)"
    values = [{"userEnteredValue": formula}]

    return add_conditional_formatting_request(index, cell_range, values, light_amber_background)


def add_conditional_formatting_tech_individual_criteria_green(sheet_id):
    requests = []

    # Binary tech criteria - green if true
    index = 25
    cell_range = a1_to_range('J3:S1000', sheet_id)
    formula = "=EQ(J3,TRUE)"
    values = [{"userEnteredValue": formula}]
    requests.append(add_conditional_formatting_request(index, cell_range, values, light_green_background))

    # Number of medium risks - green if less than 2 and not blank
    index = 26
    cell_range = a1_to_range('T3:T1000', sheet_id)
    formula = "=AND(NOT(ISBLANK(T3)), (T3 < 2))"
    values = [{"userEnteredValue": formula}]
    requests.append(add_conditional_formatting_request(index, cell_range, values, light_green_background))

    # Number of high risks - green if 0 and not blank
    index = 27
    cell_range = a1_to_range('U3:U1000', sheet_id)
    formula = "=AND(NOT(ISBLANK(U3)), EQ(U3,0))"
    values = [{"userEnteredValue": formula}]
    requests.append(add_conditional_formatting_request(index, cell_range, values, light_green_background))

    return requests


def add_conditional_formatting_atrophy_individual_criteria_red(sheet_id):
    requests = []

    # Licences expire - red if date in past
    index = 28
    cell_range = a1_to_range('V3:V1000', sheet_id)

    column = 'V'
    formula = date_in_past_condition(column)
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_red_background))

    # Dependencies go out of support - red if date in past
    index = 29
    cell_range = a1_to_range('W3:W1000', sheet_id)

    column = 'W'
    formula = date_in_past_condition(column)
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_red_background))

    # General dependency updates last applied - red if date more than a year ago
    index = 30
    cell_range = a1_to_range('X3:X1000', sheet_id)

    column = 'X'
    formula = date_equal_or_earlier_than_condition(column, days_in_future=-365)
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_red_background))

    # Oldest unapplied security patches released - red if date more than 6 months ago
    index = 31
    cell_range = a1_to_range('Y3:Y1000', sheet_id)

    column = 'Y'
    formula = date_equal_or_earlier_than_condition(column, days_in_future=(6 * 30 * -1))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_red_background))

    # Support contract expire - red if date in past
    index = 32
    cell_range = a1_to_range('Z3:Z1000', sheet_id)

    column = 'Z'
    formula = date_in_past_condition(column)
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_red_background))

    # Relevant legislation changes come into effect - red if date less than
    # 3 months in the future (or already past)
    index = 33
    cell_range = a1_to_range('AA3:AA1000', sheet_id)

    column = 'AA'
    formula = date_equal_or_earlier_than_condition(column, days_in_future=(3 * 30))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_red_background))

    return requests


def add_conditional_formatting_atrophy_individual_criteria_amber(sheet_id):
    requests = []

    # Licences expire - amber if date less than 6 months in the future (or already past)
    index = 34
    cell_range = a1_to_range('V3:V1000', sheet_id)

    column = 'V'
    formula = date_equal_or_earlier_than_condition(column, days_in_future=(6 * 30))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_amber_background))

    # Dependencies go out of support - amber if date less than 6 months in the future (or already past)
    index = 35
    cell_range = a1_to_range('W3:W1000', sheet_id)

    column = 'W'
    formula = date_equal_or_earlier_than_condition(column, days_in_future=(6 * 30))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_amber_background))

    # General dependency updates last applied - amber if date more than 6 months ago
    index = 36
    cell_range = a1_to_range('X3:X1000', sheet_id)

    column = 'X'
    formula = date_equal_or_earlier_than_condition(column, days_in_future=(6 * 30 * -1))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_amber_background))

    # Oldest unapplied security patches released - amber if date more than 3 months ago
    index = 37
    cell_range = a1_to_range('Y3:Y1000', sheet_id)

    column = 'Y'
    formula = date_equal_or_earlier_than_condition(column, days_in_future=(3 * 30 * -1))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_amber_background))

    # Support contract expire - amber if date less than 9 months in the future (or already past)
    index = 38
    cell_range = a1_to_range('Z3:Z1000', sheet_id)

    column = 'Z'
    formula = date_equal_or_earlier_than_condition(column, days_in_future=(9 * 30))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_amber_background))

    # Relevant legislation changes come into effect - amber if date less than 6 months in the future (or already past)
    index = 39
    cell_range = a1_to_range('AA3:AA1000', sheet_id)

    column = 'AA'
    formula = date_equal_or_earlier_than_condition(column, days_in_future=(6 * 30))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_amber_background))

    # Actively preventing degradation over time? - amber if false
    index = 40
    cell_range = a1_to_range('AB3:AB1000', sheet_id)
    formula = "=EQ(AB3,FALSE)"
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_amber_background))

    return requests


def add_conditional_formatting_atrophy_individual_criteria_green(sheet_id):
    requests = []

    # Licences expire - green if date more than 6 months in the future
    index = 41
    cell_range = a1_to_range('V3:V1000', sheet_id)

    column = 'V'
    formula = date_later_than_condition(column, days_in_future=(6 * 30))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_green_background))

    # Dependencies go out of support - green if date more than 6 months in the future
    index = 42
    cell_range = a1_to_range('W3:W1000', sheet_id)

    column = 'W'
    formula = date_later_than_condition(column, days_in_future=(6 * 30))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_green_background))

    # General dependency updates last applied - green if date less than 6 months ago
    index = 43
    cell_range = a1_to_range('X3:X1000', sheet_id)

    column = 'X'
    formula = date_later_than_condition(column, days_in_future=(6 * 30 * -1))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_green_background))

    # Oldest unapplied security patches released - green if date less than 3 months ago
    index = 44
    cell_range = a1_to_range('Y3:Y1000', sheet_id)

    column = 'Y'
    formula = date_later_than_condition(column, days_in_future=(3 * 30 * -1))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_green_background))

    # Support contract expire - green if date more than 9 months in the future
    index = 45
    cell_range = a1_to_range('Z3:Z1000', sheet_id)

    column = 'Z'
    formula = date_later_than_condition(column, days_in_future=(9 * 30))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_green_background))

    # Relevant legislation changes come into effect - green if date more than 6 months in the future
    index = 46
    cell_range = a1_to_range('AA3:AA1000', sheet_id)

    column = 'AA'
    formula = date_later_than_condition(column, days_in_future=(6 * 30))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_green_background))

    # Actively preventing degradation over time? - green if true
    index = 47
    cell_range = a1_to_range('AB3:AB1000', sheet_id)
    formula = "=EQ(AB3,TRUE)"
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_green_background))

    # All date-based atrophy criteria - green if not applicable (entered as "N/A")
    index = 48
    cell_range = a1_to_range('V3:AA1000', sheet_id)
    formula = "=EQ(V3, \"N/A\")"
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_green_background))

    return requests


def add_conditional_formatting_update_details_red(sheet_id):
    requests = []

    # Both update details columns - red if blank
    index = 49
    cell_range = a1_to_range('AC3:AD1000', sheet_id)
    formula = "=ISBLANK(AC3)"
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_red_background))

    # Updated on - red if in the future
    index = 50
    cell_range = a1_to_range('AD3:AD1000', sheet_id)

    column = 'AD'
    formula = date_later_than_condition(column, days_in_future=(0))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_red_background))

    # Updated on - red if more than 3 months ago
    index = 51
    cell_range = a1_to_range('AD3:AD1000', sheet_id)

    column = 'AD'
    formula = date_equal_or_earlier_than_condition(column, days_in_future=(3 * 30 * -1))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_red_background))

    return requests


def add_conditional_formatting_update_details_amber_request(sheet_id):
    # Updated on - amber if more than 2 months ago
    index = 52
    cell_range = a1_to_range('AD3:AD1000', sheet_id)

    column = 'AD'
    formula = date_equal_or_earlier_than_condition(column, days_in_future=(2 * 30 * -1))
    values = [{"userEnteredValue": formula}]

    return add_conditional_formatting_request(index, cell_range, values, light_amber_background)


def add_conditional_formatting_update_details_green_request(sheet_id):
    # Updated on - green if less than 2 months ago
    index = 53
    cell_range = a1_to_range('AD3:AD1000', sheet_id)

    column = 'AD'
    formula = date_later_than_condition(column, days_in_future=(2 * 30 * -1))
    values = [{"userEnteredValue": formula}]

    return add_conditional_formatting_request(index, cell_range, values, light_green_background)


def freeze_header_rows_and_summary_columns_request(sheet_id):
    return {
        "updateSheetProperties": {
            "properties": {
                "sheetId": sheet_id,
                "gridProperties": {
                    "frozenRowCount": "2",
                    "frozenColumnCount": "4",
                },
            },
            "fields": "gridProperties.frozenRowCount, gridProperties.frozenColumnCount"
        }
    }


def bold_and_wrap_text_in_header_rows_and_service_column(sheet_id):
    cell_ranges = [
        a1_to_range('A1:AD2', sheet_id),
        a1_to_range('A3:A1000', sheet_id),
    ]
    requests = [
        {
            "repeatCell": {
                "range": cell_range,
                "cell": {
                    "userEnteredFormat": {
                        "textFormat": {
                            "bold": True,
                        },
                        "wrapStrategy": "WRAP",
                    },
                },
                "fields": "userEnteredFormat(textFormat, wrapStrategy)"
            }
        } for cell_range in cell_ranges
    ]

    return requests


def set_column_widths(sheet_id):
    # Default column width is 100 pixels
    a1_with_widths = [
        ('A1:A1000', 230),
        ('E1:E1000', 160),
        ('F1:F1000', 200),
        ('G1:G1000', 95),
        ('H1:H1000', 85),
        ('I1:I1000', 65),
        ('AC1:AC1000', 120),
    ]

    # This API call wants the range in a different format
    def format_a1_ranges_with_column_dimension(a1, sheet_id):
        full_range = a1_to_range(a1, sheet_id)
        return {
            "sheetId": sheet_id,
            "dimension": "COLUMNS",
            "startIndex": full_range["startColumnIndex"],
            "endIndex": full_range["endColumnIndex"]
        }

    requests = [
        {
            "updateDimensionProperties": {
                "range": format_a1_ranges_with_column_dimension(a1, sheet_id),
                "properties": {
                    "pixelSize": width
                },
                "fields": "pixelSize"
            }
        } for a1, width in a1_with_widths
    ]

    return requests


def all_requests_in_order(sheet_id):
    """
    Return all the real requests, in the right order for applying as a batch.

    Uses extend and append to construct a flat list of requests, since some
    functions return multiple similar request objects.
    """
    requests = []

    requests.append(remove_all_user_entered_formatting_request(sheet_id))
    requests.extend(merge_header_row_cells(sheet_id))
    requests.extend(write_first_header_rows(sheet_id))
    requests.append(write_second_header_row_request(sheet_id))
    requests.append(set_data_validation_number_of_people_request(sheet_id))
    requests.append(set_data_validation_managed_service_request(sheet_id))
    requests.append(set_data_validation_tech_boolean_criteria_request(sheet_id))
    requests.append(set_data_validation_number_of_security_risks_request(sheet_id))
    requests.append(set_data_validation_atrophy_dates_request(sheet_id))
    requests.append(set_data_validation_preventing_degradation_request(sheet_id))
    requests.append(set_data_validation_updated_on_date_request(sheet_id))
    requests.extend(format_dates(sheet_id))
    requests.extend(set_borders(sheet_id))
    requests.append(set_default_green_background_tech_atrophy_summaries_request(sheet_id))
    requests.append(add_conditional_formatting_people_red_summary_request(sheet_id))
    requests.append(add_conditional_formatting_people_amber_summary_request(sheet_id))
    requests.append(add_conditional_formatting_people_green_summary_request(sheet_id))
    requests.extend(add_conditional_formatting_tech_summary_red(sheet_id))
    requests.append(add_conditional_formatting_tech_amber_summary_request(sheet_id))
    requests.extend(add_conditional_formatting_atrophy_summary_red(sheet_id))
    requests.extend(add_conditional_formatting_atrophy_summary_amber(sheet_id))
    requests.extend(add_conditional_formatting_tech_individual_criteria_red(sheet_id))
    requests.append(add_conditional_formatting_tech_individual_criteria_amber_request(sheet_id))
    requests.extend(add_conditional_formatting_tech_individual_criteria_green(sheet_id))
    requests.extend(add_conditional_formatting_atrophy_individual_criteria_red(sheet_id))
    requests.extend(add_conditional_formatting_atrophy_individual_criteria_amber(sheet_id))
    requests.extend(add_conditional_formatting_atrophy_individual_criteria_green(sheet_id))
    requests.extend(add_conditional_formatting_update_details_red(sheet_id))
    requests.append(add_conditional_formatting_update_details_amber_request(sheet_id))
    requests.append(add_conditional_formatting_update_details_green_request(sheet_id))
    requests.append(freeze_header_rows_and_summary_columns_request(sheet_id))
    requests.extend(bold_and_wrap_text_in_header_rows_and_service_column(sheet_id))
    requests.extend(set_column_widths(sheet_id))

    return requests
