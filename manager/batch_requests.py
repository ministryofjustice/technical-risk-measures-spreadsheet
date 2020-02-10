import os

from conditional_formatting import date_in_past_condition, \
    date_equal_or_earlier_than_condition, date_later_than_condition
from helpers import build_cell_range, red_background, amber_background, \
    green_background, add_conditional_formatting_request, light_red_background, \
    light_amber_background, light_green_background


PROTECTED_RANGE_EDITOR_USERS = os.environ['PROTECTED_RANGE_EDITOR_USERS']
PROTECTED_RANGE_EDITOR_GROUPS = os.environ['PROTECTED_RANGE_EDITOR_GROUPS']


COLUMNS = {
    'service': 'A',
    'people_summary': 'B',
    'tech_summary': 'C',
    'decay_summary': 'D',
    'service_area': 'E',
    'notes': 'F',
    'civil_servants': 'G',
    'contractors': 'H',
    'managed_service': 'I',
    'automated_tests': 'J',
    'build_promotion': 'K',
    'change_process_quick_easy_cheap': 'L',
    'deploy_quickly_easily_in_working_hours_without_downtime': 'M',
    'alerts_well_understood': 'N',
    'logs_aggregated_searchable': 'O',
    'backup_recovery_strategy': 'P',
    'good_understanding_of_security': 'Q',
    'medium_security_risks': 'R',
    'high_security_risks': 'S',
    'licences_expire': 'T',
    'major_dependencies_out_of_support': 'U',
    'oldest_unapplied_dependency_version_released': 'V',
    'support_contract_expires': 'W',
    'legislation_changes': 'X',
    'preventing_degradation_over_time': 'Y',
    'updated_by': 'Z',
    'updated_on': 'AA',
}


def remove_all_user_entered_formatting_request(sheet_id):
    """
    Do this at the start, to clear any formatting that's been copied+pasted.

    This is a single request that can be done as part of the main batch, unlike
    deleting the conditional formatting rules.
    """
    cell_range = build_cell_range(
        COLUMNS['service'], 1,
        COLUMNS['updated_on'], 1000,
        sheet_id=sheet_id
    )
    return {
        "repeatCell": {
            "range": cell_range,
            "cell": {
                "userEnteredFormat": {},
            },
            "fields": "userEnteredFormat"
        }
    }


def protect_header_rows_request(sheet_id, users_list, groups_list):
    def split_email_address_lists(l):
        return [a.strip() for a in l.split(',')]

    users = split_email_address_lists(users_list)
    groups = split_email_address_lists(groups_list)

    cell_range = build_cell_range(
        COLUMNS['service'], 1,
        COLUMNS['updated_on'], 2,
        sheet_id=sheet_id
    )

    return {
        "addProtectedRange": {
            "protectedRange": {
                "range": cell_range,
                "description": "Header rows",
                "warningOnly": False,
                "editors": {
                    "users": users,
                    "groups": groups,
                    "domainUsersCanEdit": False,
                }
            }
        }
    }


def merge_header_row_cells(sheet_id):
    cell_ranges = [
        build_cell_range(
            COLUMNS['people_summary'], 1,
            COLUMNS['decay_summary'], 1,
            sheet_id=sheet_id
        ),
        build_cell_range(
            COLUMNS['civil_servants'], 1,
            COLUMNS['managed_service'], 1,
            sheet_id=sheet_id
        ),
        build_cell_range(
            COLUMNS['automated_tests'], 1,
            COLUMNS['high_security_risks'], 1,
            sheet_id=sheet_id
        ),
        build_cell_range(
            COLUMNS['licences_expire'], 1,
            COLUMNS['preventing_degradation_over_time'], 1,
            sheet_id=sheet_id
        ),
        build_cell_range(
            COLUMNS['updated_by'], 1,
            COLUMNS['updated_on'], 1,
            sheet_id=sheet_id
        ),
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
                'range': build_cell_range(COLUMNS['people_summary'], 1, sheet_id=sheet_id),
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
                'range': build_cell_range(COLUMNS['civil_servants'], 1, sheet_id=sheet_id),
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
                'range': build_cell_range(COLUMNS['automated_tests'], 1, sheet_id=sheet_id),
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
                'range': build_cell_range(COLUMNS['licences_expire'], 1, sheet_id=sheet_id),
                'rows': [
                    {'values': [
                        {'userEnteredValue': {'stringValue': 'Decay criteria'}}
                    ]}
                ]
            }
        },
        {
            'updateCells': {
                'fields': '*',
                'range': build_cell_range(COLUMNS['updated_by'], 1, sheet_id=sheet_id),
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
    cell_range = build_cell_range(
        COLUMNS['service'], 2,
        COLUMNS['updated_on'], 2,
        sheet_id=sheet_id
    )
    values = [
        'Service',
        'People summary',
        'Tech summary',
        'Decay summary',
        'Service area',
        'Notes',
        'Number of civil servants',
        'Number of contractors',
        'Managed service?',
        'Application has automated tests?',
        'Application has build promotion through test environments?',
        'The overall process for making a change to the service is quick, easy and cheap for us?',
        'Changes to the service can be deployed quickly and easily during working hours, without downtime for users?',
        'Alerts for an incident are well understood and sent to a known place?',
        'Logs are aggregated and searchable?',
        'Application has a backup & recovery strategy?',
        'Team has a good understanding of application\'s security?',
        'Number of medium security risks',
        'Number of high security risks',
        'When do licences expire?',
        'When will major dependencies go out of support?',
        'When was the oldest unapplied version of any dependency released?',
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
    cell_range = build_cell_range(
        COLUMNS['civil_servants'], 3,
        COLUMNS['contractors'], 1000,
        sheet_id=sheet_id
    )
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
    cell_range = build_cell_range(
        COLUMNS['managed_service'], 3,
        COLUMNS['managed_service'], 1000,
        sheet_id=sheet_id
    )
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
    cell_range = build_cell_range(
        COLUMNS['automated_tests'], 3,
        COLUMNS['good_understanding_of_security'], 1000,
        sheet_id=sheet_id
    )
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
    cell_range = build_cell_range(
        COLUMNS['medium_security_risks'], 3,
        COLUMNS['high_security_risks'], 1000,
        sheet_id=sheet_id
    )
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


def set_data_validation_decay_dates_request(sheet_id):
    cell_range = build_cell_range(
        COLUMNS['licences_expire'], 3,
        COLUMNS['legislation_changes'], 1000,
        sheet_id=sheet_id
    )
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
    cell_range = build_cell_range(
        COLUMNS['preventing_degradation_over_time'], 3,
        COLUMNS['preventing_degradation_over_time'], 1000,
        sheet_id=sheet_id
    )
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
    cell_range = build_cell_range(
        COLUMNS['updated_on'], 3,
        COLUMNS['updated_on'], 1000,
        sheet_id=sheet_id
    )
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
        build_cell_range(
            COLUMNS['licences_expire'], 3,
            COLUMNS['legislation_changes'], 1000,
            sheet_id=sheet_id
        ),
        build_cell_range(
            COLUMNS['updated_on'], 3,
            COLUMNS['updated_on'], 1000,
            sheet_id=sheet_id
        ),
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
        build_cell_range(
            COLUMNS['people_summary'], 1,
            COLUMNS['people_summary'], 1000,
            sheet_id=sheet_id
        ),
        build_cell_range(
            COLUMNS['service_area'], 1,
            COLUMNS['service_area'], 1000,
            sheet_id=sheet_id
        ),
        build_cell_range(
            COLUMNS['civil_servants'], 1,
            COLUMNS['civil_servants'], 1000,
            sheet_id=sheet_id
        ),
        build_cell_range(
            COLUMNS['automated_tests'], 1,
            COLUMNS['automated_tests'], 1000,
            sheet_id=sheet_id
        ),
        build_cell_range(
            COLUMNS['licences_expire'], 1,
            COLUMNS['licences_expire'], 1000,
            sheet_id=sheet_id
        ),
        build_cell_range(
            COLUMNS['updated_by'], 1,
            COLUMNS['updated_by'], 1000,
            sheet_id=sheet_id
        ),
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


def set_default_green_background_tech_decay_summaries_request(sheet_id):
    """
    The tech and decay summary columns have a green background by default.

    Conditional formatting rules override this for red and amber.

    We need to pass values for every cell in the range we're updating, hence the
    two list comprehensions.
    """
    cell_range = build_cell_range(
        COLUMNS['tech_summary'], 3,
        COLUMNS['decay_summary'], 1000,
        sheet_id=sheet_id
    )
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
    cell_range = build_cell_range(
        COLUMNS['people_summary'], 3,
        COLUMNS['people_summary'], 1000,
        sheet_id=sheet_id
    )

    civil_servants = COLUMNS['civil_servants'] + '3'
    contractors = COLUMNS['contractors'] + '3'
    managed_service = COLUMNS['managed_service'] + '3'

    formula = f"=AND(({civil_servants}=0), ({contractors}=0), ({managed_service}=FALSE))"
    values = [{"userEnteredValue": formula}]

    return add_conditional_formatting_request(index, cell_range, values, red_background)


def add_conditional_formatting_people_amber_summary_request(sheet_id):
    index = 1
    cell_range = build_cell_range(
        COLUMNS['people_summary'], 3,
        COLUMNS['people_summary'], 1000,
        sheet_id=sheet_id
    )

    civil_servants = COLUMNS['civil_servants'] + '3'
    contractors = COLUMNS['contractors'] + '3'
    managed_service = COLUMNS['managed_service'] + '3'

    formula = "".join([
        "=OR(",
            f"AND(({civil_servants}=1), ({contractors}=0), ({managed_service}=FALSE)), ",
            f"AND(({civil_servants}=0), ({contractors}>=1), ({managed_service}=FALSE)), ",
            f"AND(({civil_servants}=0), ({contractors}=0), ({managed_service}=TRUE)), ",
            f"AND(({civil_servants}=0), ({contractors}>=1), ({managed_service}=TRUE)), ",
            f"AND(({civil_servants}=1), ({contractors}=1), ({managed_service}=FALSE))",
        ")"
    ])
    values = [{"userEnteredValue": formula}]

    return add_conditional_formatting_request(index, cell_range, values, amber_background)


def add_conditional_formatting_people_green_summary_request(sheet_id):
    index = 2
    cell_range = build_cell_range(
        COLUMNS['people_summary'], 3,
        COLUMNS['people_summary'], 1000,
        sheet_id=sheet_id
    )

    civil_servants = COLUMNS['civil_servants'] + '3'
    contractors = COLUMNS['contractors'] + '3'
    managed_service = COLUMNS['managed_service'] + '3'

    formula = "".join([
        "=OR(",
            f"({civil_servants}>=2), ",
            f"AND(({civil_servants}=1), {contractors}>=2), ",
            f"AND(({civil_servants}=1), ({managed_service}=TRUE))",
        ")"
    ])
    values = [{"userEnteredValue": formula}]

    return add_conditional_formatting_request(index, cell_range, values, green_background)


def add_conditional_formatting_tech_summary_red(sheet_id):
    requests = []

    cell_range = build_cell_range(
        COLUMNS['tech_summary'], 3,
        COLUMNS['tech_summary'], 1000,
        sheet_id=sheet_id
    )

    # Tech binary criteria - ease and risk of making changes
    index = 3
    first_cell = COLUMNS['automated_tests'] + '3'
    last_cell = COLUMNS['backup_recovery_strategy'] + '3'
    formula = f"=COUNTIF({first_cell}:{last_cell},\"FALSE\") >= 2"
    values = [{"userEnteredValue": formula}]
    requests.append(add_conditional_formatting_request(index, cell_range, values, red_background))

    # Understanding of security
    index = 4
    cell = COLUMNS['good_understanding_of_security'] + '3'
    formula = f"=({cell}=FALSE)"
    values = [{"userEnteredValue": formula}]
    requests.append(add_conditional_formatting_request(index, cell_range, values, red_background))

    # Number of medium risks
    index = 5
    cell = COLUMNS['medium_security_risks'] + '3'
    formula = f"={cell} > 5"
    values = [{"userEnteredValue": formula}]
    requests.append(add_conditional_formatting_request(index, cell_range, values, red_background))

    # Number of high risks
    index = 6
    cell = COLUMNS['high_security_risks'] + '3'
    formula = f"={cell} >= 1"
    values = [{"userEnteredValue": formula}]
    requests.append(add_conditional_formatting_request(index, cell_range, values, red_background))

    return requests


def add_conditional_formatting_tech_amber_summary_request(sheet_id):
    # Number of medium risks
    index = 7
    cell_range = build_cell_range(
        COLUMNS['tech_summary'], 3,
        COLUMNS['tech_summary'], 1000,
        sheet_id=sheet_id
    )
    cell = COLUMNS['medium_security_risks'] + '3'
    formula = f"=AND({cell} >= 2, {cell} <= 5)"
    values = [{"userEnteredValue": formula}]

    return add_conditional_formatting_request(index, cell_range, values, amber_background)


def add_conditional_formatting_decay_summary_red(sheet_id):
    requests = []

    cell_range = build_cell_range(
        COLUMNS['decay_summary'], 3,
        COLUMNS['decay_summary'], 1000,
        sheet_id=sheet_id
    )

    # Licences expire - red if date in past
    index = 8
    formula = date_in_past_condition(COLUMNS['licences_expire'])
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, red_background))

    # Dependencies go out of support - red if date in past
    index = 9
    formula = date_in_past_condition(COLUMNS['major_dependencies_out_of_support'])
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, red_background))

    # Oldest unapplied dependency version released - red if date more than a year ago
    index = 10
    formula = date_equal_or_earlier_than_condition(COLUMNS['oldest_unapplied_dependency_version_released'], days_in_future=-365)
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, red_background))

    # FIXME: index 11 removed herw

    # Support contract expire - red if date in past
    index = 12
    formula = date_in_past_condition(COLUMNS['support_contract_expires'])
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, red_background))

    # Relevant legislation changes come into effect - red if date less than
    # 3 months in the future (or already past)
    index = 13
    formula = date_equal_or_earlier_than_condition(COLUMNS['legislation_changes'], days_in_future=(3 * 30))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, red_background))

    return requests


def add_conditional_formatting_decay_summary_amber(sheet_id):
    requests = []

    cell_range = build_cell_range(
        COLUMNS['decay_summary'], 3,
        COLUMNS['decay_summary'], 1000,
        sheet_id=sheet_id
    )

    # Licences expire - amber if date less than 6 months in the future (or already past)
    index = 14
    formula = date_equal_or_earlier_than_condition(COLUMNS['licences_expire'], days_in_future=(6 * 30))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, amber_background))

    # Dependencies go out of support - amber if date less than 6 months in the future (or already past)
    index = 15
    formula = date_equal_or_earlier_than_condition(COLUMNS['major_dependencies_out_of_support'], days_in_future=(6 * 30))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, amber_background))

    # Oldest unapplied dependency version released - amber if date more than 3 months ago
    index = 16
    formula = date_equal_or_earlier_than_condition(COLUMNS['oldest_unapplied_dependency_version_released'], days_in_future=(3 * 30 * -1))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, amber_background))

    # FIXME: index 17 removed here

    # Support contract expire - amber if date less than 9 months in the future (or already past)
    index = 18
    formula = date_equal_or_earlier_than_condition(COLUMNS['support_contract_expires'], days_in_future=(9 * 30))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, amber_background))

    # Relevant legislation changes come into effect - amber if date less than 6 months in the future (or already past)
    index = 19
    formula = date_equal_or_earlier_than_condition(COLUMNS['legislation_changes'], days_in_future=(6 * 30))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, amber_background))

    # Actively preventing degradation over time? - amber if false
    index = 20
    cell = COLUMNS['preventing_degradation_over_time'] + '3'
    formula = f"=({cell}=FALSE)"
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, amber_background))

    return requests


def add_conditional_formatting_tech_individual_criteria_red(sheet_id):
    requests = []

    # Binary tech criteria - red if false
    index = 21
    cell_range = build_cell_range(
        COLUMNS['automated_tests'], 3,
        COLUMNS['good_understanding_of_security'], 1000,
        sheet_id=sheet_id
    )
    cell = COLUMNS['automated_tests'] + '3'
    formula = f"=({cell}=FALSE)"
    values = [{"userEnteredValue": formula}]
    requests.append(add_conditional_formatting_request(index, cell_range, values, light_red_background))

    # Number of medium risks - red if more than 5
    index = 22
    cell_range = build_cell_range(
        COLUMNS['medium_security_risks'], 3,
        COLUMNS['medium_security_risks'], 1000,
        sheet_id=sheet_id
    )
    cell = COLUMNS['medium_security_risks'] + '3'
    formula = f"={cell} > 5"
    values = [{"userEnteredValue": formula}]
    requests.append(add_conditional_formatting_request(index, cell_range, values, light_red_background))

    # Number of high risks - red if one or more
    index = 23
    cell_range = build_cell_range(
        COLUMNS['high_security_risks'], 3,
        COLUMNS['high_security_risks'], 1000,
        sheet_id=sheet_id
    )
    cell = COLUMNS['high_security_risks'] + '3'
    formula = f"={cell} >= 1"
    values = [{"userEnteredValue": formula}]
    requests.append(add_conditional_formatting_request(index, cell_range, values, light_red_background))

    return requests


def add_conditional_formatting_tech_individual_criteria_amber_request(sheet_id):
    # Number of medium risks - amber if between 2 and 5 inclusive
    index = 24
    cell_range = build_cell_range(
        COLUMNS['medium_security_risks'], 3,
        COLUMNS['medium_security_risks'], 1000,
        sheet_id=sheet_id
    )
    cell = COLUMNS['medium_security_risks'] + '3'
    formula = f"=AND({cell} >= 2, {cell} <= 5)"
    values = [{"userEnteredValue": formula}]

    return add_conditional_formatting_request(index, cell_range, values, light_amber_background)


def add_conditional_formatting_tech_individual_criteria_green(sheet_id):
    requests = []

    # Binary tech criteria - green if true
    index = 25
    cell_range = build_cell_range(
        COLUMNS['automated_tests'], 3,
        COLUMNS['good_understanding_of_security'], 1000,
        sheet_id=sheet_id
    )
    cell = COLUMNS['automated_tests'] + '3'
    formula = f"=({cell}=TRUE)"
    values = [{"userEnteredValue": formula}]
    requests.append(add_conditional_formatting_request(index, cell_range, values, light_green_background))

    # Number of medium risks - green if less than 2 and not blank
    index = 26
    cell_range = build_cell_range(
        COLUMNS['medium_security_risks'], 3,
        COLUMNS['medium_security_risks'], 1000,
        sheet_id=sheet_id
    )
    cell = COLUMNS['medium_security_risks'] + '3'
    formula = f"=AND(NOT(ISBLANK({cell})), ({cell} < 2))"
    values = [{"userEnteredValue": formula}]
    requests.append(add_conditional_formatting_request(index, cell_range, values, light_green_background))

    # Number of high risks - green if 0 and not blank
    index = 27
    cell_range = build_cell_range(
        COLUMNS['high_security_risks'], 3,
        COLUMNS['high_security_risks'], 1000,
        sheet_id=sheet_id
    )
    cell = COLUMNS['high_security_risks'] + '3'
    formula = f"=AND(NOT(ISBLANK({cell})), ({cell}=0))"
    values = [{"userEnteredValue": formula}]
    requests.append(add_conditional_formatting_request(index, cell_range, values, light_green_background))

    return requests


def add_conditional_formatting_decay_individual_criteria_red(sheet_id):
    requests = []

    # Licences expire - red if date in past
    index = 28
    cell_range = build_cell_range(
        COLUMNS['licences_expire'], 3,
        COLUMNS['licences_expire'], 1000,
        sheet_id=sheet_id
    )

    formula = date_in_past_condition(COLUMNS['licences_expire'])
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_red_background))

    # Dependencies go out of support - red if date in past
    index = 29
    cell_range = build_cell_range(
        COLUMNS['major_dependencies_out_of_support'], 3,
        COLUMNS['major_dependencies_out_of_support'], 1000,
        sheet_id=sheet_id
    )

    formula = date_in_past_condition(COLUMNS['major_dependencies_out_of_support'])
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_red_background))

    # Oldest unapplied dependency version released - red if date more than a year ago
    index = 30
    cell_range = build_cell_range(
        COLUMNS['oldest_unapplied_dependency_version_released'], 3,
        COLUMNS['oldest_unapplied_dependency_version_released'], 1000,
        sheet_id=sheet_id
    )

    formula = date_equal_or_earlier_than_condition(COLUMNS['oldest_unapplied_dependency_version_released'], days_in_future=-365)
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_red_background))

    # FIXME: index 31 removed here

    # Support contract expire - red if date in past
    index = 32
    cell_range = build_cell_range(
        COLUMNS['support_contract_expires'], 3,
        COLUMNS['support_contract_expires'], 1000,
        sheet_id=sheet_id
    )

    formula = date_in_past_condition(COLUMNS['support_contract_expires'])
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_red_background))

    # Relevant legislation changes come into effect - red if date less than
    # 3 months in the future (or already past)
    index = 33
    cell_range = build_cell_range(
        COLUMNS['legislation_changes'], 3,
        COLUMNS['legislation_changes'], 1000,
        sheet_id=sheet_id
    )

    formula = date_equal_or_earlier_than_condition(COLUMNS['legislation_changes'], days_in_future=(3 * 30))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_red_background))

    return requests


def add_conditional_formatting_decay_individual_criteria_amber(sheet_id):
    requests = []

    # Licences expire - amber if date less than 6 months in the future (or already past)
    index = 34
    cell_range = build_cell_range(
        COLUMNS['licences_expire'], 3,
        COLUMNS['licences_expire'], 1000,
        sheet_id=sheet_id
    )

    formula = date_equal_or_earlier_than_condition(COLUMNS['licences_expire'], days_in_future=(6 * 30))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_amber_background))

    # Dependencies go out of support - amber if date less than 6 months in the future (or already past)
    index = 35
    cell_range = build_cell_range(
        COLUMNS['major_dependencies_out_of_support'], 3,
        COLUMNS['major_dependencies_out_of_support'], 1000,
        sheet_id=sheet_id
    )

    formula = date_equal_or_earlier_than_condition(COLUMNS['major_dependencies_out_of_support'], days_in_future=(6 * 30))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_amber_background))

    # Oldest unapplied dependency version released - amber if date more than 3 months ago
    index = 36
    cell_range = build_cell_range(
        COLUMNS['oldest_unapplied_dependency_version_released'], 3,
        COLUMNS['oldest_unapplied_dependency_version_released'], 1000,
        sheet_id=sheet_id
    )

    formula = date_equal_or_earlier_than_condition(COLUMNS['oldest_unapplied_dependency_version_released'], days_in_future=(3 * 30 * -1))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_amber_background))

    # FIXME: index 37 removed here

    # Support contract expire - amber if date less than 9 months in the future (or already past)
    index = 38
    cell_range = build_cell_range(
        COLUMNS['support_contract_expires'], 3,
        COLUMNS['support_contract_expires'], 1000,
        sheet_id=sheet_id
    )

    formula = date_equal_or_earlier_than_condition(COLUMNS['support_contract_expires'], days_in_future=(9 * 30))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_amber_background))

    # Relevant legislation changes come into effect - amber if date less than 6 months in the future (or already past)
    index = 39
    cell_range = build_cell_range(
        COLUMNS['legislation_changes'], 3,
        COLUMNS['legislation_changes'], 1000,
        sheet_id=sheet_id
    )

    formula = date_equal_or_earlier_than_condition(COLUMNS['legislation_changes'], days_in_future=(6 * 30))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_amber_background))

    # Actively preventing degradation over time? - amber if false
    index = 40
    cell_range = build_cell_range(
        COLUMNS['preventing_degradation_over_time'], 3,
        COLUMNS['preventing_degradation_over_time'], 1000,
        sheet_id=sheet_id
    )
    cell = COLUMNS['preventing_degradation_over_time'] + '3'
    formula = f"=({cell}=FALSE)"
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_amber_background))

    return requests


def add_conditional_formatting_decay_individual_criteria_green(sheet_id):
    requests = []

    # Licences expire - green if date more than 6 months in the future
    index = 41
    cell_range = build_cell_range(
        COLUMNS['licences_expire'], 3,
        COLUMNS['licences_expire'], 1000,
        sheet_id=sheet_id
    )

    formula = date_later_than_condition(COLUMNS['licences_expire'], days_in_future=(6 * 30))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_green_background))

    # Dependencies go out of support - green if date more than 6 months in the future
    index = 42
    cell_range = build_cell_range(
        COLUMNS['major_dependencies_out_of_support'], 3,
        COLUMNS['major_dependencies_out_of_support'], 1000,
        sheet_id=sheet_id
    )

    formula = date_later_than_condition(COLUMNS['major_dependencies_out_of_support'], days_in_future=(6 * 30))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_green_background))

    # Oldest unapplied dependency version released - green if date less than 3 months ago
    index = 43
    cell_range = build_cell_range(
        COLUMNS['oldest_unapplied_dependency_version_released'], 3,
        COLUMNS['oldest_unapplied_dependency_version_released'], 1000,
        sheet_id=sheet_id
    )

    formula = date_later_than_condition(COLUMNS['oldest_unapplied_dependency_version_released'], days_in_future=(3 * 30 * -1))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_green_background))

    # FIXME: index 44 removed here

    # Support contract expire - green if date more than 9 months in the future
    index = 45
    cell_range = build_cell_range(
        COLUMNS['support_contract_expires'], 3,
        COLUMNS['support_contract_expires'], 1000,
        sheet_id=sheet_id
    )

    formula = date_later_than_condition(COLUMNS['support_contract_expires'], days_in_future=(9 * 30))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_green_background))

    # Relevant legislation changes come into effect - green if date more than 6 months in the future
    index = 46
    cell_range = build_cell_range(
        COLUMNS['legislation_changes'], 3,
        COLUMNS['legislation_changes'], 1000,
        sheet_id=sheet_id
    )

    formula = date_later_than_condition(COLUMNS['legislation_changes'], days_in_future=(6 * 30))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_green_background))

    # Actively preventing degradation over time? - green if true
    index = 47
    cell_range = build_cell_range(
        COLUMNS['preventing_degradation_over_time'], 3,
        COLUMNS['preventing_degradation_over_time'], 1000,
        sheet_id=sheet_id
    )
    cell = COLUMNS['preventing_degradation_over_time'] + '3'
    formula = f"=({cell}=TRUE)"
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_green_background))

    # All date-based decay criteria - green if not applicable (entered as "N/A")
    index = 48
    cell_range = build_cell_range(
        COLUMNS['licences_expire'], 3,
        COLUMNS['legislation_changes'], 1000,
        sheet_id=sheet_id
    )
    cell = COLUMNS['licences_expire'] + '3'
    formula = f"=({cell}=\"N/A\")"
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_green_background))

    return requests


def add_conditional_formatting_update_details_red(sheet_id):
    requests = []

    # Both update details columns - red if blank
    index = 49
    cell_range = build_cell_range(
        COLUMNS['updated_by'], 3,
        COLUMNS['updated_on'], 1000,
        sheet_id=sheet_id
    )
    cell = COLUMNS['updated_by'] + '3'
    formula = f"=ISBLANK({cell})"
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_red_background))

    # Updated on - red if in the future
    index = 50
    cell_range = build_cell_range(
        COLUMNS['updated_on'], 3,
        COLUMNS['updated_on'], 1000,
        sheet_id=sheet_id
    )

    formula = date_later_than_condition(COLUMNS['updated_on'], days_in_future=(0))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_red_background))

    # Updated on - red if more than 3 months ago
    index = 51
    cell_range = build_cell_range(
        COLUMNS['updated_on'], 3,
        COLUMNS['updated_on'], 1000,
        sheet_id=sheet_id
    )

    formula = date_equal_or_earlier_than_condition(COLUMNS['updated_on'], days_in_future=(3 * 30 * -1))
    values = [{"userEnteredValue": formula}]

    requests.append(add_conditional_formatting_request(index, cell_range, values, light_red_background))

    return requests


def add_conditional_formatting_update_details_amber_request(sheet_id):
    # Updated on - amber if more than 2 months ago
    index = 52
    cell_range = build_cell_range(
        COLUMNS['updated_on'], 3,
        COLUMNS['updated_on'], 1000,
        sheet_id=sheet_id
    )

    formula = date_equal_or_earlier_than_condition(COLUMNS['updated_on'], days_in_future=(2 * 30 * -1))
    values = [{"userEnteredValue": formula}]

    return add_conditional_formatting_request(index, cell_range, values, light_amber_background)


def add_conditional_formatting_update_details_green_request(sheet_id):
    # Updated on - green if less than 2 months ago
    index = 53
    cell_range = build_cell_range(
        COLUMNS['updated_on'], 3,
        COLUMNS['updated_on'], 1000,
        sheet_id=sheet_id
    )

    formula = date_later_than_condition(COLUMNS['updated_on'], days_in_future=(2 * 30 * -1))
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
        build_cell_range(
            COLUMNS['service'], 1,
            COLUMNS['updated_on'], 2,
            sheet_id=sheet_id
        ),
        build_cell_range(
            COLUMNS['service'], 3,
            COLUMNS['service'], 1000,
            sheet_id=sheet_id
        ),
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
    column_names_with_widths = [
        ('service', 230),
        ('service_area', 160),
        ('notes', 200),
        ('civil_servants', 95),
        ('contractors', 85),
        ('managed_service', 65),
        ('updated_by', 120),
    ]

    # This API call wants the range in a different format
    def format_range_with_dimension(column_name, sheet_id):
        full_range = build_cell_range(COLUMNS[column_name], 1, COLUMNS[column_name], 1000, sheet_id=sheet_id)
        return {
            "sheetId": sheet_id,
            "dimension": "COLUMNS",
            "startIndex": full_range["startColumnIndex"],
            "endIndex": full_range["endColumnIndex"]
        }

    requests = [
        {
            "updateDimensionProperties": {
                "range": format_range_with_dimension(column_name, sheet_id),
                "properties": {
                    "pixelSize": width
                },
                "fields": "pixelSize"
            }
        } for column_name, width in column_names_with_widths
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
    requests.append(protect_header_rows_request(sheet_id, PROTECTED_RANGE_EDITOR_USERS, PROTECTED_RANGE_EDITOR_GROUPS))
    requests.extend(merge_header_row_cells(sheet_id))
    requests.extend(write_first_header_rows(sheet_id))
    requests.append(write_second_header_row_request(sheet_id))
    requests.append(set_data_validation_number_of_people_request(sheet_id))
    requests.append(set_data_validation_managed_service_request(sheet_id))
    requests.append(set_data_validation_tech_boolean_criteria_request(sheet_id))
    requests.append(set_data_validation_number_of_security_risks_request(sheet_id))
    requests.append(set_data_validation_decay_dates_request(sheet_id))
    requests.append(set_data_validation_preventing_degradation_request(sheet_id))
    requests.append(set_data_validation_updated_on_date_request(sheet_id))
    requests.extend(format_dates(sheet_id))
    requests.extend(set_borders(sheet_id))
    requests.append(set_default_green_background_tech_decay_summaries_request(sheet_id))
    requests.append(add_conditional_formatting_people_red_summary_request(sheet_id))
    requests.append(add_conditional_formatting_people_amber_summary_request(sheet_id))
    requests.append(add_conditional_formatting_people_green_summary_request(sheet_id))
    requests.extend(add_conditional_formatting_tech_summary_red(sheet_id))
    requests.append(add_conditional_formatting_tech_amber_summary_request(sheet_id))
    requests.extend(add_conditional_formatting_decay_summary_red(sheet_id))
    requests.extend(add_conditional_formatting_decay_summary_amber(sheet_id))
    requests.extend(add_conditional_formatting_tech_individual_criteria_red(sheet_id))
    requests.append(add_conditional_formatting_tech_individual_criteria_amber_request(sheet_id))
    requests.extend(add_conditional_formatting_tech_individual_criteria_green(sheet_id))
    requests.extend(add_conditional_formatting_decay_individual_criteria_red(sheet_id))
    requests.extend(add_conditional_formatting_decay_individual_criteria_amber(sheet_id))
    requests.extend(add_conditional_formatting_decay_individual_criteria_green(sheet_id))
    requests.extend(add_conditional_formatting_update_details_red(sheet_id))
    requests.append(add_conditional_formatting_update_details_amber_request(sheet_id))
    requests.append(add_conditional_formatting_update_details_green_request(sheet_id))
    requests.append(freeze_header_rows_and_summary_columns_request(sheet_id))
    requests.extend(bold_and_wrap_text_in_header_rows_and_service_column(sheet_id))
    requests.extend(set_column_widths(sheet_id))

    return requests
