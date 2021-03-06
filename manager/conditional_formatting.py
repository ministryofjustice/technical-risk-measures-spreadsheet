def date_in_past_condition(column):
    """
    Returns true if field is not empty and date is in the past.

    This condition can also be expressed using date_comparison_condition, but
    this version is more readable in the spreadsheet. Not sure if it's worth
    keeping this one for that reason, though.
    """
    return f"=AND(NOT(ISBLANK({column}3)), {column}3<TODAY())"


def date_equal_or_earlier_than_condition(column, days_in_future):
    """
    Returns true if field is not empty, and date is equal to or earlier than days_in_future.

    column must be a string.
    days_in_future can be a string or an integer.

    Make days_in_future negative to represent a date in the past.
    This handles the date in the cell being either past or future in any case.
    """
    return f"=AND(NOT(ISBLANK({column}3)), IF({column}3<TODAY(),DATEDIF({column}3,TODAY(),\"D\")*-1,DATEDIF(TODAY(), {column}3, \"D\")) <= {days_in_future})"


def date_later_than_condition(column, days_in_future):
    """
    Returns true if field is not empty, and date is later than days_in_future.

    column must be a string.
    days_in_future can be a string or an integer.

    Make days_in_future negative to represent a date in the past.
    This handles the date in the cell being either past or future in any case.
    """
    return f"=AND(NOT(ISBLANK({column}3)), IF({column}3<TODAY(),DATEDIF({column}3,TODAY(),\"D\")*-1,DATEDIF(TODAY(), {column}3, \"D\")) > {days_in_future})"
