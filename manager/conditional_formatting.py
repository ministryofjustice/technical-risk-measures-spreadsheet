def date_in_past_condition(column):
    """
    Field is not empty and date is in the past.
    """
    return f"=AND(NOT(ISBLANK({column}3)), {column}3<TODAY())"
