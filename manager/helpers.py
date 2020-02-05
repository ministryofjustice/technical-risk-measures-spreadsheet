def column_letters_to_number(letters):
    """
    Convert a column in letters to a zero-indexed number.
    """
    total = 0
    remaining_letters = letters
    while len(remaining_letters) > 0:
        current_letter, *remaining_letters = remaining_letters
        total = (ord(current_letter) - 64) + (total * 26)
    return total - 1


def split_cell_string(cell):
    """
    Split a cell specifier into its letters and numbers.

    This uses letters from the whole string, rather than only from the start,
    so cells must be specified correctly.

    This function also takes care of case-insensitivity - upper and lower case
    letters refer to the same column in spreadsheets so we should do the same
    here.
    """
    letters = ''
    numbers = ''
    for char in cell:
        if char.isalpha():
            letters += char
        elif char.isdigit():
            numbers += char
        else:
            raise ValueError('Found a non-alphanumeric character in cell specifier')
    return (letters.upper(), numbers)


def a1_to_range(cells, sheet_id):
    """
    Convert a range of cells from A1 notation to a range object.

    A1 notation is more intuitive to work with, but the Sheets API requires
    range objects.
    """
    if ':' in cells:
        start, end = cells.split(sep=':', maxsplit=1)
    else:
        start = end = cells

    start_letters, start_numbers = split_cell_string(start)
    start_column_index = column_letters_to_number(start_letters)
    start_row_index = int(start_numbers) - 1

    end_letters, end_numbers = split_cell_string(end)
    end_column_index = column_letters_to_number(end_letters) + 1
    end_row_index = int(end_numbers)

    return {
        "sheetId": sheet_id,
        "startColumnIndex": start_column_index,
        "startRowIndex": start_row_index,
        "endColumnIndex": end_column_index,
        "endRowIndex": end_row_index
    }


def build_cell_range(first_column, first_row, last_column=None, last_row=None, sheet_id=None):
    """
    Construct a range object from the given arguments.

    Specify the values of the column and row arguments as in A1 notation - ie
    using the values displayed in the spreadsheet for them (not zero-indexed).

    The resulting range is inclusive of the first and last rows & columns in
    the arguments (as in A1 notation, but unlike the range object format).

    last_column and last_row are optional arguments; if either is not provided,
    the resulting range will represent the single cell of the first_column and
    first_row.

    Calling this function looks more cumbersome than A1 notation, but is clearer
    when using columns from a reference dict (rather than hardcoding letters
    throughout).
    """
    # I think it's clearer to have sheet_id as the last argument, but this
    # means it needs to have a default value since it comes after the two
    # optional "last" arguments:
    if sheet_id is None:
        raise TypeError('sheet_id is required')

    if last_column is None or last_row is None:
        last_column = first_column
        last_row = first_row

    start_column_index = column_letters_to_number(first_column)
    start_row_index = int(first_row) - 1

    end_column_index = column_letters_to_number(last_column) + 1
    end_row_index = int(last_row)

    return {
        "sheetId": sheet_id,
        "startColumnIndex": start_column_index,
        "startRowIndex": start_row_index,
        "endColumnIndex": end_column_index,
        "endRowIndex": end_row_index
    }


red_background = {
    "backgroundColor": {
        "blue": 0,
        "green": 0,
        "red": 1,
    }
}


amber_background = {
    "backgroundColor": {
        "blue": 0,
        "green": 0.6,
        "red": 1,
    }
}


green_background = {
    "backgroundColor": {
        "blue": 0,
        "green": 1,
        "red": 0,
    }
}


light_red_background = {
    "backgroundColor": {
        "blue": 0.8,
        "green": 0.8,
        "red": 0.957,
    }
}


light_amber_background = {
    "backgroundColor": {
        "blue": 0.804,
        "green": 0.899,
        "red": 0.989,
    }
}


light_green_background = {
    "backgroundColor": {
        "blue": 0.828,
        "green": 0.918,
        "red": 0.853,
    }
}


def add_conditional_formatting_request(index, cell_range, values, format):
    """
    Build a request to add a conditional formatting rule using a custom formula.
    """
    return {
        "addConditionalFormatRule": {
            "index": index,
            "rule": {
                "ranges": [cell_range],
                "booleanRule": {
                    "condition": {
                        "values": values,
                        "type": "CUSTOM_FORMULA",
                    },
                    "format": format
                }
            }
        }
    }
