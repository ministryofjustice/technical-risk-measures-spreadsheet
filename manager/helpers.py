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


def add_conditional_formatting_request(index, range, values, format):
    """
    Build a request to add a conditional formatting rule using a custom formula.
    """
    return {
        "addConditionalFormatRule": {
            "index": index,
            "rule": {
                "ranges": [range],
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
