def column_letter_to_number(letter):
    """
    Convert a single column letter to a zero-indexed number.
    """
    # Convert to ASCII and reduce A to 0:
    return ord(letter) - 65


def split_cell_string(cell):
    """
    Split a cell specifier into its letters and numbers.

    This uses letters from the whole string, rather than only from the start,
    so cells must be specified correctly.
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
    return (letters, numbers)


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
    start_column_index = column_letter_to_number(start_letters.upper())
    start_row_index = int(start_numbers) - 1

    end_letters, end_numbers = split_cell_string(end)
    end_column_index = column_letter_to_number(end_letters.upper()) + 1
    end_row_index = int(end_numbers)

    return {
        "sheetId": sheet_id,
        "startColumnIndex": start_column_index,
        "startRowIndex": start_row_index,
        "endColumnIndex": end_column_index,
        "endRowIndex": end_row_index
    }
