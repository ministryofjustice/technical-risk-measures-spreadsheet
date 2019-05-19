def column_letter_to_number(letter):
    """
    Convert a single column letter to a zero-indexed number.
    """
    # Convert to ASCII and reduce A to 0:
    return ord(letter) - 65


def split_cell_string(cell):
    """
    Split a cell specifier into letters and numbers.

    This doesn't do anything clever, so cells need to be specified correctly.
    """
    letters = ''
    numbers = ''
    for char in cell:
        if char.isalpha():
            letters += char
        else:
            numbers += char
    return (letters, numbers)
