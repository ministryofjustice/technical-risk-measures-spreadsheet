def column_letter_to_number(letter):
    """
    Convert a single column letter to a zero-indexed number.
    """
    # Convert to ASCII and reduce A to 0:
    return ord(letter) - 65
