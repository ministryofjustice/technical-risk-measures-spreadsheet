from manager.helpers import column_letter_to_number


def test_column_letter_to_number_A():
    assert column_letter_to_number('A') == 0


def test_column_letter_to_number_B():
    assert column_letter_to_number('B') == 1
