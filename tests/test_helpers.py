from manager.helpers import column_letter_to_number, split_cell_string


def test_column_letter_to_number_A():
    assert column_letter_to_number('A') == 0


def test_column_letter_to_number_B():
    assert column_letter_to_number('B') == 1


def test_split_cell_string_A1():
    assert split_cell_string('A1') == ('A', '1')


def test_split_cell_string_KLR947():
    assert split_cell_string('KLR947') == ('KLR', '947')


def test_split_cell_string_invalid_cell():
    # This is an invalid cell specifier - this test is here to demonstrate that
    # split_cell_string isn't doing anything clever, so its input must be good.
    assert split_cell_string('94J4NM20X') == ('JNMX', '94420')
