import pytest

from manager.conditional_formatting import date_in_past_condition, date_comparison_condition
from manager.helpers import column_letters_to_number, split_cell_string, a1_to_range


def test_column_letters_to_number_A():
    assert column_letters_to_number('A') == 0


def test_column_letters_to_number_B():
    assert column_letters_to_number('B') == 1


def test_column_letters_to_number_Z():
    assert column_letters_to_number('Z') == 25


def test_column_letters_to_number_AC():
    assert column_letters_to_number('AC') == 28


def test_split_cell_string_A1():
    assert split_cell_string('A1') == ('A', '1')


def test_split_cell_string_KLR947():
    assert split_cell_string('KLR947') == ('KLR', '947')


def test_split_cell_string_case_insensitive():
    assert split_cell_string('klr947') == ('KLR', '947')


def test_split_cell_string_invalid_cell():
    # This is an invalid cell specifier - this test is here to demonstrate that
    # split_cell_string isn't doing anything clever, so its input must be good.
    assert split_cell_string('94J4NM20X') == ('JNMX', '94420')


def test_a1_to_range_A1():
    expected_range = {
        "sheetId": 0,
        "startColumnIndex": 0,
        "startRowIndex": 0,
        "endColumnIndex": 1,
        "endRowIndex": 1
    }
    assert a1_to_range('A1', 0) == expected_range


def test_a1_to_range_A12():
    expected_range = {
        "sheetId": 0,
        "startColumnIndex": 0,
        "startRowIndex": 11,
        "endColumnIndex": 1,
        "endRowIndex": 12
    }
    assert a1_to_range('A12', 0) == expected_range


def test_a1_to_range_non_alphanumeric():
    expected_range = {
        "sheetId": 0,
        "startColumnIndex": 0,
        "startRowIndex": 11,
        "endColumnIndex": 1,
        "endRowIndex": 12
    }
    with pytest.raises(ValueError):
        a1_to_range('A&3', 0)


def test_a1_to_range_C2E11():
    expected_range = {
        "sheetId": 5,
        "startColumnIndex": 2,
        "startRowIndex": 1,
        "endColumnIndex": 5,
        "endRowIndex": 11
    }
    assert a1_to_range('C2:E11', 5) == expected_range


def test_a1_to_range_A2AC2():
    expected_range = {
        "sheetId": 123,
        "startColumnIndex": 0,
        "startRowIndex": 1,
        "endColumnIndex": 29,
        "endRowIndex": 2
    }
    assert a1_to_range('A2:AC2', 123) == expected_range


def test_date_in_past_condition():
    expected_formula = '=AND(NOT(ISBLANK(CD3)), CD3<TODAY())'
    assert date_in_past_condition('CD') == expected_formula


def test_date_comparison_condition():
    expected_formula = "=AND(NOT(ISBLANK(X3)), IF(X3<TODAY(),DATEDIF(X3,TODAY(),\"D\")*-1,DATEDIF(TODAY(), X3, \"D\")) <= -365)"
    assert date_comparison_condition('X', '<=', -365)
