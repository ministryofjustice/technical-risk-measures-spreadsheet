import pytest

from manager.conditional_formatting import date_in_past_condition, \
    date_equal_or_earlier_than_condition
from manager.helpers import column_letters_to_number, build_cell_range


def test_column_letters_to_number_A():
    assert column_letters_to_number('A') == 0


def test_column_letters_to_number_B():
    assert column_letters_to_number('B') == 1


def test_column_letters_to_number_Z():
    assert column_letters_to_number('Z') == 25


def test_column_letters_to_number_AC():
    assert column_letters_to_number('AC') == 28


def test_build_cell_range_with_last_values():
    expected_range = {
        "sheetId": 5,
        "startColumnIndex": 2,
        "startRowIndex": 1,
        "endColumnIndex": 5,
        "endRowIndex": 11
    }
    assert build_cell_range('C', 2, 'E', 11, sheet_id=5) == expected_range


def test_build_cell_range_without_last_values_for_single_cell():
    expected_range = {
        "sheetId": 5,
        "startColumnIndex": 2,
        "startRowIndex": 1,
        "endColumnIndex": 3,
        "endRowIndex": 2
    }
    assert build_cell_range('C', 2, sheet_id=5) == expected_range


def test_build_cell_range_missing_sheet_id():
    with pytest.raises(TypeError) as excinfo:
        build_cell_range('C', 2, 'D', 6)
    assert 'sheet_id is required' in str(excinfo.value)


def test_date_in_past_condition():
    expected_formula = '=AND(NOT(ISBLANK(CD3)), CD3<TODAY())'
    assert date_in_past_condition('CD') == expected_formula


def test_date_equal_or_earlier_than_condition():
    expected_formula = "=AND(NOT(ISBLANK(X3)), IF(X3<TODAY(),DATEDIF(X3,TODAY(),\"D\")*-1,DATEDIF(TODAY(), X3, \"D\")) <= -365)"
    assert date_equal_or_earlier_than_condition('X', days_in_future=-365)
