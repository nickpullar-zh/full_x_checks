from strategies.x_checks.variable_builder import (
    build_variables_string,
    _build_variable_name,
    _build_accounts_string,
)


# ---------------------------------------------------------------------------
# _build_variable_name
# ---------------------------------------------------------------------------

def test_name_empty_accounts_returns_blank():
    # REGRESSION: empty list crashed before guard was added; now returns '<blank>'
    assert _build_variable_name([], []) == '<blank>'


def test_name_single_account_no_types():
    assert _build_variable_name(['A246'], []) == 'A246'


def test_name_multiple_accounts_adds_ff():
    assert _build_variable_name(['A246', 'A247'], []) == 'A246ff'


def test_name_multiple_accounts_uses_sorted_first():
    # Sorted: A246 comes before A247 — name starts with A246
    assert _build_variable_name(['A247', 'A246'], []) == 'A246ff'


def test_name_with_single_movement_type():
    assert _build_variable_name(['A246'], ['1']) == 'A246ToM1'


def test_name_multiple_types_adds_ff():
    assert _build_variable_name(['A246'], ['1', '2']) == 'A246ToM1ff'


def test_name_strips_trailing_dot_zero():
    # REGRESSION: old code used .replace('.0', '') which hit any '.0' substring
    # New code uses endswith('.0') slice — only strips genuine trailing .0
    assert _build_variable_name(['A246.0'], []) == 'A246'


def test_name_dot_zero_in_middle_preserved():
    # .0 mid-string must NOT be stripped
    assert _build_variable_name(['A2.046'], []) == 'A2.046'


def test_name_filters_nan_string_in_types():
    # REGRESSION: 'nan' string from pandas NaN must not appear in variable name
    assert _build_variable_name(['A246'], ['nan']) == 'A246'


def test_name_filters_empty_string_in_types():
    assert _build_variable_name(['A246'], ['']) == 'A246'


# ---------------------------------------------------------------------------
# _build_accounts_string
# ---------------------------------------------------------------------------

def test_accounts_string_single():
    assert _build_accounts_string(['A246']) == 'A246'


def test_accounts_string_sorted_caret_delimited():
    assert _build_accounts_string(['B', 'A']) == 'A^B'


def test_accounts_string_empty_returns_blank():
    assert _build_accounts_string([]) == '<blank>'


def test_accounts_string_filters_nan():
    assert _build_accounts_string(['A246', 'nan']) == 'A246'


# ---------------------------------------------------------------------------
# build_variables_string
# ---------------------------------------------------------------------------

def test_build_variables_string_single():
    raw = [{'fs_accounts': ['A246'], 'movement_types': ['1']}]
    result = build_variables_string(raw)
    assert result == 'Name:A246ToM1;FS Account:A246;Movement Types:1'


def test_build_variables_string_multiple():
    raw = [
        {'fs_accounts': ['A246'], 'movement_types': []},
        {'fs_accounts': ['A300'], 'movement_types': ['2']},
    ]
    result = build_variables_string(raw)
    parts = result.split('|')
    assert len(parts) == 2
    assert 'Name:A246' in parts[0]
    assert 'Name:A300ToM2' in parts[1]


def test_build_variables_string_empty_list():
    assert build_variables_string([]) == ''
