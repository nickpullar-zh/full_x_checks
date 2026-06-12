from strategies.x_checks.fip_extraction import (
    _safe_split,
    _canonical_var_name,
    _build_excl_formula,
    _get_x_check_information,
    _VAR_HEADER,
    _BLANK_LINE,
    _BLOCK_END,
    _FORMULA_HEADER,
    _SEPARATOR,
)


# ---------------------------------------------------------------------------
# _safe_split
# ---------------------------------------------------------------------------

def test_safe_split_valid_index():
    assert _safe_split('a b c', 1) == 'b'


def test_safe_split_first_token():
    assert _safe_split('hello world', 0) == 'hello'


def test_safe_split_out_of_bounds_returns_default():
    # REGRESSION: bare line.split()[5] raised IndexError on short lines
    assert _safe_split('a b', 5) == ''


def test_safe_split_custom_default():
    assert _safe_split('a', 5, 'MISSING') == 'MISSING'


def test_safe_split_empty_line():
    assert _safe_split('', 0) == ''


def test_safe_split_extra_whitespace():
    assert _safe_split('  a   b  ', 1) == 'b'


# ---------------------------------------------------------------------------
# _canonical_var_name — converts any excl notation to excl.acc.type=N
# ---------------------------------------------------------------------------

def test_canonical_var_name_single_type():
    # excl.acc.type2 (no = sign) → canonical
    assert _canonical_var_name('LIN_00380excl.acc.type2', ['2']) == \
        'LIN_00380excl.acc.type=2'


def test_canonical_var_name_multi_type():
    # multi-type: types sorted numerically, comma-joined
    assert _canonical_var_name('LIN_00380excl.acc.type:1,4', ['1', '4']) == \
        'LIN_00380excl.acc.type=1,4'


def test_canonical_var_name_aff_notation():
    # excl.2-Aff (previously unhandled gap) → canonical
    assert _canonical_var_name('SLSTLSAS_11520n-lifeLOBsexcl.2-Aff', ['2']) == \
        'SLSTLSAS_11520n-lifeLOBsexcl.acc.type=2'


def test_canonical_var_name_tom_preserved():
    # TOM movement-type suffix after excl must be preserved
    assert _canonical_var_name('LIN_00380excl.acc.type2TOML09', ['2']) == \
        'LIN_00380excl.acc.type=2TOML09'


def test_canonical_var_name_tom_camelcase_preserved():
    assert _canonical_var_name('LIN_00380excl.acc.type=2ToML07', ['2']) == \
        'LIN_00380excl.acc.type=2ToML07'


def test_canonical_var_name_plain_name_appends_suffix():
    # Plain variable name (no excl token) with @2A@ types → suffix appended
    # e.g. A159_09 has 'Account Type 2' row but no 'excl' in the variable name
    assert _canonical_var_name('A159_09', ['2']) == 'A159_09excl.acc.type=2'


def test_canonical_var_name_plain_name_multi_type():
    # Plain variable name + multiple @2A@ types → comma-separated, sorted
    assert _canonical_var_name('A159_09', ['2', '1']) == 'A159_09excl.acc.type=1,2'


def test_canonical_var_name_plain_name_with_tom():
    # Plain name with ToM/TOM suffix and no excl → suffix inserted before ToM/TOM
    assert _canonical_var_name('A159_09ToML09', ['2']) == 'A159_09excl.acc.type=2ToML09'
    assert _canonical_var_name('A159_09TOML07', ['2']) == 'A159_09excl.acc.type=2TOML07'


def test_canonical_var_name_no_types_no_excl_unchanged():
    # No types and no excl text in name → unchanged
    assert _canonical_var_name('A246', []) == 'A246'


def test_canonical_var_name_no_types_strips_existing_excl():
    # Existing excl text is unreliable — always strip it, even when no @2A@ types
    assert _canonical_var_name('LIN_00380excl.acc.type2', []) == 'LIN_00380'


def test_canonical_var_name_no_types_strips_excl_preserves_tom():
    # Strip excl text but preserve ToM/TOM suffix
    assert _canonical_var_name('LIN_00380excl.acc.type2ToML07', []) == 'LIN_00380ToML07'


def test_canonical_var_name_types_sorted():
    # Types should be sorted numerically — 4 before 2 in input, 1,4 in output would be wrong
    assert _canonical_var_name('LIN_00380excl.type4,1', ['4', '1']) == \
        'LIN_00380excl.acc.type=1,4'


# ---------------------------------------------------------------------------
# _build_excl_formula — replaces variable names in formula using ExclAccountTypes
# ---------------------------------------------------------------------------

def test_build_excl_formula_replaces_single_variable():
    formula = 'ABS(VAL_YTD(IAN_00051excl.acc.type=2))>=CONST(1,'+"'USD','E')"
    variables = {
        0: {'Variable': 'IAN_00051excl.acc.type=2', 'ExclAccountTypes': ['2']},
    }
    result = _build_excl_formula(formula, variables)
    assert 'IAN_00051excl.acc.type=2' in result


def test_build_excl_formula_fixes_aff_notation():
    # The key gap: excl.2-Aff in formula should become excl.acc.type=2
    formula = 'ABS(VAL_YTD(SLSTlLSAS_11520excl.2-Aff))'
    variables = {
        0: {'Variable': 'SLSTlLSAS_11520excl.2-Aff', 'ExclAccountTypes': ['2']},
    }
    result = _build_excl_formula(formula, variables)
    assert 'excl.acc.type=2' in result
    assert 'excl.2-Aff' not in result


def test_build_excl_formula_no_excl_types_unchanged():
    formula = 'VAL_YTD(A246)+VAL_YTD(B123)>=CONST(0,'+"'USD','E')"
    variables = {
        0: {'Variable': 'A246', 'ExclAccountTypes': []},
        1: {'Variable': 'B123', 'ExclAccountTypes': []},
    }
    assert _build_excl_formula(formula, variables) == formula


def test_build_excl_formula_strips_excl_when_no_types():
    # Variable name has pre-written excl text but no @2A@ rows → strip the excl text
    formula = 'VAL_YTD(LIN_00380excl.2-Aff)>=CONST(0,'+"'USD','E')"
    variables = {
        0: {'Variable': 'LIN_00380excl.2-Aff', 'ExclAccountTypes': []},
    }
    result = _build_excl_formula(formula, variables)
    assert 'LIN_00380' in result
    assert 'excl' not in result
    assert '2-Aff' not in result


def test_build_excl_formula_multi_type():
    formula = 'ABS(VAL_YTD(LIN_00380excl.acc.type:1,4TOML09))'
    variables = {
        0: {'Variable': 'LIN_00380excl.acc.type:1,4TOML09', 'ExclAccountTypes': ['1', '4']},
    }
    result = _build_excl_formula(formula, variables)
    assert 'excl.acc.type=1,4' in result
    assert 'TOML09' in result


def test_build_excl_formula_plain_name_appends_suffix():
    # A159_09 case: variable name contains no excl token but @2A@ Account Type 2
    # row was captured → formula should be rewritten with the suffix appended
    formula = 'ABS(VAL_YTD(A159_09))>=CONST(0,'+"'USD','E')"
    variables = {
        0: {'Variable': 'A159_09', 'ExclAccountTypes': ['2']},
    }
    result = _build_excl_formula(formula, variables)
    assert 'A159_09excl.acc.type=2' in result


def test_build_excl_formula_plain_name_multi_type():
    formula = 'VAL_YTD(A159_09)'
    variables = {
        0: {'Variable': 'A159_09', 'ExclAccountTypes': ['2', '1']},
    }
    result = _build_excl_formula(formula, variables)
    assert 'A159_09excl.acc.type=1,2' in result


# ---------------------------------------------------------------------------
# _get_x_check_information — ExclAccountTypes captured from @2A@ rows
# ---------------------------------------------------------------------------

def _make_block(*lines):
    """Helper: dict[int, str] from a sequence of lines."""
    return {i: line for i, line in enumerate(lines)}


def test_get_x_check_information_excl_account_types_single():
    # Variable with one @2A@ Account Type row → ExclAccountTypes=['2']
    block = _make_block(
        _FORMULA_HEADER,
        'VAL_YTD(IAN_00051excl.acc.type=2)',
        _BLOCK_END,
        _VAR_HEADER,
        'IAN_00051excl.acc.type=2',
        _BLANK_LINE,
        '|-FS Account @2A@ Account Type 2 affiliated |',
        _BLOCK_END,
    )
    result = _get_x_check_information(block)
    var = result['Variables'][0]
    assert var['ExclAccountTypes'] == ['2']


def test_get_x_check_information_excl_account_types_multi():
    # Variable with two @2A@ Account Type rows → ExclAccountTypes=['1','4']
    block = _make_block(
        _FORMULA_HEADER,
        'VAL_YTD(LIN_00380excl.acc.type:1,4)',
        _BLOCK_END,
        _VAR_HEADER,
        'LIN_00380excl.acc.type:1,4',
        _BLANK_LINE,
        '|-FS Account @2A@ Account Type 1 3rd party |',
        '|-FS Account @2A@ Account Type 4 linked |',
        _BLOCK_END,
    )
    result = _get_x_check_information(block)
    var = result['Variables'][0]
    assert var['ExclAccountTypes'] == ['1', '4']


def test_get_x_check_information_no_excl_account_types():
    # Normal variable with no @2A@ Account Type rows → ExclAccountTypes=[]
    block = _make_block(
        _FORMULA_HEADER,
        'VAL_YTD(A246)+VAL_YTD(B123)',
        _BLOCK_END,
        _VAR_HEADER,
        'A246',
        _BLANK_LINE,
        _BLOCK_END,
    )
    result = _get_x_check_information(block)
    var = result['Variables'].get(0, {})
    assert var.get('ExclAccountTypes', []) == []
