import pandas as pd
from strategies.x_checks.ebx_extraction import (
    _should_use_qu, _should_use_pct, _create_formula,
    _VERSION_GAAP_PREFIX, extract_ebx,
)


# ---------------------------------------------------------------------------
# _should_use_qu
# ---------------------------------------------------------------------------

def test_should_use_qu_match():
    dict_account = {"S20210DE": {}, "S23070DE": {}}
    qu_accounts = {"S20210DE", "S99999XX"}
    assert _should_use_qu(dict_account, qu_accounts) is True


def test_should_use_qu_no_match():
    dict_account = {"A246": {}, "A247": {}}
    qu_accounts = {"S20210DE", "S23070DE"}
    assert _should_use_qu(dict_account, qu_accounts) is False


def test_should_use_qu_empty_set():
    dict_account = {"A246": {}}
    assert _should_use_qu(dict_account, set()) is False


def test_should_use_qu_none():
    dict_account = {"A246": {}}
    assert _should_use_qu(dict_account, None) is False


# ---------------------------------------------------------------------------
# _create_formula with use_qu
# ---------------------------------------------------------------------------

def _make_row():
    return {
        "Operator 1": ">=", "Operator 2": "",
        "Limit 1": "0", "Limit 2": "",
    }


def test_create_formula_use_qu_produces_qu_ytd():
    variables = [{"Variable-Name": "S20210DE", "Operator": "+"}]
    formula = _create_formula(variables, False, _make_row(), use_qu=True)
    assert "QU_YTD" in formula
    assert "VAL_YTD" not in formula


def test_create_formula_use_qu_takes_priority_over_use_lc():
    # use_qu should win when both flags are True
    variables = [{"Variable-Name": "S20210DE", "Operator": "+"}]
    formula = _create_formula(variables, False, _make_row(), use_lc=True, use_qu=True)
    assert "QU_YTD" in formula
    assert "LC_YTD" not in formula
    assert "VAL_YTD" not in formula


def test_create_formula_use_qu_const_unchanged():
    # CONST stays as CONST (not CONST_LC) for QU accounts
    variables = [{"Variable-Name": "S20210DE", "Operator": "+"}]
    row = {"Operator 1": ">=", "Operator 2": "", "Limit 1": "1000", "Limit 2": ""}
    formula = _create_formula(variables, False, row, use_qu=True)
    assert "QU_YTD" in formula
    assert "CONST(" in formula
    assert "CONST_LC" not in formula


def test_create_formula_default_still_val_ytd():
    # Ensure existing behaviour unchanged when use_qu=False
    variables = [{"Variable-Name": "A246", "Operator": "+"}]
    formula = _create_formula(variables, False, _make_row())
    assert "VAL_YTD" in formula
    assert "QU_YTD" not in formula


# ---------------------------------------------------------------------------
# _should_use_pct
# ---------------------------------------------------------------------------

def test_should_use_pct_x_in_column():
    assert _should_use_pct({'%': 'X'}) is True


def test_should_use_pct_empty_column():
    assert _should_use_pct({'%': ''}) is False


def test_should_use_pct_missing_column():
    assert _should_use_pct({}) is False


# ---------------------------------------------------------------------------
# _create_formula with use_pct
# ---------------------------------------------------------------------------

def test_create_formula_use_pct_zero_limit():
    # Limit = 0 → '0,000000%'  (the real-world case from S002_00)
    variables = [{"Variable-Name": "S73101", "Operator": "+"}]
    row = {"Operator 1": ">", "Operator 2": "", "Limit 1": "0", "Limit 2": "", "%": "X"}
    formula = _create_formula(variables, False, row, use_pct=True)
    assert formula == "VAL_YTD(S73101)>'0,000000%'"


def test_create_formula_use_pct_nonzero_integer_limit():
    # Limit = 5 → '5,000000%'
    variables = [{"Variable-Name": "A246", "Operator": "+"}]
    row = {"Operator 1": ">=", "Operator 2": "", "Limit 1": "5", "Limit 2": "", "%": "X"}
    formula = _create_formula(variables, False, row, use_pct=True)
    assert formula == "VAL_YTD(A246)>='5,000000%'"


def test_create_formula_use_pct_decimal_limit():
    # Limit = 1.5 → '1,500000%'  (would be wrong '1,000000%' if truncated to int first)
    variables = [{"Variable-Name": "A246", "Operator": "+"}]
    row = {"Operator 1": ">=", "Operator 2": "", "Limit 1": "1.5", "Limit 2": "", "%": "X"}
    formula = _create_formula(variables, False, row, use_pct=True)
    assert formula == "VAL_YTD(A246)>='1,500000%'"


def test_create_formula_use_pct_no_const_wrapper():
    # use_pct must NOT wrap the right-hand side in CONST(...)
    variables = [{"Variable-Name": "A246", "Operator": "+"}]
    row = {"Operator 1": ">=", "Operator 2": "", "Limit 1": "100", "Limit 2": "", "%": "X"}
    formula = _create_formula(variables, False, row, use_pct=True)
    assert "CONST" not in formula
    assert "'100,000000%'" in formula


# ---------------------------------------------------------------------------
# _VERSION_GAAP_PREFIX constant
# ---------------------------------------------------------------------------

def test_version_gaap_prefix_contains_expected_mappings():
    assert _VERSION_GAAP_PREFIX['100'] == 'IFRSN'
    assert _VERSION_GAAP_PREFIX['190'] == 'IFRSN'
    assert _VERSION_GAAP_PREFIX['600'] == 'SLST'
    assert _VERSION_GAAP_PREFIX['800'] == 'SST'


# ---------------------------------------------------------------------------
# Version spanning via extract_ebx
# ---------------------------------------------------------------------------

def _make_ebx_df(rows: list[dict]) -> pd.DataFrame:
    """Build a minimal EBX DataFrame from a list of row dicts."""
    defaults = {
        'X-Check No.': 'X1', 'Account No.': 'A', 'SubA No.': '', 'Account description': '',
        'SubA Description': '', 'Operator (X-Check Term)': '+', 'Absolute (result)': '',
        'Operator 1': '>=', 'Operator 2': '', 'Limit 1': '0', 'Limit 2': '', '%': '',
        'Category': '', 'Version Spanning Validation': '', 'Ending Balance Prior Year': '',
        'SST account category': '', 'SII account category': '',
    }
    return pd.DataFrame([{**defaults, **r} for r in rows])


def test_version_spanning_suffix_mode():
    # Same account with two different version numbers → suffix mode: accountvN
    df = _make_ebx_df([
        {'Account No.': 'A246', 'Operator (X-Check Term)': '+', 'Version Spanning Validation': '100'},
        {'Account No.': 'A246', 'Operator (X-Check Term)': '-', 'Version Spanning Validation': '800'},
    ])
    results = extract_ebx(df, apply_version_spanning=True)
    assert len(results) == 1
    formula = results[0]['EBX Formula']
    assert 'A246v100' in formula
    assert 'A246v800' in formula


def test_version_spanning_prefix_mode():
    # Different accounts each with one version → prefix mode: GAAPaccount
    df = _make_ebx_df([
        {'Account No.': 'A246', 'Operator (X-Check Term)': '+', 'Version Spanning Validation': '190'},
        {'Account No.': 'B100', 'Operator (X-Check Term)': '-', 'Version Spanning Validation': '600'},
    ])
    results = extract_ebx(df, apply_version_spanning=True)
    formula = results[0]['EBX Formula']
    assert 'IFRSNА246'.replace('А', 'A') in formula or 'IFRSNA246' in formula  # GAAP prefix
    assert 'SLSTB100' in formula


def test_version_spanning_disabled_no_change():
    df = _make_ebx_df([
        {'Account No.': 'A246', 'Version Spanning Validation': '100'},
        {'Account No.': 'A246', 'Version Spanning Validation': '800'},
    ])
    results = extract_ebx(df, apply_version_spanning=False)
    formula = results[0]['EBX Formula']
    assert 'v100' not in formula
    assert 'v800' not in formula


# ---------------------------------------------------------------------------
# Prior year balance via _create_formula
# ---------------------------------------------------------------------------

def test_prior_year_balance_uses_p_val_per():
    variables = [{"Variable-Name": "29201ff", "Operator": "+",
                  "use_p_val_per": True, "py_suffix": False}]
    row = {"Operator 1": "=", "Operator 2": "", "Limit 1": "0", "Limit 2": ""}
    formula = _create_formula(variables, False, row)
    assert "P_VAL_PER(29201ff,'0','1')" in formula
    assert "VAL_YTD" not in formula


def test_prior_year_balance_py_suffix_applied():
    variables = [{"Variable-Name": "29201ff", "Operator": "+",
                  "use_p_val_per": True, "py_suffix": True}]
    row = {"Operator 1": "=", "Operator 2": "", "Limit 1": "0", "Limit 2": ""}
    formula = _create_formula(variables, False, row)
    assert "P_VAL_PER(29201ffPY,'0','1')" in formula


def test_prior_year_balance_no_py_suffix_equity():
    # Category = Shareholders' Equity → no PY suffix (py_suffix=False)
    variables = [{"Variable-Name": "29201ff", "Operator": "+",
                  "use_p_val_per": True, "py_suffix": False}]
    row = {"Operator 1": "=", "Operator 2": "", "Limit 1": "0", "Limit 2": ""}
    formula = _create_formula(variables, False, row)
    assert "P_VAL_PER(29201ff,'0','1')" in formula
    assert "PY" not in formula


def test_prior_year_balance_overrides_lc_ytd():
    # When use_p_val_per is present, use_lc should be suppressed — VAL_YTD not LC_YTD
    variables = [
        {"Variable-Name": "var1", "Operator": "+",  "use_p_val_per": True,  "py_suffix": False},
        {"Variable-Name": "var2", "Operator": "+",  "use_p_val_per": False, "py_suffix": False},
    ]
    row = {"Operator 1": "=", "Operator 2": "", "Limit 1": "0", "Limit 2": ""}
    formula = _create_formula(variables, False, row, use_lc=True)
    assert "VAL_YTD(var2)" in formula  # LC_YTD suppressed
    assert "LC_YTD" not in formula


def test_prior_year_balance_disabled_uses_val_ytd():
    variables = [{"Variable-Name": "A246", "Operator": "+",
                  "use_p_val_per": False, "py_suffix": False}]
    row = {"Operator 1": ">=", "Operator 2": "", "Limit 1": "0", "Limit 2": ""}
    formula = _create_formula(variables, False, row)
    assert "VAL_YTD(A246)" in formula
    assert "P_VAL_PER" not in formula
