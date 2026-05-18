from strategies.x_checks.ebx_extraction import _should_use_qu, _should_use_pct, _create_formula


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


def test_create_formula_use_pct_nonzero_limit():
    # Limit = 5 → '5,000000%'
    variables = [{"Variable-Name": "A246", "Operator": "+"}]
    row = {"Operator 1": ">=", "Operator 2": "", "Limit 1": "5", "Limit 2": "", "%": "X"}
    formula = _create_formula(variables, False, row, use_pct=True)
    assert formula == "VAL_YTD(A246)>='5,000000%'"


def test_create_formula_use_pct_no_const_wrapper():
    # use_pct must NOT wrap the right-hand side in CONST(...)
    variables = [{"Variable-Name": "A246", "Operator": "+"}]
    row = {"Operator 1": ">=", "Operator 2": "", "Limit 1": "100", "Limit 2": "", "%": "X"}
    formula = _create_formula(variables, False, row, use_pct=True)
    assert "CONST" not in formula
    assert "'100,000000%'" in formula
