from strategies.x_checks.ebx_extraction import _should_use_qu, _create_formula


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
