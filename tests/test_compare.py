from strategies.x_checks.compare import _compare_formulas, _compare_variables, compare


# ---------------------------------------------------------------------------
# _compare_formulas
# ---------------------------------------------------------------------------

def test_formulas_exact_match():
    match, formula = _compare_formulas("VAL_YTD(A)>=0", "VAL_YTD(A)>=0")
    assert match is True
    assert formula == "VAL_YTD(A)>=0"


def test_formulas_case_insensitive_match():
    match, _ = _compare_formulas("val_ytd(A)>=0", "VAL_YTD(A)>=0")
    assert match is True


def test_formulas_genuine_mismatch():
    match, _ = _compare_formulas("VAL_YTD(A)>=0", "VAL_YTD(B)>=0")
    assert match is False


def test_formulas_empty_bracket_vars_no_crash():
    # REGRESSION: formula contains '+' but no VAL_YTD calls — bracket_vars is []
    # Before fix: bracket_vars[0] raised IndexError
    # After fix: guarded with 'if bracket_vars:', returns (False, formula) cleanly
    fip = "CONST(1,'USD','E')+CONST(2,'USD','E')"
    ebx = "CONST(2,'USD','E')+CONST(1,'USD','E')"
    match, returned_formula = _compare_formulas(fip, ebx)
    assert match is False
    assert returned_formula == fip  # formula unchanged when no replacement made


def test_formulas_addition_reorder_no_crash():
    # Addition-only path: same var sets, different order — must not crash even if
    # it cannot fully reorder. The reorder logic fires but produces a valid (bool, str)
    fip = "VAL_YTD(A)+VAL_YTD(B)>=0"
    ebx = "VAL_YTD(B)+VAL_YTD(A)>=0"
    match, returned = _compare_formulas(fip, ebx)
    assert isinstance(match, bool)
    assert isinstance(returned, str)


# ---------------------------------------------------------------------------
# _compare_variables
# ---------------------------------------------------------------------------

def test_variables_same_order():
    assert _compare_variables("A|B", "A|B") is True


def test_variables_different_order():
    assert _compare_variables("A|B", "B|A") is True


def test_variables_case_insensitive():
    assert _compare_variables("name:a;fs account:a246", "Name:A;FS Account:A246") is True


def test_variables_both_empty():
    # REGRESSION: '' split on '|' gives [''], filtering to [] — both sides empty
    # Before fix: [''] == [''] returned True for empty, which was correct,
    # but empty-string parts from pipe-only strings could cause false matches
    assert _compare_variables("", "") is True


def test_variables_one_side_empty():
    assert _compare_variables("", "A|B") is False


def test_variables_pipe_separated_match():
    assert _compare_variables("A|B|C", "C|A|B") is True


def test_variables_mismatch():
    assert _compare_variables("A|B", "A|C") is False


# ---------------------------------------------------------------------------
# compare (integration of the two sides)
# ---------------------------------------------------------------------------

def test_compare_missing_fip():
    ebx = [{"X-Check Number": "X1", "EBX Formula": "VAL_YTD(A)>=0", "EBX Variables": "A"}]
    fip = []
    rows = compare(ebx, fip)
    assert len(rows) == 1
    row = rows[0]
    assert row["X-Check Number"] == "X1"
    assert row["Formula Match"] == "Not Found"
    assert row["FIP Formula"] == "Not Found"
    assert row["EBX Formula"] == "VAL_YTD(A)>=0"


def test_compare_missing_ebx():
    ebx = []
    fip = [{"X-Check Number": "X1", "FIP Formula": "VAL_YTD(A)>=0",
            "FIP Variables": "A", "FIP Variable (Builder)": "A"}]
    rows = compare(ebx, fip)
    assert len(rows) == 1
    row = rows[0]
    assert row["Formula Match"] == "Not Found"
    assert row["EBX Formula"] == "Not Found"


def test_compare_both_present_match():
    ebx = [{"X-Check Number": "X1", "EBX Formula": "VAL_YTD(A)>=0", "EBX Variables": "A"}]
    fip = [{"X-Check Number": "X1", "FIP Formula": "VAL_YTD(A)>=0",
            "FIP Variables": "A", "FIP Variable (Builder)": "A"}]
    rows = compare(ebx, fip)
    assert len(rows) == 1
    assert rows[0]["Formula Match"] == "Match"
    assert rows[0]["Variables Match"] == "Match"


def test_compare_both_present_mismatch():
    ebx = [{"X-Check Number": "X1", "EBX Formula": "VAL_YTD(A)>=0", "EBX Variables": "A"}]
    fip = [{"X-Check Number": "X1", "FIP Formula": "VAL_YTD(B)>=0",
            "FIP Variables": "B", "FIP Variable (Builder)": "B"}]
    rows = compare(ebx, fip)
    assert rows[0]["Formula Match"] == "MisMatch"
    assert rows[0]["Variables Match"] == "MisMatch"


def test_compare_deduplicates_ebx():
    # Two EBX rows for the same X-Check — only the first should be used
    ebx = [
        {"X-Check Number": "X1", "EBX Formula": "FIRST", "EBX Variables": "A"},
        {"X-Check Number": "X1", "EBX Formula": "SECOND", "EBX Variables": "B"},
    ]
    fip = [{"X-Check Number": "X1", "FIP Formula": "FIRST",
            "FIP Variables": "A", "FIP Variable (Builder)": "A"}]
    rows = compare(ebx, fip)
    assert len(rows) == 1
    assert rows[0]["EBX Formula"] == "FIRST"
