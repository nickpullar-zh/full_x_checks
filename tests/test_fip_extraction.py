from strategies.x_checks.fip_extraction import _safe_split, _normalise_excl_suffix


def test_safe_split_valid_index():
    assert _safe_split('a b c', 1) == 'b'


def test_safe_split_first_token():
    assert _safe_split('hello world', 0) == 'hello'


def test_safe_split_out_of_bounds_returns_default():
    # REGRESSION: bare line.split()[5] raised IndexError on short lines;
    # _safe_split returns default instead
    assert _safe_split('a b', 5) == ''


def test_safe_split_custom_default():
    assert _safe_split('a', 5, 'MISSING') == 'MISSING'


def test_safe_split_empty_line():
    assert _safe_split('', 0) == ''


def test_safe_split_extra_whitespace():
    # split() with no args collapses whitespace — index is token-based not char-based
    assert _safe_split('  a   b  ', 1) == 'b'


# ---------------------------------------------------------------------------
# _normalise_excl_suffix — all 7 FIP notation variants
# ---------------------------------------------------------------------------

def test_normalise_excl_variant1_equals_spaces():
    assert _normalise_excl_suffix("VAL_YTD(LIN_00380excl. acc. type = 2)") == \
        "VAL_YTD(LIN_00380excl.acc.type=2)"


def test_normalise_excl_variant2_equals_spaces_type1():
    assert _normalise_excl_suffix("VAL_YTD(LIN_00380excl. acc. type = 1)") == \
        "VAL_YTD(LIN_00380excl.acc.type=1)"


def test_normalise_excl_variant3_colon_space_type2():
    assert _normalise_excl_suffix("VAL_YTD(LIN_00380excl. acc. type: 2)") == \
        "VAL_YTD(LIN_00380excl.acc.type=2)"


def test_normalise_excl_variant4_colon_nospace_multitype():
    assert _normalise_excl_suffix("VAL_YTD(LIN_00380excl. acc. type:1,4)") == \
        "VAL_YTD(LIN_00380excl.acc.type=1,4)"


def test_normalise_excl_variant5_concatenated():
    assert _normalise_excl_suffix("VAL_YTD(LIN_00380ffexcl.acc.type2)") == \
        "VAL_YTD(LIN_00380ffexcl.acc.type=2)"


def test_normalise_excl_variant6_space_separator():
    assert _normalise_excl_suffix("VAL_YTD(LIN_00380excl. acc. type 2)") == \
        "VAL_YTD(LIN_00380excl.acc.type=2)"


def test_normalise_excl_variant7_acct_abbreviation():
    assert _normalise_excl_suffix("VAL_YTD(LIN_00380excl.acct.type 2)") == \
        "VAL_YTD(LIN_00380excl.acc.type=2)"


def test_normalise_excl_typo_exl():
    # 'exl' (missing 'c') variant — seen in AL167_00 and AS130_00
    assert _normalise_excl_suffix("VAL_YTD(IAN_00023exl.acc.type:2)") == \
        "VAL_YTD(IAN_00023excl.acc.type=2)"


def test_normalise_excl_tom_suffix_preserved():
    # TOM movement-type tag after the excl suffix must be preserved
    assert _normalise_excl_suffix("VAL_YTD(LIN_00380excl. acc. type: 2 TOM L09)") == \
        "VAL_YTD(LIN_00380excl.acc.type=2 TOM L09)"


def test_normalise_excl_no_pattern_unchanged():
    formula = "VAL_YTD(A246)+VAL_YTD(B123)>=CONST(0,'USD','E')"
    assert _normalise_excl_suffix(formula) == formula


# ---------------------------------------------------------------------------
# Family B: parenthetical annotation variants (LA003–LA006 series)
# ---------------------------------------------------------------------------

def test_normalise_excl_family_b_only_2_affiliated():
    # LA003_09 / LA004_09: (only 2-affiliated) → (only.type=2)
    result = _normalise_excl_suffix("LC_YTD(BSN (only 2-affiliated) ToM 354ff)")
    assert result == "LC_YTD(BSN (only.type=2) ToM 354ff)"


def test_normalise_excl_family_b_only_2_affiliated_spaced_variant():
    # LA004_09 uses TOM (uppercase) and hyphen — same normalisation
    result = _normalise_excl_suffix("LC_YTD(BSN (only 2-affiliated) TOM 670ff)")
    assert result == "LC_YTD(BSN (only.type=2) TOM 670ff)"


def test_normalise_excl_family_b_without_3rd_party():
    # LA003_09 / LA004_09 / LA006_09: (without 3rd party) → (excl.type=1)
    result = _normalise_excl_suffix("LC_YTD(SN_12895ff (without 3rd party) ToM 660ff)")
    assert result == "LC_YTD(SN_12895ff (excl.type=1) ToM 660ff)"


def test_normalise_excl_family_b_without_affiliated():
    # LA005_09: (without affiliated) → (excl.type=2); slash separator preserved
    result = _normalise_excl_suffix("LC_YTD(SN_12895ff(without affiliated)/ToM 660ff)")
    assert result == "LC_YTD(SN_12895ff(excl.type=2)/ToM 660ff)"


def test_normalise_excl_family_b_tom_and_surrounding_text_preserved():
    # Confirm nothing outside the annotation is changed
    result = _normalise_excl_suffix(
        "ABS(LC_YTD(28832)+LC_YTD(BSN (only 2-affiliated) TOM 670ff)"
        "-LC_YTD(SN_12895ff (without 3rd party) ToM 670ff))<=CONST(5,'USD','E')"
    )
    assert "(only.type=2)" in result
    assert "(excl.type=1)" in result
    assert "LC_YTD(28832)" in result
    assert "CONST(5,'USD','E')" in result
