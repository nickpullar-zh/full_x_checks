from strategies.x_checks.fip_extraction import _safe_split


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
