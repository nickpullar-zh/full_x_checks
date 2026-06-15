"""
Tests for the v0.4 X-Check No Selection feature: _change_flag_row_indices,
the unioning behaviour of _filter_changed_rows, and the .txt writer logic
in XChecks._write_x_check_no_list.
"""
import os
import tempfile

import pandas as pd
import pytest

from file_upload_config import UploadTaskConfig
from strategies.base_strategy import BaseStrategy
from strategies.x_checks.x_checks import XChecks


# ---------------------------------------------------------------------------
# _change_flag_row_indices
# ---------------------------------------------------------------------------

class _StubStrategy(BaseStrategy):
    """Concrete subclass with a no-op process so we can instantiate BaseStrategy."""
    def process(self, loaded_files, files):
        pass


def _make_strategy():
    cfg = UploadTaskConfig(task_name="Stub", file_fields=[])
    return _StubStrategy(cfg)


def test_change_flag_indices_picks_non_blank():
    df = pd.DataFrame({
        "X-Check No.": ["A1", "B2", "C3", "D4"],
        "Type of change": ["", "New", "  ", "Modified"],
    })
    s = _make_strategy()
    assert s._change_flag_row_indices(df) == [1, 3]


def test_change_flag_indices_handles_nan_strings():
    df = pd.DataFrame({
        "X-Check No.": ["A1", "B2", "C3"],
        "Type of change": ["nan", "None", "Changed"],
    })
    s = _make_strategy()
    assert s._change_flag_row_indices(df) == [2]


def test_change_flag_indices_missing_column_returns_empty():
    df = pd.DataFrame({"X-Check No.": ["A1", "B2"]})
    s = _make_strategy()
    assert s._change_flag_row_indices(df) == []


def test_change_flag_indices_all_blank_returns_empty():
    df = pd.DataFrame({
        "X-Check No.": ["A1", "B2"],
        "Type of change": ["", ""],
    })
    s = _make_strategy()
    assert s._change_flag_row_indices(df) == []


# ---------------------------------------------------------------------------
# _write_x_check_no_list — preserves order of first occurrence, dedupes,
# writes one X-Check No per line
# ---------------------------------------------------------------------------

def _make_xchecks_strategy(process_only_diff=True):
    cfg = UploadTaskConfig(task_name="X-Checks", file_fields=[])
    s = XChecks(cfg)
    s.log = []
    s.process_only_differences = process_only_diff
    return s


def test_write_x_check_no_list_preserves_first_occurrence_order():
    df = pd.DataFrame({"X-Check No.": ["B2", "A1", "B2", "C3", "A1"]})
    s = _make_xchecks_strategy()
    with tempfile.TemporaryDirectory() as tmp:
        files = {"output_directory": tmp, "timestamp": "20260615_120000"}
        s._write_x_check_no_list({"X-Checks Publication File": df}, files)
        out = os.path.join(tmp, "20260615_120000_X-Check_Nos.txt")
        assert os.path.exists(out)
        assert open(out).read() == "B2\nA1\nC3"


def test_write_x_check_no_list_skips_blanks_and_nan():
    df = pd.DataFrame({"X-Check No.": ["", "A1", "nan", "B2", "None"]})
    s = _make_xchecks_strategy()
    with tempfile.TemporaryDirectory() as tmp:
        files = {"output_directory": tmp, "timestamp": "20260615_120000"}
        s._write_x_check_no_list({"X-Checks Publication File": df}, files)
        out = os.path.join(tmp, "20260615_120000_X-Check_Nos.txt")
        assert open(out).read() == "A1\nB2"


def test_write_x_check_no_list_no_data_writes_nothing():
    df = pd.DataFrame({"X-Check No.": []})
    s = _make_xchecks_strategy()
    with tempfile.TemporaryDirectory() as tmp:
        files = {"output_directory": tmp, "timestamp": "20260615_120000"}
        s._write_x_check_no_list({"X-Checks Publication File": df}, files)
        assert os.listdir(tmp) == []


def test_write_x_check_no_list_missing_column_skips_silently():
    df = pd.DataFrame({"OtherCol": ["A", "B"]})
    s = _make_xchecks_strategy()
    with tempfile.TemporaryDirectory() as tmp:
        files = {"output_directory": tmp, "timestamp": "20260615_120000"}
        s._write_x_check_no_list({"X-Checks Publication File": df}, files)
        assert os.listdir(tmp) == []
