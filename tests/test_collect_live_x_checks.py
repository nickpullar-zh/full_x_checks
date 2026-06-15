"""
Tests for the Collect Live X-Checks standalone task: writes a .txt of in-scope
X-Check Nos and copies the same text to the clipboard.
"""
import os
from unittest.mock import patch

import pandas as pd
import pytest
from openpyxl import Workbook

from file_upload_config import UploadTaskConfig
from strategies.collect_live_x_checks import CollectLiveXChecks


def _write_xlsx(tmp_path, headers, rows):
    path = os.path.join(str(tmp_path), "ebx.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "cross checks all"
    ws.append(headers)
    for r in rows:
        ws.append(r)
    wb.save(path)
    return path


def _make_strategy():
    cfg = UploadTaskConfig(task_name="Collect Live X-Checks", file_fields=[])
    s = CollectLiveXChecks(cfg)
    s.log = []
    return s


def _files_dict(path, out_dir):
    return {
        "files":           {"X-Checks Publication File": path},
        "sheet_names":     {"X-Checks Publication File": "cross checks all"},
        "output_directory": out_dir,
        "timestamp":       "20260615_120000",
    }


@patch.object(CollectLiveXChecks, "_copy_to_clipboard")
def test_writes_txt_and_copies_clipboard(mock_clip, tmp_path):
    headers = ["X-Check No.", "Status", "Type of Change", "Exclude Z-Core", "Category"]
    rows = [
        ["A1", "ACTIVE",   "Modified", "",  ""],
        ["B2", "INACTIVE", "Modified", "",  ""],
        ["C3", "ACTIVE",   "New",      "X", ""],   # Excluded by Z-Core
        ["D4", "ACTIVE",   "Modified", "",  ""],
    ]
    path = _write_xlsx(tmp_path, headers, rows)
    df = pd.read_excel(path, sheet_name="cross checks all")

    s = _make_strategy()
    out_dir = str(tmp_path)
    s.process({"X-Checks Publication File": df}, _files_dict(path, out_dir))

    out = os.path.join(out_dir, "20260615_120000_X-Check_Nos.txt")
    assert open(out).read() == "A1\nD4"
    mock_clip.assert_called_once_with("A1\nD4")


@patch.object(CollectLiveXChecks, "_copy_to_clipboard")
def test_no_in_scope_writes_nothing(mock_clip, tmp_path):
    headers = ["X-Check No.", "Status", "Type of Change", "Exclude Z-Core", "Category"]
    rows = [["A1", "INACTIVE", "Modified", "", ""]]
    path = _write_xlsx(tmp_path, headers, rows)
    df = pd.read_excel(path, sheet_name="cross checks all")

    s = _make_strategy()
    out_dir = str(tmp_path)
    s.process({"X-Checks Publication File": df}, _files_dict(path, out_dir))

    out = os.path.join(out_dir, "20260615_120000_X-Check_Nos.txt")
    assert not os.path.exists(out)
    mock_clip.assert_not_called()


@patch.object(CollectLiveXChecks, "_copy_to_clipboard", side_effect=RuntimeError("boom"))
def test_clipboard_failure_does_not_break_run(mock_clip, tmp_path):
    headers = ["X-Check No.", "Status", "Type of Change", "Exclude Z-Core", "Category"]
    rows = [["A1", "ACTIVE", "Modified", "", ""]]
    path = _write_xlsx(tmp_path, headers, rows)
    df = pd.read_excel(path, sheet_name="cross checks all")

    s = _make_strategy()
    out_dir = str(tmp_path)
    s.process({"X-Checks Publication File": df}, _files_dict(path, out_dir))

    # File still written despite clipboard failure
    out = os.path.join(out_dir, "20260615_120000_X-Check_Nos.txt")
    assert open(out).read() == "A1"
    # Log captured the failure
    assert any("Clipboard copy failed" in str(entry) for entry in s.log)


def test_missing_ebx_aborts_cleanly(tmp_path):
    s = _make_strategy()
    out_dir = str(tmp_path)
    files = {
        "files":            {},
        "sheet_names":      {},
        "output_directory": out_dir,
        "timestamp":        "20260615_120000",
    }
    # Should not raise
    s.process({}, files)
    assert os.listdir(out_dir) == []
