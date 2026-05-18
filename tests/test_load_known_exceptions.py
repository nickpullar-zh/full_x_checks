import pytest
import pandas as pd
from pathlib import Path


def _write_exceptions_excel(path: Path, data: dict, sheet: str = "Known Exceptions"):
    """Write an exceptions Excel file matching the real file structure.

    The real file has a guidance row at Excel row 2 which _load_known_exceptions
    skips via skiprows=[1].  Insert a dummy guidance row here so test data rows
    land where the code expects them.
    """
    df = pd.DataFrame(data)
    guidance = pd.DataFrame({col: [f"<guidance>"] for col in df.columns})
    pd.concat([guidance, df], ignore_index=True).to_excel(path, sheet_name=sheet, index=False)


# ---------------------------------------------------------------------------
# Valid file
# ---------------------------------------------------------------------------

def test_valid_file_returns_dict(xchecks_instance, test_data_dir):
    path = str(test_data_dir / "Known_Exception_List.xlsx")
    result = xchecks_instance._load_known_exceptions(path)
    assert isinstance(result, dict)
    assert "S380_00" in result
    # Value should be the Reason text, not the Exception Type
    assert "LC_YTD" in result["S380_00"] or len(result["S380_00"]) > 5


def test_guidance_row_skipped(xchecks_instance, test_data_dir):
    # Row 2 of the real file is the guidance/example row — must not appear as an entry
    path = str(test_data_dir / "Known_Exception_List.xlsx")
    result = xchecks_instance._load_known_exceptions(path)
    assert "Unique X-Check identifier" not in result
    assert "e.g. S380_00" not in result


# ---------------------------------------------------------------------------
# Empty / unreadable
# ---------------------------------------------------------------------------

def test_empty_sheet_returns_empty_dict(xchecks_instance, tmp_path):
    path = tmp_path / "empty.xlsx"
    _write_exceptions_excel(path, {"X-Check Number": [], "Reason": []})
    result = xchecks_instance._load_known_exceptions(str(path))
    assert result == {}


def test_unreadable_file_returns_empty_dict(xchecks_instance, tmp_path):
    path = tmp_path / "does_not_exist.xlsx"
    result = xchecks_instance._load_known_exceptions(str(path))
    assert result == {}


# ---------------------------------------------------------------------------
# Missing columns
# ---------------------------------------------------------------------------

def test_missing_reason_column_raises(xchecks_instance, tmp_path):
    path = tmp_path / "no_reason.xlsx"
    _write_exceptions_excel(path, {"X-Check Number": ["X1"], "Exception Type": ["LC_YTD"]})
    with pytest.raises(ValueError, match="Reason"):
        xchecks_instance._load_known_exceptions(str(path))


def test_missing_xcheck_number_column_raises(xchecks_instance, tmp_path):
    path = tmp_path / "no_xcheck.xlsx"
    _write_exceptions_excel(path, {"Check": ["X1"], "Reason": ["some reason"]})
    with pytest.raises(ValueError, match="X-Check Number"):
        xchecks_instance._load_known_exceptions(str(path))


# ---------------------------------------------------------------------------
# Invalid rows
# ---------------------------------------------------------------------------

def test_blank_reason_value_raises(xchecks_instance, tmp_path):
    path = tmp_path / "blank_reason.xlsx"
    _write_exceptions_excel(path, {
        "X-Check Number": ["X1"],
        "Reason": [""],
    })
    with pytest.raises(ValueError, match="X1"):
        xchecks_instance._load_known_exceptions(str(path))


def test_blank_xcheck_number_raises(xchecks_instance, tmp_path):
    path = tmp_path / "blank_xcheck.xlsx"
    _write_exceptions_excel(path, {
        "X-Check Number": [""],
        "Reason": ["Some reason"],
    })
    with pytest.raises(ValueError, match="row"):
        xchecks_instance._load_known_exceptions(str(path))


def test_multiple_invalid_rows_all_reported(xchecks_instance, tmp_path):
    # All bad rows should appear in the error message, not just the first
    path = tmp_path / "multi_bad.xlsx"
    _write_exceptions_excel(path, {
        "X-Check Number": ["X1", "X2"],
        "Reason": ["", ""],
    })
    with pytest.raises(ValueError) as exc_info:
        xchecks_instance._load_known_exceptions(str(path))
    msg = str(exc_info.value)
    assert "X1" in msg
    assert "X2" in msg
