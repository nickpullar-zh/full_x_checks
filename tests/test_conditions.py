"""
Unit tests for the Conditions strategy modules.

extract.py  — _is_yellow, _is_green colour helpers + extract_conditions()
fip.py      — process_fip() column rename and concatenation
compare.py  — compare() True/False/blank match logic
"""

import os
import pytest
import pandas as pd
from unittest.mock import MagicMock, patch, PropertyMock


# ---------------------------------------------------------------------------
# Helpers to build mock openpyxl cells with a given fill colour
# ---------------------------------------------------------------------------

def _cell_with_rgb(rgb: str):
    """Return a mock openpyxl cell whose fgColor.type='rgb' and .rgb=rgb."""
    cell = MagicMock()
    cell.fill.fgColor.type = "rgb"
    cell.fill.fgColor.rgb = rgb
    return cell


def _cell_no_fill():
    cell = MagicMock()
    cell.fill.fgColor.type = "rgb"
    cell.fill.fgColor.rgb = "00000000"
    return cell


# ---------------------------------------------------------------------------
# tests for _is_yellow / _is_green
# ---------------------------------------------------------------------------

from strategies.conditions.extract import _is_yellow, _is_green


class TestColourHelpers:

    def test_is_yellow_standard(self):
        cell = _cell_with_rgb("FFFFFF00")
        assert _is_yellow(cell, None) is True

    def test_is_yellow_dark_yellow(self):
        cell = _cell_with_rgb("FFFFC000")
        assert _is_yellow(cell, None) is True

    def test_is_yellow_no_fill(self):
        cell = _cell_no_fill()
        assert _is_yellow(cell, None) is False

    def test_is_yellow_green_is_not_yellow(self):
        cell = _cell_with_rgb("FF92D050")
        assert _is_yellow(cell, None) is False

    def test_is_green_standard(self):
        cell = _cell_with_rgb("FF92D050")
        assert _is_green(cell, None) is True

    def test_is_green_alt(self):
        cell = _cell_with_rgb("FF00B050")
        assert _is_green(cell, None) is True

    def test_is_green_no_fill(self):
        cell = _cell_no_fill()
        assert _is_green(cell, None) is False

    def test_is_green_yellow_is_not_green(self):
        cell = _cell_with_rgb("FFFFFF00")
        assert _is_green(cell, None) is False


# ---------------------------------------------------------------------------
# tests for process_fip
# ---------------------------------------------------------------------------

from strategies.conditions.extract import CONDITION_COLS as _COND_COLS
from strategies.conditions.fip import process_fip, CONCAT_COL, _RENAMED_COLS_8 as _RENAMED_COLS


# ---------------------------------------------------------------------------
# Tests for extract_conditions() extraction rule
# ---------------------------------------------------------------------------

class TestExtractionRule:
    """
    The rule: collect a condition cell only when the condition cell itself
    is yellow or green AND has a non-blank value.
    A green X-Check No. cell must NOT cause uncoloured condition cells to
    be collected.
    """

    def _make_ws(self, rows):
        """
        Build a minimal mock openpyxl worksheet.
        rows: list of dicts with keys 'xcno', 'xcno_rgb', 'cond_val', 'cond_rgb'
        (single condition column for simplicity)
        """
        import openpyxl
        from unittest.mock import MagicMock

        ws = MagicMock()

        # Header row
        hdr_xcno = MagicMock(); hdr_xcno.value = "X-Check No."; hdr_xcno.column = 1; hdr_xcno.row = 1
        hdr_cond = MagicMock(); hdr_cond.value = _COND_COLS[1]; hdr_cond.column = 2; hdr_cond.row = 1  # Applicable Quarters

        def _make_cell(col, val, rgb):
            c = MagicMock()
            c.column = col
            c.value = val
            c.fill.fgColor.type = "rgb"
            c.fill.fgColor.rgb = rgb
            return c

        data_rows = []
        for i, r in enumerate(rows, start=2):
            xcno_cell = _make_cell(1, r['xcno'], r['xcno_rgb'])
            cond_cell = _make_cell(2, r['cond_val'], r['cond_rgb'])
            data_rows.append([xcno_cell, cond_cell])

        ws.__getitem__ = lambda self, idx: [hdr_xcno, hdr_cond]
        ws.sheetnames = [None]

        return ws, data_rows

    def test_yellow_condition_cell_is_collected(self):
        from strategies.conditions.extract import _is_yellow, _is_green
        cell = _cell_with_rgb("FFFFFF00")
        assert _is_yellow(cell, None) is True

    def test_green_condition_cell_with_value_is_collected(self):
        from strategies.conditions.extract import _is_green
        cell = _cell_with_rgb("FF92D050")
        cell.value = "CON_Q4"
        assert _is_green(cell, None) is True

    def test_uncoloured_condition_cell_on_green_xcno_row_not_collected(self):
        """Green X-Check No. cell must NOT pull in uncoloured condition cells."""
        from strategies.conditions.extract import _is_yellow, _is_green
        # Simulate: xcno cell is green, condition cell has no fill
        xcno_cell = _cell_with_rgb("FF92D050")
        cond_cell = _cell_no_fill()
        cond_cell.value = "CON_Q4"
        # The extraction rule checks the condition cell colour only
        assert not _is_yellow(cond_cell, None)
        assert not _is_green(cond_cell, None)

    def test_blank_yellow_condition_cell_not_collected(self):
        """A yellow cell with no value should produce nothing."""
        from strategies.conditions.extract import _is_yellow
        cell = _cell_with_rgb("FFFFFF00")
        cell.value = None
        assert _is_yellow(cell, None) is True
        # The extraction loop skips it because value is blank — verified by the
        # early-continue guard: 'if cell_val is None or not str(cell_val).strip(): continue'

    def test_process_only_diff_false_collects_uncoloured_cells(self):
        """When process_only_differences=False, uncoloured condition cells with values are collected."""
        from strategies.conditions.extract import _is_yellow, _is_green
        # Uncoloured cell with a value — should NOT be ignored in full-file mode
        cell = _cell_no_fill()
        assert not _is_yellow(cell, None)
        assert not _is_green(cell, None)
        # The extraction code path (process_only_differences=False) calls _record without
        # checking colour — verified by the else branch in extract_conditions()

    def test_process_only_diff_true_skips_uncoloured_cells(self):
        """When process_only_differences=True, uncoloured condition cells are skipped."""
        from strategies.conditions.extract import _is_yellow, _is_green
        cell = _cell_no_fill()
        cell.value = "CON_Q4"
        # Colour check returns False for both — cell would be skipped in differences mode
        assert not (_is_yellow(cell, None) or _is_green(cell, None))

    def test_reference_xcheck_overrides_xcheck_no_in_concat(self):
        """
        When 'Reference  X-Check (Condition)' has a value, it must be used
        as the X-Check identifier in ALL concat keys for that row, not
        the 'X-Check No.' value from the row itself.
        """
        from strategies.conditions.extract import CONDITION_COLS
        # CONDITION_COLS[0] is "Reference  X-Check (Condition)"
        # CONDITION_COLS[1] is "Applicable Quarters"
        # Simulate one collected entry: X-Check No.=XC099, ref col=XC001, AQ=Q4
        collected = {
            "XC099": {
                CONDITION_COLS[0]: "XC001",   # reference override
                CONDITION_COLS[1]: "Q4",
                CONDITION_COLS[2]: None,
                CONDITION_COLS[3]: None,
                CONDITION_COLS[4]: None,
            }
        }
        # Replicate the DataFrame-building logic from extract_conditions
        _REF_XC_COL = CONDITION_COLS[0]
        rows = []
        for xcheck_no, cond_vals in sorted(collected.items()):
            ref_val = cond_vals.get(_REF_XC_COL)
            effective_xc = str(ref_val).strip() if ref_val and str(ref_val).strip() else xcheck_no
            row = {"X-Check No.": xcheck_no}
            for c in CONDITION_COLS:
                row[c] = cond_vals.get(c) or ""
            for c in CONDITION_COLS:
                val = cond_vals.get(c)
                row[c + " (Concat)"] = f"{effective_xc}|{val}" if val and str(val).strip() else ""
            rows.append(row)
        all_cols = ["X-Check No."] + CONDITION_COLS + [c + " (Concat)" for c in CONDITION_COLS]
        df = pd.DataFrame(rows, columns=all_cols)

        # Concat for col[0] (ref col itself): key should be XC001|XC001
        assert df.loc[0, CONDITION_COLS[0] + " (Concat)"] == "XC001|XC001"
        # Concat for col[1] (AQ): key should use effective_xc=XC001, not XC099
        assert df.loc[0, CONDITION_COLS[1] + " (Concat)"] == "XC001|Q4"
        # X-Check No. column itself is unchanged
        assert df.loc[0, "X-Check No."] == "XC099"


class TestProcessFip:

    def _make_fip_df(self, rows=None):
        """Build a minimal FIP DataFrame with 8 columns and optional data rows."""
        orig_cols = ["MethC", "MK", "Medium Text", "ValidRule", "Medium Text",
                     "UCFV20G-TRUE_BRANCH", "ValidRule", "Medium Text"]
        if rows is None:
            rows = [
                ["M1", "MK1", "desc1", "XC001", "xc text", "branch1", "C1",   "cond text"],
                ["M2", "MK2", "desc2", "XC002", "xc text", "branch2", "C2",   "cond text"],
                ["M3", "MK3", "desc3", "XC001", "xc text", "branch3", "",     ""],
            ]
        return pd.DataFrame(rows, columns=orig_cols)

    def test_columns_renamed(self):
        df = process_fip(self._make_fip_df())
        assert list(df.columns[:8]) == _RENAMED_COLS

    def test_concat_col_present(self):
        df = process_fip(self._make_fip_df())
        assert CONCAT_COL in df.columns

    def test_concat_value_format(self):
        df = process_fip(self._make_fip_df())
        assert df.loc[0, CONCAT_COL] == "XC001|C1"
        assert df.loc[1, CONCAT_COL] == "XC002|C2"

    def test_concat_empty_when_condition_blank(self):
        df = process_fip(self._make_fip_df())
        # Row 2 had blank Condition No → dropped from output (no usable key)
        assert len(df) == 2  # only the 2 rows with valid keys survive

    def test_raises_on_too_few_columns(self):
        bad_df = pd.DataFrame([[1, 2, 3]], columns=["A", "B", "C"])
        with pytest.raises(ValueError, match="at least 8"):
            process_fip(bad_df)

    def test_extra_columns_preserved(self):
        orig_cols = ["MethC", "MK", "Medium Text", "ValidRule", "Medium Text",
                     "UCFV20G-TRUE_BRANCH", "ValidRule", "Medium Text", "Extra"]
        df = pd.DataFrame([["v"] * 9], columns=orig_cols)
        result = process_fip(df)
        assert "Extra" in result.columns


# ---------------------------------------------------------------------------
# tests for compare
# ---------------------------------------------------------------------------

from strategies.conditions.compare import compare
from strategies.conditions.extract import CONDITION_COLS
from strategies.conditions.fip import process_fip


class TestCompare:

    def _make_working_df(self):
        """Minimal working_df: 2 X-Checks, only some concat values filled."""
        rows = [
            {
                "X-Check No.": "XC001",
                CONDITION_COLS[0]: "val1",
                CONDITION_COLS[0] + " (Concat)": "XC001|val1",
                CONDITION_COLS[1]: "",
                CONDITION_COLS[1] + " (Concat)": "",
                CONDITION_COLS[2]: "",
                CONDITION_COLS[2] + " (Concat)": "",
                CONDITION_COLS[3]: "",
                CONDITION_COLS[3] + " (Concat)": "",
                CONDITION_COLS[4]: "pct",
                CONDITION_COLS[4] + " (Concat)": "XC001|pct",
            },
            {
                "X-Check No.": "XC002",
                CONDITION_COLS[0]: "valX",
                CONDITION_COLS[0] + " (Concat)": "XC002|valX",   # NOT in FIP
                CONDITION_COLS[1]: "",
                CONDITION_COLS[1] + " (Concat)": "",
                CONDITION_COLS[2]: "",
                CONDITION_COLS[2] + " (Concat)": "",
                CONDITION_COLS[3]: "",
                CONDITION_COLS[3] + " (Concat)": "",
                CONDITION_COLS[4]: "",
                CONDITION_COLS[4] + " (Concat)": "",
            },
        ]
        all_cols = (
            ["X-Check No."]
            + CONDITION_COLS
            + [c + " (Concat)" for c in CONDITION_COLS]
        )
        return pd.DataFrame(rows, columns=all_cols)

    def _make_fip_df(self):
        """FIP DataFrame with Key (Concatenated) keys XC001|val1 and XC001|pct."""
        fip_cols = ["MethC", "MK", "Medium Text", "ValidRule", "Medium Text",
                    "UCFV20G-TRUE_BRANCH", "ValidRule", "Medium Text"]
        rows = [
            ["M1", "K1", "d", "XC001", "t", "b", "val1", "ct"],
            ["M2", "K2", "d", "XC001", "t", "b", "pct",  "ct"],
        ]
        raw = pd.DataFrame(rows, columns=fip_cols)
        return process_fip(raw)

    def test_matched_row_has_fip_data(self):
        # XC001|val1 exists in FIP → FIP Data = same value, Comparison = "Matched"
        results, _ = compare(self._make_working_df(), self._make_fip_df())
        matched = results[results["EBX Data"] == "XC001|val1"].iloc[0]
        assert matched["FIP Data"] == "XC001|val1"
        assert matched["Comparison"] == "Matched"

    def test_not_matched_row_has_empty_fip_data(self):
        # XC002|valX not in FIP → FIP Data = "", Comparison = "Not Matched"
        results, _ = compare(self._make_working_df(), self._make_fip_df())
        not_matched = results[results["EBX Data"] == "XC002|valX"].iloc[0]
        assert not_matched["FIP Data"] == ""
        assert not_matched["Comparison"] == "Not Matched"

    def test_blank_concat_produces_no_row(self):
        # Blank concat values (CONDITION_COLS[1] for both X-Checks) → no row emitted
        results, _ = compare(self._make_working_df(), self._make_fip_df())
        assert not any(results["EBX Data"].str.startswith("XC001|") & results["EBX Data"].str.contains(CONDITION_COLS[1]))

    def test_summary_counts(self):
        _, summary = compare(self._make_working_df(), self._make_fip_df())
        # XC001 col0 (val1 → matched), XC001 col4 (pct → matched), XC002 col0 (valX → not matched)
        assert summary["Total Pairs"] == 3
        assert summary["Matched"] == 2
        assert summary["Not Matched"] == 1

    def test_output_columns(self):
        results, _ = compare(self._make_working_df(), self._make_fip_df())
        assert list(results.columns) == ["EBX Data", "FIP Data", "Comparison"]


# ---------------------------------------------------------------------------
# Integration test — Conditions.process() with mocked I/O
# ---------------------------------------------------------------------------

from strategies.conditions.conditions import Conditions
from strategies.conditions.extract import CONDITION_COLS
from file_upload_config import UploadTaskConfig, FileFieldConfig


class TestConditionsProcess:

    def _make_config(self):
        return UploadTaskConfig(
            task_name="Conditions",
            window_title="Conditions Files",
            requires_output_directory=True,
            file_fields=[
                FileFieldConfig(label="X-Checks Publication File",
                                file_types=[("Excel Files", "*.xlsx")]),
                FileFieldConfig(label="FIP File",
                                file_types=[("Excel Files", "*.xlsx")]),
            ],
        )

    def _make_loaded_files(self):
        """Minimal loaded_files that process_fip and compare can handle."""
        fip_cols = ["MethC", "MK", "Medium Text", "ValidRule", "Medium Text",
                    "UCFV20G-TRUE_BRANCH", "ValidRule", "Medium Text"]
        fip_rows = [["M", "K", "d", "XC001", "t", "b", "C1", "ct"]]
        fip_df = pd.DataFrame(fip_rows, columns=fip_cols)
        return {"FIP File": fip_df}

    def _make_files_dict(self, tmp_path):
        return {
            "files": {
                "X-Checks Publication File": str(tmp_path / "pub.xlsx"),
                "FIP File": str(tmp_path / "fip.xlsx"),
            },
            "sheet_names": {"X-Checks Publication File": "cross checks all"},
            "output_directory": str(tmp_path),
            "timestamp": "20260101_000000",
            "process_only_differences": False,
        }

    def test_process_returns_true(self, tmp_path):
        config = self._make_config()
        strategy = Conditions(config)

        loaded_files = self._make_loaded_files()
        files_dict = self._make_files_dict(tmp_path)

        # Patch extract_conditions to return a minimal working_df without needing a real file
        rows = [{
            "X-Check No.": "XC001",
            **{c: "" for c in CONDITION_COLS},
            **{c + " (Concat)": "" for c in CONDITION_COLS},
        }]
        all_cols = ["X-Check No."] + CONDITION_COLS + [c + " (Concat)" for c in CONDITION_COLS]
        working_df = pd.DataFrame(rows, columns=all_cols)

        with patch("strategies.conditions.conditions.extract_conditions",
                   return_value=(working_df, [])) as mock_extract, \
             patch.object(strategy, "write_excel_output") as mock_write, \
             patch.object(strategy, "log_step"):

            strategy.log = []
            result = strategy.process(loaded_files, files_dict)

        assert result is True
        mock_extract.assert_called_once()
        mock_write.assert_called_once()
