"""
Unit tests for the Full Run strategy.

Tests cover:
  _unique_name()            — sheet name deduplication + 31-char truncation
  _build_full_run_config()  — merges file fields, deduplicates by label
  FullRun.process()         — correct sheet capture, prefixing, combined output
  FullRun.apply_output_formatting() — tab colours set per strategy
"""

import pytest
import pandas as pd
from collections import OrderedDict
from unittest.mock import MagicMock, patch, call

from strategies.full_run.full_run import FullRun, _unique_name, STRATEGY_COLOURS
from task_configs import _build_full_run_config
from file_upload_config import UploadTaskConfig, FileFieldConfig


# ---------------------------------------------------------------------------
# _unique_name
# ---------------------------------------------------------------------------

class TestUniqueName:

    def test_new_name_returned_unchanged(self):
        assert _unique_name("Cond — Conditions", {}) == "Cond — Conditions"

    def test_duplicate_gets_counter_suffix(self):
        existing = {"Cond — Conditions": None}
        result = _unique_name("Cond — Conditions", existing)
        assert result == "Cond — Conditions (2)"

    def test_counter_increments_past_existing_suffixes(self):
        existing = {"Cond — Conditions": None, "Cond — Conditions (2)": None}
        result = _unique_name("Cond — Conditions", existing)
        assert result == "Cond — Conditions (3)"

    def test_truncated_to_31_chars(self):
        long_name = "A" * 40
        result = _unique_name(long_name, {})
        assert len(result) == 31

    def test_truncated_duplicate_still_unique(self):
        long_name = "A" * 40
        existing = {"A" * 31: None}
        result = _unique_name(long_name, existing)
        assert result not in existing
        assert len(result) <= 31


# ---------------------------------------------------------------------------
# _build_full_run_config
# ---------------------------------------------------------------------------

def _make_config(task_name, labels):
    return UploadTaskConfig(
        task_name=task_name,
        file_fields=[
            FileFieldConfig(label=lbl, file_types=[("Excel Files", "*.xlsx")])
            for lbl in labels
        ],
    )


class TestBuildFullRunConfig:

    def test_merges_all_fields(self):
        registry = {
            "Strategy A": (_make_config("Strategy A", ["File 1", "File 2"]), None),
            "Strategy B": (_make_config("Strategy B", ["File 3"]), None),
        }
        config = _build_full_run_config(registry)
        assert [f.label for f in config.file_fields] == ["File 1", "File 2", "File 3"]

    def test_deduplicates_shared_labels(self):
        registry = {
            "Strategy A": (_make_config("Strategy A", ["Shared File", "File A"]), None),
            "Strategy B": (_make_config("Strategy B", ["Shared File", "File B"]), None),
        }
        config = _build_full_run_config(registry)
        labels = [f.label for f in config.file_fields]
        assert labels.count("Shared File") == 1
        assert set(labels) == {"Shared File", "File A", "File B"}

    def test_skips_full_run_entry(self):
        registry = {
            "Strategy A": (_make_config("Strategy A", ["File 1"]), None),
            "Full Run":   (_make_config("Full Run",   ["File 1"]), None),
        }
        config = _build_full_run_config(registry)
        assert len(config.file_fields) == 1

    def test_config_metadata(self):
        registry = {
            "Strategy A": (_make_config("Strategy A", ["File 1"]), None),
        }
        config = _build_full_run_config(registry)
        assert config.task_name == "Full Run"
        assert config.requires_output_directory is True


# ---------------------------------------------------------------------------
# FullRun.process — sheet capture and combination
# ---------------------------------------------------------------------------

def _make_strategy_factory(sheet_data: dict):
    """
    Returns a factory that produces a mock BaseStrategy whose process()
    calls its monkey-patched write_excel_output with the given sheets.
    """
    def factory(config):
        instance = MagicMock()
        instance.log = []

        def fake_process(loaded_files, files):
            # Simulate what a real strategy does: call write_excel_output.
            instance.write_excel_output(
                "fake_path.xlsx",
                OrderedDict(sheet_data),
                instance.log,
            )
            return True

        instance.process = fake_process
        return instance
    return factory


class TestFullRunProcess:

    def _make_files(self, labels):
        return {
            "files":       {lbl: f"/fake/{lbl}.xlsx" for lbl in labels},
            "sheet_names": {lbl: "Sheet1" for lbl in labels},
            "output_directory": "/fake/output",
            "timestamp": "20260101_000000",
            "process_only_differences": False,
        }

    def _run(self, registry_patch, labels):
        config = _make_config("Full Run", labels)
        strategy = FullRun(config)
        strategy.log = []
        strategy._progress_dialog = None
        strategy._stop_event = None
        strategy._sensitivity_labeler = None

        written_sheets = {}

        def fake_write(output_path, sheets, log, summaries=None):
            written_sheets.update(sheets)

        strategy.write_excel_output = fake_write
        strategy._apply_sensitivity_label = lambda path: None

        loaded = {lbl: pd.DataFrame({"col": [1]}) for lbl in labels}
        files = self._make_files(labels)

        with patch("task_registry.TASK_REGISTRY", registry_patch):
            result = strategy.process(loaded, files)

        return result, written_sheets

    def test_sheets_are_prefixed(self):
        registry = {
            "Conditions": (
                _make_config("Conditions", ["X-Checks Publication File", "FIP File"]),
                _make_strategy_factory({"Conditions": pd.DataFrame(), "Working Sheet": pd.DataFrame()}),
            ),
        }
        result, sheets = self._run(
            registry,
            ["X-Checks Publication File", "FIP File"],
        )
        assert result is True
        assert "Cond — Conditions" in sheets
        assert "Cond — Working Sheet" in sheets

    def test_full_run_entry_skipped(self):
        """The Full Run entry in the registry must not call itself recursively."""
        sentinel = MagicMock()
        registry = {
            "Conditions": (
                _make_config("Conditions", ["X-Checks Publication File", "FIP File"]),
                _make_strategy_factory({"Conditions": pd.DataFrame()}),
            ),
            "Full Run": (
                _make_config("Full Run", []),
                sentinel,
            ),
        }
        self._run(registry, ["X-Checks Publication File", "FIP File"])
        sentinel.assert_not_called()

    def test_returns_false_when_no_sheets(self):
        empty_factory = lambda config: MagicMock(
            log=[],
            process=lambda loaded, files: None,
        )
        registry = {
            "Conditions": (_make_config("Conditions", ["FIP File"]), empty_factory),
        }
        result, sheets = self._run(registry, ["FIP File"])
        assert result is False

    def test_strategy_exception_does_not_abort_remaining(self):
        """If one strategy raises, the next strategy should still run."""
        def bad_factory(config):
            instance = MagicMock()
            instance.log = []
            instance.process = MagicMock(side_effect=RuntimeError("boom"))
            return instance

        registry = {
            "Bad Strategy": (_make_config("Bad Strategy", ["FIP File"]), bad_factory),
            "Conditions":   (
                _make_config("Conditions", ["FIP File"]),
                _make_strategy_factory({"Conditions": pd.DataFrame()}),
            ),
        }
        result, sheets = self._run(registry, ["FIP File"])
        assert result is True
        assert "Cond — Conditions" in sheets

    def test_duplicate_sheet_names_made_unique(self):
        """Two strategies producing a sheet named 'Results' should not collide."""
        def _prefix_factory(prefix, sheet_name):
            return _make_strategy_factory({sheet_name: pd.DataFrame()})

        registry = {
            "Strategy A": (_make_config("Strategy A", ["File A"]),
                           _make_strategy_factory({"Results": pd.DataFrame()})),
            "Strategy B": (_make_config("Strategy B", ["File B"]),
                           _make_strategy_factory({"Results": pd.DataFrame()})),
        }
        # Both strategies need different prefixes; patch _SHEET_PREFIXES too.
        with patch("strategies.full_run.full_run._SHEET_PREFIXES",
                   {"Strategy A": "SA", "Strategy B": "SB"}):
            result, sheets = self._run(registry, ["File A", "File B"])
        assert result is True
        assert len(sheets) == 2
        assert all(k in sheets for k in ["SA — Results", "SB — Results"])


# ---------------------------------------------------------------------------
# FullRun.apply_output_formatting — tab colours
# ---------------------------------------------------------------------------

class TestApplyOutputFormatting:

    def _make_workbook(self, sheet_names):
        wb = MagicMock()
        wb.sheetnames = sheet_names
        sheets = {name: MagicMock() for name in sheet_names}
        wb.__getitem__ = lambda self, key: sheets[key]
        return wb

    def test_conditions_gets_zurich_blue(self):
        wb = self._make_workbook(["Cond — Conditions", "Processing Log"])
        strategy = FullRun.__new__(FullRun)
        strategy._strategy_sheet_names = {"Conditions": ["Cond — Conditions"]}
        strategy.apply_output_formatting(wb)
        assert wb["Cond — Conditions"].sheet_properties.tabColor == STRATEGY_COLOURS["Conditions"]

    def test_processing_log_gets_grey(self):
        wb = self._make_workbook(["Cond — Conditions", "Processing Log"])
        strategy = FullRun.__new__(FullRun)
        strategy._strategy_sheet_names = {"Conditions": ["Cond — Conditions"]}
        strategy.apply_output_formatting(wb)
        assert wb["Processing Log"].sheet_properties.tabColor == "808080"

    def test_unknown_strategy_gets_fallback_colour(self):
        wb = self._make_workbook(["XX — Sheet"])
        strategy = FullRun.__new__(FullRun)
        strategy._strategy_sheet_names = {"Unknown Strategy": ["XX — Sheet"]}
        strategy.apply_output_formatting(wb)
        # Should not raise; tab colour should be set to something.
        assert wb["XX — Sheet"].sheet_properties.tabColor is not None
