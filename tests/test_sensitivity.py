"""
Unit tests for the sensitivity-label module. Excel COM is mocked because
the test runner doesn't have a live Excel; runtime correctness is verified
by the smoke test in v0.5.12 (it round-trips a real label and reads it back).
"""
import os
import pytest

from strategies.sensitivity import (
    ExcelLabeler,
    SITE_ID,
    label_for,
    _LABELS,
)


def test_label_for_known_levels():
    for name in _LABELS:
        label_id, label_friendly = label_for(name)
        assert label_id and label_friendly == name


def test_label_for_unknown_raises():
    with pytest.raises(KeyError):
        label_for("NotALevel")


def test_internal_use_only_id_matches_vba_constant():
    label_id, _ = label_for("Internal_Use_Only")
    assert label_id == "9108d454-5c13-4905-93be-12ec8059c842"


def test_site_id_matches_vba_constant():
    assert SITE_ID == "473672ba-cd07-4371-a2ae-788b4c61840e"


def test_labeler_returns_failure_for_missing_file(tmp_path):
    labeler = ExcelLabeler()
    ok, msg = labeler.label_file(str(tmp_path / "does_not_exist.xlsx"),
                                 "Internal_Use_Only")
    assert ok is False
    assert "not found" in msg.lower()
    labeler.close()


def test_labeler_returns_failure_for_unknown_level(tmp_path):
    # Create a real empty xlsx so the missing-file path is excluded.
    from openpyxl import Workbook
    p = tmp_path / "x.xlsx"
    Workbook().save(p)
    labeler = ExcelLabeler()
    ok, msg = labeler.label_file(str(p), "NotARealLevel")
    assert ok is False
    assert "Unknown sensitivity level" in msg
    labeler.close()
