"""
Parses the Validation Methods xlsx (sheet 'Validation Methods').

Layout (Q1 2026 file):
  - Row 1, cols C..BM: Validation Event names (the 'Accounting Principles').
  - Rows 4-6 hold the current period block ('from 2023'):
      * row 4 = Warning row
      * rows 5, 6 = Error rows
  - A cell merged across rows 4-5 (or 4-6) for a single column means that
    event accepts EITHER 'w' or 'e' on a cross-checks-all row, i.e. severity = Both.
  - Method cells contain newline-separated entries like
        "V900A - Part A (2023 onwards)"
    The leading code before " - " is the method ID we keep.
  - "-" or empty cells mean 'no entry' for that severity.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Optional

import openpyxl
from openpyxl.cell.cell import MergedCell


SHEET_NAME = "Validation Methods"
EVENT_ROW = 1
WARNING_ROW = 4
ERROR_ROWS = (5, 6)
EVENT_COL_START = 3   # column C


@dataclass
class EventDefinition:
    """
    Severity declared for a Validation Event in the current-period block,
    plus the method codes the user should expect to see for each severity.

    For severity == "Both", the methods on the merged cell are returned via
    methods_w (and copied into methods_e) so a downstream cross-checks-all 'w'
    or 'e' can both look up applicable methods.
    """
    event: str
    severity: str            # "Warning" | "Error" | "Both"
    methods_w: list[str] = field(default_factory=list)
    methods_e: list[str] = field(default_factory=list)


_METHOD_LINE_RE = re.compile(r"^\s*([A-Za-z0-9]+)")


def _extract_method_codes(cell_value) -> list[str]:
    """
    Splits a cell's content on newlines and pulls the leading code from each line.
    'V900A - Part A...' -> 'V900A'
    Empty / dash / whitespace-only lines are dropped.
    """
    if cell_value is None:
        return []
    text = str(cell_value).strip()
    if not text or text == "-":
        return []
    codes: list[str] = []
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line or line == "-":
            continue
        m = _METHOD_LINE_RE.match(line)
        if m:
            codes.append(m.group(1))
    return codes


def _is_blank(cell_value) -> bool:
    """Empty, whitespace-only, or '-' cells count as blank."""
    if cell_value is None:
        return True
    s = str(cell_value).strip()
    return s == "" or s == "-"


def _column_is_in_warning_merge_with_error(ws, col: int) -> bool:
    """True if a single merged region in this column spans the Warning row AND
    at least one Error row (i.e. the cell's content covers Warning+Error =
    severity 'Both' for this column)."""
    for rng in ws.merged_cells.ranges:
        if rng.min_col <= col <= rng.max_col:
            covers_warning = rng.min_row <= WARNING_ROW <= rng.max_row
            covers_error   = any(rng.min_row <= er <= rng.max_row for er in ERROR_ROWS)
            if covers_warning and covers_error:
                return True
    return False


def _read_cell_through_merges(ws, row: int, col: int):
    """Returns the value at (row,col), following merged-cell semantics so
    every cell in a merged range yields the merged value (not None)."""
    cell = ws.cell(row=row, column=col)
    if not isinstance(cell, MergedCell):
        return cell.value
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            return ws.cell(row=rng.min_row, column=rng.min_col).value
    return None


def parse_validation_methods(filepath: str, subset: list[str]) -> list[EventDefinition]:
    """
    Reads the Validation Methods workbook and returns one EventDefinition per
    event in `subset` that has any non-empty content in rows 4-6. Events with
    nothing recorded are silently dropped (per the spec's "do nothing" rule).

    `subset` is matched case-sensitively against row-1 event names; events not
    found in the file are skipped (the caller can warn separately).
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    if SHEET_NAME not in wb.sheetnames:
        wb.close()
        raise ValueError(f"Sheet '{SHEET_NAME}' not found in {filepath}")
    ws = wb[SHEET_NAME]

    # Build event-name -> column index map (only for columns present in subset)
    subset_set = set(subset)
    col_for_event: dict[str, int] = {}
    for col in range(EVENT_COL_START, ws.max_column + 1):
        v = ws.cell(row=EVENT_ROW, column=col).value
        if v is None:
            continue
        name = str(v).strip()
        if name in subset_set and name not in col_for_event:
            col_for_event[name] = col

    results: list[EventDefinition] = []

    for event in subset:
        col = col_for_event.get(event)
        if col is None:
            continue

        warning_val = _read_cell_through_merges(ws, WARNING_ROW, col)
        error_vals  = [_read_cell_through_merges(ws, r, col) for r in ERROR_ROWS]

        warning_methods = _extract_method_codes(warning_val)
        error_methods: list[str] = []
        for v in error_vals:
            error_methods.extend(_extract_method_codes(v))

        is_both = _column_is_in_warning_merge_with_error(ws, col)

        if is_both:
            # The merged-cell content was written into warning_methods AND
            # the error rows already (since _read_cell_through_merges follows
            # the merge). Deduplicate while preserving order.
            merged_methods = list(dict.fromkeys(warning_methods + error_methods))
            if not merged_methods:
                continue   # Both declared but no method content -> ignore
            results.append(EventDefinition(
                event=event,
                severity="Both",
                methods_w=merged_methods,
                methods_e=merged_methods,
            ))
            continue

        # Independent Warning + Error cells
        if warning_methods and error_methods:
            # Per spec: emit two definitions (Warning + Error) so the comparator
            # can match each independently against the cross-checks-all letter.
            results.append(EventDefinition(
                event=event, severity="Warning",
                methods_w=warning_methods, methods_e=[],
            ))
            results.append(EventDefinition(
                event=event, severity="Error",
                methods_w=[], methods_e=error_methods,
            ))
        elif warning_methods:
            results.append(EventDefinition(
                event=event, severity="Warning",
                methods_w=warning_methods, methods_e=[],
            ))
        elif error_methods:
            results.append(EventDefinition(
                event=event, severity="Error",
                methods_w=[], methods_e=error_methods,
            ))
        # else: nothing recorded for this event -> drop

    wb.close()
    return results


def list_all_event_names(filepath: str) -> list[str]:
    """Returns every Validation Event name in row 1 (cols C..BM), in column order.
    Used by the form's multi-select to build the checkbox list."""
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    if SHEET_NAME not in wb.sheetnames:
        wb.close()
        return []
    ws = wb[SHEET_NAME]
    names: list[str] = []
    for col in range(EVENT_COL_START, ws.max_column + 1):
        v = ws.cell(row=EVENT_ROW, column=col).value
        if v is not None:
            n = str(v).strip()
            if n:
                names.append(n)
    wb.close()
    return names
