"""
Accounting Principles comparator.

The FIP Methods Rules and Condition sheet (col A 'Key' = '<Method>|<X-Check>',
col J 'MT' = 'W' or 'E') is the source of truth for which (Method, X-Check)
combinations actually exist. We emit one output row per FIP entry whose
Method belongs to one of our subset events and whose X-Check is in scope.

Match rule:
  Compare FIP letter (W/E, case-insensitive) to cross-checks-all letter
  (w/e, case-insensitive) at the same X-Check row, in the column whose
  header equals the event name.

Output columns: X-Check No. | Event | Expected | FIP | Actual | Method | Match
  Expected = severity from validation_methods (Warning|Error|Both)
  FIP      = the W/E from FIP for this (Method, X-Check)
  Actual   = the w/e from cross-checks-all for this (X-Check, Event)
"""
from __future__ import annotations

import pandas as pd

from .validation_methods import EventDefinition


X_CHECK_NO_COL = "X-Check No."
FIP_KEY_COL    = "Key"
FIP_MT_COL     = "MT"
FIP_VALIDRULE  = "ValidRule"   # contains the X-Check No. (e.g. 'A047_00')
FIP_MK_COL     = "MK"          # method code (e.g. 'V900W')


def _norm_letter(raw) -> str:
    """Strip + lowercase the cell value; treat 'nan'/'none' as empty."""
    if raw is None:
        return ""
    s = str(raw).strip().lower()
    return "" if s in ("", "nan", "none") else s


def compare(
    definitions: list[EventDefinition],
    cross_checks_df: pd.DataFrame,
    in_scope_x_checks: list[str],
    fip_methods_df: pd.DataFrame,
) -> list[dict]:
    """
    Returns rows ready to be turned into an output DataFrame, in the order:
      X-Check No. > Event > Method.
    """
    if X_CHECK_NO_COL not in cross_checks_df.columns:
        return []
    if FIP_KEY_COL not in fip_methods_df.columns or FIP_MT_COL not in fip_methods_df.columns:
        return []

    in_scope_set = set(in_scope_x_checks)

    # Build: method -> [EventDefinition] (a method may appear under multiple
    # events / definitions, but the SAME definition mustn't double-register
    # because it has the method in both methods_w and methods_e).
    method_to_definitions: dict[str, list[EventDefinition]] = {}
    for d in definitions:
        seen_methods: set = set()
        for m in (*d.methods_w, *d.methods_e):
            if m in seen_methods:
                continue
            seen_methods.add(m)
            method_to_definitions.setdefault(m, []).append(d)

    # Build: X-Check No. -> cross-checks-all row (first occurrence wins)
    xcheck_to_row: dict[str, pd.Series] = {}
    for _, row in cross_checks_df.iterrows():
        xc = str(row[X_CHECK_NO_COL]).strip()
        if xc in in_scope_set and xc not in xcheck_to_row:
            xcheck_to_row[xc] = row

    rows: list[dict] = []

    # Walk FIP. Each row gives us a concrete (method, x_check) pair AND its W/E.
    for _, fip_row in fip_methods_df.iterrows():
        key = str(fip_row[FIP_KEY_COL]).strip()
        if "|" not in key:
            continue
        method, xcheck = key.split("|", 1)
        method = method.strip()
        xcheck = xcheck.strip()

        if xcheck not in xcheck_to_row:
            continue
        defs = method_to_definitions.get(method)
        if not defs:
            continue

        fip_letter = _norm_letter(fip_row[FIP_MT_COL])    # 'w' or 'e'
        ccrow = xcheck_to_row[xcheck]

        for d in defs:
            if d.event not in cross_checks_df.columns:
                continue
            actual = _norm_letter(ccrow.get(d.event))
            if actual == "":
                continue
            verdict = "Match" if fip_letter == actual else "MisMatch"
            rows.append({
                "X-Check No.": xcheck,
                "Event":       d.event,
                "Expected":    d.severity,
                "FIP":         fip_letter,
                "Actual":      actual,
                "Method":      method,
                "Match":       verdict,
            })

    rows.sort(key=lambda r: (r["X-Check No."], r["Event"], r["Method"]))
    return rows
