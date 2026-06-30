"""
Accounting Principles comparator.

For each (V-code, X-Check No.) pair where FIP has a row, look up the
cross-checks-all letter via priority-ordered MethodBinding records:
  black-font bindings come before grey, and within each colour group the
  leftmost column wins. The first binding with a non-empty actual letter on
  cross-checks-all is the one we attribute the row to. Emit ONE output row
  per (X-Check, V-code).

Output columns: X-Check No. | Event | Expected | FIP | Actual | Method | Match
  Expected = severity declared on the winning binding (Warning|Error|Both)
  FIP      = the W/E from FIP for this (Method, X-Check)
  Actual   = the w/e from cross-checks-all for the winning event's column
"""
from __future__ import annotations

import pandas as pd

from .validation_methods import EventDefinition, MethodBinding


X_CHECK_NO_COL = "X-Check No."
FIP_KEY_COL    = "Key"
FIP_MT_COL     = "MT"


def _norm_letter(raw) -> str:
    """Strip + lowercase the cell value; treat 'nan'/'none' as empty."""
    if raw is None:
        return ""
    s = str(raw).strip().lower()
    return "" if s in ("", "nan", "none") else s


def _norm_event_name(name: str) -> str:
    """Squash spaces and hyphens, lowercase. Used to match validation-methods
    event names against cross-checks-all column headers despite punctuation
    differences (e.g. 'DE-GAAP RFD' vs 'DE GAAP RFD')."""
    return "".join(ch for ch in str(name).lower() if ch not in (" ", "-"))


def _build_event_to_column(cc_df: pd.DataFrame, events: list[str]) -> dict[str, str]:
    """
    {event_name: actual_cross_checks_all_column}, matching by punctuation-
    insensitive name. If multiple cc_df columns normalise to the same form
    (e.g. pandas appended '.1' to a duplicate header), the first wins. Events
    with no match are absent from the dict.
    """
    norm_to_col: dict[str, str] = {}
    for col in cc_df.columns:
        base = str(col).rsplit(".", 1)[0] if str(col).rsplit(".", 1)[-1].isdigit() else str(col)
        n = _norm_event_name(base)
        if n and n not in norm_to_col:
            norm_to_col[n] = col

    out: dict[str, str] = {}
    for ev in events:
        n = _norm_event_name(ev)
        if n in norm_to_col:
            out[ev] = norm_to_col[n]
    return out


# ---------------------------------------------------------------------------
# Backwards-compatible compare() — used by existing tests.
# Wraps the EventDefinition input into MethodBinding records, then defers to
# compare_with_bindings(). Kept so v0.5.* test suite still passes.
# ---------------------------------------------------------------------------

def compare(
    definitions: list[EventDefinition],
    cross_checks_df: pd.DataFrame,
    in_scope_x_checks: list[str],
    fip_methods_df: pd.DataFrame,
) -> list[dict]:
    bindings: list[MethodBinding] = []
    # Use sequential column indices so the test definitions retain their order.
    for idx, d in enumerate(definitions):
        col = idx + 1
        if d.severity == "Both":
            for m in (*d.methods_w, *d.methods_e):
                if m not in [b.method for b in bindings if b.event == d.event and b.severity == "Both"]:
                    bindings.append(MethodBinding(
                        method=m, event=d.event, severity="Both",
                        font="black", column=col,
                    ))
        else:
            for m in d.methods_w:
                bindings.append(MethodBinding(
                    method=m, event=d.event, severity="Warning",
                    font="black", column=col,
                ))
            for m in d.methods_e:
                bindings.append(MethodBinding(
                    method=m, event=d.event, severity="Error",
                    font="black", column=col,
                ))
    return compare_with_bindings(bindings, cross_checks_df, in_scope_x_checks, fip_methods_df)


# ---------------------------------------------------------------------------
# Real comparator: priority-ordered binding walk, ONE row per (X-Check, V-code).
# ---------------------------------------------------------------------------

def compare_with_bindings(
    bindings: list[MethodBinding],
    cross_checks_df: pd.DataFrame,
    in_scope_x_checks: list[str],
    fip_methods_df: pd.DataFrame,
) -> list[dict]:
    if X_CHECK_NO_COL not in cross_checks_df.columns:
        return []
    if FIP_KEY_COL not in fip_methods_df.columns or FIP_MT_COL not in fip_methods_df.columns:
        return []

    in_scope_set = set(in_scope_x_checks)

    # Group bindings by V-code, sort by (font priority, column).
    # font priority: black=0, grey=1.
    method_to_ordered_bindings: dict[str, list[MethodBinding]] = {}
    for b in bindings:
        method_to_ordered_bindings.setdefault(b.method, []).append(b)
    for m in method_to_ordered_bindings:
        method_to_ordered_bindings[m].sort(
            key=lambda b: (0 if b.font == "black" else 1, b.column)
        )

    # First-occurrence X-Check No. -> cc row
    xcheck_to_row: dict[str, pd.Series] = {}
    for _, row in cross_checks_df.iterrows():
        xc = str(row[X_CHECK_NO_COL]).strip()
        if xc in in_scope_set and xc not in xcheck_to_row:
            xcheck_to_row[xc] = row

    # event -> actual cross-checks-all column header
    all_events = sorted({b.event for b in bindings})
    event_to_col = _build_event_to_column(cross_checks_df, all_events)

    rows: list[dict] = []

    # Walk FIP. Each row gives a concrete (V-code, X-Check) + W/E.
    for _, fip_row in fip_methods_df.iterrows():
        key = str(fip_row[FIP_KEY_COL]).strip()
        if "|" not in key:
            continue
        method, xcheck = (s.strip() for s in key.split("|", 1))

        if xcheck not in xcheck_to_row:
            continue
        ordered = method_to_ordered_bindings.get(method)
        if not ordered:
            continue

        ccrow = xcheck_to_row[xcheck]
        fip_letter = _norm_letter(fip_row[FIP_MT_COL])

        # First binding in priority order whose cc column has a non-empty actual wins.
        winning: MethodBinding | None = None
        winning_actual = ""
        for b in ordered:
            cc_col = event_to_col.get(b.event)
            if cc_col is None:
                continue
            actual = _norm_letter(ccrow.get(cc_col))
            if actual:
                winning = b
                winning_actual = actual
                break
        if winning is None:
            continue

        verdict = "Match" if fip_letter == winning_actual else "MisMatch"
        rows.append({
            "X-Check No.": xcheck,
            "Event":       winning.event,
            "Expected":    winning.severity,
            "FIP":         fip_letter,
            "Actual":      winning_actual,
            "Method":      method,
            "Match":       verdict,
        })

    rows.sort(key=lambda r: (r["X-Check No."], r["Event"], r["Method"]))
    return rows
