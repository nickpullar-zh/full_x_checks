"""
Integration parity tests.

Runs the core pipeline (extract_ebx -> extract_fip -> compare) against real
test data and asserts the output is identical to the committed golden fixtures.

If this test fails after a code change, either:
  - You introduced a regression (investigate and fix)
  - You intentionally improved the output (re-run generate_golden_fixtures.py
    to advance the baseline, then commit the updated CSV)
"""
import pandas as pd
import pytest
from pathlib import Path

ROOT = Path(__file__).parent.parent
FIXTURES = Path(__file__).parent / "fixtures"

COMPARE_COLS = [
    "Formula Match", "EBX Formula", "FIP Formula",
    "Variables Match", "EBX Variables", "FIP Variables",
]

PAIRS = [
    pytest.param(
        ROOT / "test_data" / "20251205 EPM X-Checks - Original - Copy.xlsx",
        "cross checks all",
        ROOT / "test_data" / "20251205 FIP X-Checks - Original.txt",
        FIXTURES / "golden_pair1.csv",
        id="pair1",
    ),
    pytest.param(
        ROOT / "test_data" / "20260313 Cross Checks All.xlsx",
        "cross checks all",
        ROOT / "test_data" / "20260318 FIP X-Checks.txt",
        FIXTURES / "golden_pair2.csv",
        id="pair2",
    ),
]


@pytest.mark.parametrize("ebx_path,ebx_sheet,fip_path,golden_path", PAIRS)
def test_core_pipeline_matches_golden(ebx_path, ebx_sheet, fip_path, golden_path):
    from strategies.x_checks.ebx_extraction import extract_ebx
    from strategies.x_checks.fip_extraction import extract_fip
    from strategies.x_checks.compare import compare

    ebx_df = pd.read_excel(ebx_path, sheet_name=ebx_sheet)
    fip_text = Path(fip_path).read_text(encoding="utf-8", errors="replace")

    x_check_list = sorted(set(
        str(x) for x in ebx_df["X-Check No."].tolist()
        if str(x) not in ("nan", "", "NaN", "None")
    ))

    ebx_results = extract_ebx(ebx_df)
    fip_results = extract_fip(fip_text, x_check_list)
    rows = compare(ebx_results, fip_results)

    actual = (
        pd.DataFrame(rows)
        .sort_values("X-Check Number")
        .reset_index(drop=True)
        .set_index("X-Check Number")
        .astype(str)
    )
    golden = (
        pd.read_csv(golden_path)
        .sort_values("X-Check Number")
        .reset_index(drop=True)
        .set_index("X-Check Number")
        .astype(str)
    )

    differences = []
    all_keys = sorted(set(actual.index) | set(golden.index))

    for xcheck in all_keys:
        if xcheck not in golden.index:
            differences.append(f"{xcheck}: only in new output")
            continue
        if xcheck not in actual.index:
            differences.append(f"{xcheck}: missing from new output (was in golden)")
            continue
        for col in COMPARE_COLS:
            a = actual.loc[xcheck, col] if col in actual.columns else "N/A"
            g = golden.loc[xcheck, col] if col in golden.columns else "N/A"
            if str(a).strip() != str(g).strip():
                differences.append(f"{xcheck} [{col}]: got {a!r}, expected {g!r}")

    assert not differences, (
        f"{len(differences)} difference(s) found:\n" + "\n".join(differences[:30])
    )
