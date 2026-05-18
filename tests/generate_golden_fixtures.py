"""
Run this script once (from the project root) to generate the golden CSV fixtures
used by tests/test_integration.py.

    python tests/generate_golden_fixtures.py

Re-run only when you intentionally change pipeline behaviour and want to
advance the baseline. Never run it to "fix" a failing test — investigate first.
"""
import sys
from pathlib import Path

ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))

import pandas as pd
from strategies.x_checks.ebx_extraction import extract_ebx
from strategies.x_checks.fip_extraction import extract_fip
from strategies.x_checks.compare import compare

COLS = ["X-Check Number", "Formula Match", "EBX Formula", "FIP Formula",
        "Variables Match", "EBX Variables", "FIP Variables"]

PAIRS = [
    {
        "name": "pair1",
        "ebx": ROOT / "test_data" / "20251205 EPM X-Checks - Original - Copy.xlsx",
        "ebx_sheet": "cross checks all",
        "fip": ROOT / "test_data" / "20251205 FIP X-Checks - Original.txt",
    },
    {
        "name": "pair2",
        "ebx": ROOT / "test_data" / "20260313 Cross Checks All.xlsx",
        "ebx_sheet": "cross checks all",
        "fip": ROOT / "test_data" / "20260318 FIP X-Checks.txt",
    },
]

FIXTURES_DIR = Path(__file__).parent / "fixtures"
FIXTURES_DIR.mkdir(exist_ok=True)

for pair in PAIRS:
    print(f"Generating {pair['name']}...")
    ebx_df = pd.read_excel(pair["ebx"], sheet_name=pair["ebx_sheet"])
    fip_text = pair["fip"].read_text(encoding="utf-8", errors="replace")

    x_check_list = sorted(set(
        str(x) for x in ebx_df["X-Check No."].tolist()
        if str(x) not in ("nan", "", "NaN", "None")
    ))

    ebx_results = extract_ebx(ebx_df)
    fip_results = extract_fip(fip_text, x_check_list)
    rows = compare(ebx_results, fip_results)

    df = pd.DataFrame(rows)[COLS].sort_values("X-Check Number").reset_index(drop=True)
    out = FIXTURES_DIR / f"golden_{pair['name']}.csv"
    df.to_csv(out, index=False)
    print(f"  Written {len(df)} rows to {out}")

print("Done.")
