# Included RUs / Excluded RUs — Analysis & Current Status

**Date:** 2026-05-19  
**Branch:** `v0.3-X-Checks`  
**Author:** Analysis generated from `EPM X-Checks file with 3345 Data rows on sheet cross checks all.xlsx`

---

## 1. What These Columns Are

The EBX "cross checks all" sheet contains two columns that define **scope** — i.e. which Reporting Units (RUs) an X-Check is applied to when run in the consolidation system (FIP):

| Column | Meaning |
|---|---|
| **Included RUs** | The X-Check runs **only** for the listed RUs. All other RUs are excluded by default. |
| **Excluded RUs** | The X-Check runs for **all** RUs **except** the listed ones. |

RU codes follow the pattern `CON_XXXXX`, referring to specific legal entities or entity groups within the Zurich Insurance consolidation hierarchy (e.g. `CON_LFLOB` = Life FLoB entities, `CON_DE` = Germany, `CON_ZIC` = Zurich Insurance Company).

---

## 2. Scale of the Data

### Included RUs

| Metric | Value |
|---|---|
| Rows with a value | 417 |
| X-Checks affected | **93** |
| Unique RU codes | 13 |

**RU codes and the number of X-Checks each applies to:**

| RU Code | X-Checks |
|---|---|
| CON_LFLOB | 45 |
| CON_ZIC | 17 |
| CON_DE | 12 |
| CON_RR | 7 |
| CON_ARG2 | 2 |
| CON_201000 | 2 |
| CON_CH2 | 2 |
| CON_123SII | 1 |
| CON_GI17 | 1 |
| CON_PC&BBA | 1 |
| CON_ZICZR | 1 |
| CON_ZIP | 1 |
| CON_ZLIC | 1 |

### Excluded RUs

| Metric | Value |
|---|---|
| Rows with a value | 171 |
| X-Checks affected | **30** |
| Unique RU codes | 25 |

**Selected RU codes:**

| RU Code | X-Checks | RU Code | X-Checks |
|---|---|---|---|
| CON_51012 | 2 | CON_BERMUD | 1 |
| CON_ZIP | 2 | CON_CH | 2 |
| CON_ARG2 | 2 | CON_DE | 1 |
| CON_910208 | 2 | CON_DIR | 1 |
| CON_400069 | 1 | CON_FA | 1 |
| CON_920011 | 1 | CON_RR | 1 |
| CON_30054 | 1 | CON_ReG&L | 1 |
| CON_10000 | 1 | CON_SISTER | 1 |
| CON_SWISS | 1 | CON_Takaf | 1 |

### Combined

- **122 X-Checks** have either an Included or Excluded RU (or both).
- **1 X-Check** has both simultaneously: `L019_00` (Included: `CON_ZIC`, Excluded: `CON_SISTER`).

---

## 3. Key Finding: RU Filtering Does NOT Appear in the FIP Formula

This is the most important architectural point. Looking at the FIP text file, **neither Included RUs nor Excluded RUs appear anywhere in the formula string or variable selection sections** of the corresponding X-Check blocks.

**Example — A313_00 (Included: CON_123SII):**
The FIP formula is:
```
ABS( VAL_YTD( 16461 ) ) >= CONST( 1.000 , 'USD' , 'E' )
```
The selection shows `Segment @28@  *` — the wildcard `*`, meaning all segments. There is no reference to `CON_123SII` anywhere in the FIP block.

**Example — A142_00 (Excluded: CON_400069):**
The FIP formula is:
```
ABS( VAL_YTD( 16903ff ToM 553 ) + VAL_YTD( 16901ff ToM 554 ) + ... ) <= CONST( 5 , 'USD' , 'E' )
```
Again, `CON_400069` does not appear in the formula or selections. The `Segment @28@  *` wildcard is present as usual.

**Conclusion:** The Included/Excluded RU filtering is enforced at the **execution layer** of the consolidation system (i.e. when FIP runs the X-Check, it applies the scope filter before executing the formula). The FIP text export only captures the formula definition — not the execution scope. The EBX publication file is the single source of truth for RU scope.

---

## 4. Impact on the Current Comparison

Of the 122 X-Checks with RU filtering defined:

| Formula Match Result | Count |
|---|---|
| Match | 89 (73%) |
| MisMatch | 33 (27%) |
| Not Found | 0 |

The 33 mismatches in this group are **not caused by** the Included/Excluded RU columns — they are caused by other factors (e.g. version spanning, GAAP prefix differences, variable ordering). The RU filtering columns have no bearing on whether the EBX and FIP formula strings match.

---

## 5. Current Implementation Status

**The Included RUs and Excluded RUs columns are read from the EBX sheet but are not used anywhere in the application code.**

Specifically:
- `ebx_extraction.py` — does not read these columns
- `fip_extraction.py` — these values do not appear in the FIP text, so nothing to extract
- `compare.py` — no comparison or flagging of RU scope
- `x_checks.py` — no output column for RU scope
- Output Excel — no column showing which RUs the X-Check applies to

---

## 6. What Could Be Implemented

Since the RU filter does not affect the formula comparison, implementing it would be about **enriching the output** rather than changing the comparison logic. Two distinct capabilities are possible:

### 6a. Add RU Scope as Informational Output Columns

Add two columns to the output sheet:

| Column | Source | Purpose |
|---|---|---|
| `Included RUs` | EBX `Included RUs` column (de-duplicated per X-Check) | Shows which RUs this check is restricted to |
| `Excluded RUs` | EBX `Excluded RUs` column (de-duplicated per X-Check) | Shows which RUs are excluded from this check |

This is purely informational — the user can see at a glance that a mismatch only affects certain entities, or that a "pass" only applies to a limited scope.

**Effort:** Low — read both columns during `extract_ebx()`, pass through to comparison output, add to the DataFrame.

### 6b. Filter the Comparison by RU

If the user is running the X-Checks tool for a **specific RU** (e.g. a single legal entity close), they could optionally specify the RU, and the tool would:
- Exclude X-Checks where that RU is in the `Excluded RUs` list
- For X-Checks with `Included RUs` set: only include them if that RU is in the list

This would reduce noise — mismatches for X-Checks that don't apply to the current RU wouldn't appear.

**Effort:** Medium — requires a new optional input (the target RU code), a filter step after extraction, and UI support.

---

## 7. Recommendation

The RU scope columns are documentation-quality metadata that belongs in the output for traceability. **Option 6a (informational output columns)** is the most practical next step:

1. It requires no changes to comparison logic.
2. It gives users context when reviewing mismatches — knowing a check is `CON_LFLOB`-only immediately scopes the impact.
3. It can be implemented in a single session and validated with the existing test infrastructure.

Option 6b (RU filtering) is a useful future enhancement but requires more design input on how the "target RU" would be specified and whether it should apply universally or per-run.

---

## 8. Summary Table

| Dimension | Detail |
|---|---|
| Affected X-Checks | 122 (93 Included only, 30 Excluded only, 1 both) |
| Total in golden pair1 | 122 |
| Mismatch rate in this group | 27% (33 of 122) |
| Root cause of mismatches | Other factors (not RU scope) |
| Appears in FIP text | No — RU scope is not in the formula export |
| Currently implemented | No |
| Recommended next step | Add Included/Excluded RUs as informational output columns |
| Effort for next step | Low |
