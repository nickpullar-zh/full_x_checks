# Exclude Account Type — FIP Notation Variants

Sources:
- `docs/x-checks_Formulas.xlsx`, sheet `final` (5,688 rows) — Family A variants
- `test_data/20260318 FIP X-Checks.txt` — Family B variants (LA003–LA006)

Two structurally distinct families of exclusion notation have been observed.

---

## Family A — `excl.acc.type` suffix variants

The exclusion appears as a **suffix appended to the end of the variable name**, inside the VAL_YTD/LC_YTD call, e.g. `LC_YTD(LIN_00380excl. acc. type = 2 ToM L07)`.  
In several cases a TOM/ToM movement-type tag follows the exclusion suffix — this is a separate concept and is not part of the exclusion pattern itself.

---

## Variant inventory

### 1. `excl. acc. type = 2` — equals, spaces, type 2

| X-Check | Full suffix as seen |
|---------|---------------------|
| AS447_09 | `excl. acc. type = 2` (no TOM tag) |
| LE417_09 | `excl. acc. type = 2 ToM L07` |
| LE419_09 | `excl. acc. type = 2 ToM L87` |

---

### 2. `excl. acc. type = 1` — equals, spaces, type 1

| X-Check | Full suffix as seen |
|---------|---------------------|
| LE418_09 | `excl. acc. type = 1 ToM L07` |
| LE420_09 | `excl. acc. type = 1 ToM L87` |

---

### 3. `excl. acc. type: 2` — colon (with space), type 2

| X-Check | Full suffix as seen |
|---------|---------------------|
| LE413_60 | `excl. acc. type: 2 TOM L09` |
| LE415_60 | `excl. acc. type: 2 TOM L08` |

---

### 4. `excl. acc. type:1,4` — colon (no space), multi-type (1 and 4)

| X-Check | Full suffix as seen |
|---------|---------------------|
| LE414_09 | `excl. acc. type:1,4 TOM L09` |
| LE414_60 | `excl. acc. type:1,4 TOM L09` |
| LE416_09 | `excl. acc. type:1,4 TOM L08` |
| LE416_60 | `excl. acc. type:1,4 TOM L08` |

---

### 5. `excl.acc.type2` — no separators, concatenated, type 2

| X-Check | Full suffix as seen |
|---------|---------------------|
| LE413_09 | `excl.acc.type2 TOM L09` |
| LE415_09 | `excl.acc.type2 TOM L08` |

---

### 6. `excl. acc. type 2` — space only as separator, type 2

| X-Check | Full suffix as seen |
|---------|---------------------|
| LS233_17 | `excl. acc. type 2` |

---

### 7. `excl.acct.type 2` — abbreviated "acct", no space before 2

| X-Check | Full suffix as seen |
|---------|---------------------|
| LS235_17 | `excl.acct.type 2` |

---

### 8. `excl. acct. type 2` — abbreviated "acct" with spaces, type 2

| X-Check | Full suffix as seen |
|---------|---------------------|
| LS237_17 | `excl. acct. type 2` |

---

## Summary table

| # | Core pattern | Separator | Abbrev | Types seen | X-Checks |
|---|-------------|-----------|--------|------------|----------|
| 1 | `excl. acc. type = N` | equals + spaces | acc | 1, 2 | AS447_09, LE417_09, LE418_09, LE419_09, LE420_09 |
| 2 | `excl. acc. type: N` | colon + space | acc | 2 | LE413_60, LE415_60 |
| 3 | `excl. acc. type:N,M` | colon, no space | acc | 1,4 (multi) | LE414_09, LE414_60, LE416_09, LE416_60 |
| 4 | `excl.acc.typeN` | none (concatenated) | acc | 2 | LE413_09, LE415_09 |
| 5 | `excl. acc. type N` | space only | acc | 2 | LS233_17 |
| 6 | `excl.acct.type N` | space before N only | **acct** | 2 | LS235_17 |
| 7 | `excl. acct. type N` | spaces throughout | **acct** | 2 | LS237_17 |

**Total: 7 distinct core patterns across 16 X-Checks.**

---

## Family A — Normalisation notes

- Strip the TOM/ToM movement-type tag (` TOM L07`, ` ToM L87`, etc.) before pattern matching — it is not part of the exclusion.
- `acc` and `acct` are two distinct abbreviations (not just typos of each other) — both must be handled.
- The separator between `type` and the number varies across: `=`, `:`, ` `, or nothing.
- Multi-type values use comma-separated lists (e.g. `1,4`) — the parser must handle N values, not just a single digit.
- All variants start with `excl` (no confirmed `exl` typo variant in this sheet, though one was observed elsewhere in AL167_00).

---

## Family B — Parenthetical annotation variants

Source: `test_data/20260318 FIP X-Checks.txt` only (not present in the 20251205 file).

The exclusion appears as a **parenthetical annotation embedded within the variable name**, between the account identifier and the ToM suffix, e.g. `LC_YTD(BSN (only 2-affiliated) ToM 354ff)`.

### 9. `(only 2-affiliated)` — include-only type 2

Meaning: restrict to affiliated accounts only (equivalent to excluding all other types).

| X-Check | Full variable name as seen |
|---------|---------------------------|
| LA003_09 | `BSN (only 2-affiliated) ToM 354ff` |
| LA004_09 | `BSN (only 2-affiliated) TOM 670ff` |

---

### 10. `(without 3rd party)` — exclude type 1

Meaning: exclude 3rd party accounts (type 1); retain affiliated and others.

| X-Check | Full variable name as seen |
|---------|---------------------------|
| LA003_09 | `SN_12895ff (without 3rd party) ToM 660ff` |
| LA004_09 | `SN_12895ff (without 3rd party) ToM 670ff` |
| LA006_09 | `SN_12895ff (without 3rd party) ToM 660ff` |

---

### 11. `(without affiliated)` — exclude type 2

Meaning: exclude affiliated accounts (type 2); retain 3rd party and others.  
Note: uses `/` as separator before ToM instead of a space.

| X-Check | Full variable name as seen |
|---------|---------------------------|
| LA005_09 | `SN_12895ff(without affiliated)/ToM 660ff` |

---

## Family B — Structural notes

- The annotation sits **between** the account name and the ToM suffix, not at the end of the variable name.
- Spacing is inconsistent: `BSN (only 2-affiliated) ToM` uses spaces around the parentheses; `SN_12895ff(without affiliated)/ToM` has no space before `(` and uses `/` as separator before ToM.
- Family B uses natural-language descriptions (`only`, `without`) rather than type numbers — the type number must be inferred from the description using the known type list (Type 1 = 3rd party, Type 2 = Affiliated).
- `(only 2-affiliated)` and `(without affiliated)` express opposite exclusions despite both involving type 2: "only 2" = keep type 2 / exclude everything else; "without affiliated" = exclude type 2 / keep everything else.
- Family B variants are **not handled** by the current `_EXCL_SUFFIX_RE` regex — they require a separate normalisation pass.

---

## Combined summary

| Family | # | Core pattern | Position in variable name | Types | X-Checks |
|--------|---|-------------|--------------------------|-------|----------|
| A | 1 | `excl. acc. type = N` | suffix (end) | 1, 2 | AS447_09, LE417–420_09 |
| A | 2 | `excl. acc. type: N` | suffix (end) | 2 | LE413_60, LE415_60 |
| A | 3 | `excl. acc. type:N,M` | suffix (end) | 1,4 | LE414, LE416 |
| A | 4 | `excl.acc.typeN` | suffix (end) | 2 | LE413_09, LE415_09 |
| A | 5 | `excl. acc. type N` | suffix (end) | 2 | LS233_17 |
| A | 6 | `excl.acct.type N` | suffix (end) | 2 | LS235_17 |
| A | 7 | `excl. acct. type N` | suffix (end) | 2 | LS237_17 |
| B | 8 | `(only 2-affiliated)` | mid (before ToM) | 2 | LA003_09, LA004_09 |
| B | 9 | `(without 3rd party)` | mid (before ToM) | 1 | LA003_09, LA004_09, LA006_09 |
| B | 10 | `(without affiliated)` | mid (before ToM) | 2 | LA005_09 |

**Total: 10 distinct patterns across 20 X-Checks.**
