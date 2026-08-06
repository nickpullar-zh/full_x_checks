---
name: ui-style-guide
description: Reference for how file upload forms, dialog boxes, and Excel outputs must look in this project. Read before making any UI or output formatting change.
user-invocable: true
---

# X-Checks Application — UI & Output Style Guide

Read this before writing or modifying any tkinter UI code or Excel output formatting.

---

## Fonts

All UI text uses **Zurich Sans** (bundled under `templates/fonts/`).
- Titles / section headings: `"Zurich Sans Semibold"`, bold
- Body / labels / hints: `"Zurich Sans"`
- All font sizes are scaled to screen resolution via `_ui_scale()` — never hardcode px sizes directly:
  ```python
  scale    = _ui_scale()          # float in [0.70, 1.0]
  F_TITLE   = max(7, round(14 * scale))
  F_SECTION = max(7, round(10 * scale))
  F_BODY    = max(7, round(9  * scale))
  F_SMALL   = max(7, round(8  * scale))
  ```

Excel output fonts: `"Zurich Sans"` for body, `"Zurich Sans Semibold"` for headers.  
Test plan generators (in `docs/`) use the same fonts.

---

## Colours

### Brand palette (use these constants — never invent new hex codes)

| Name | Hex | Use |
|---|---|---|
| Dark Blue | `#23366F` | Primary text, headers, borders |
| Light Blue | `#91BFE3` | Header fills, section highlights |
| Logic Blue | `#D6E4F7` | Logic row tint (test plans) |
| White | `#FFFFFF` | Background |
| Alt Grey | `#ECEEEF` | Alternating row tint |
| Grey (tab) | `#808080` | Processing Log tab colour |

### Comparison output colours (from `BaseStrategy` class constants)

Always reference these — never redefine them in a strategy:

```python
FILL_GREEN  = PatternFill(start_color="C6EFCE", ...)   # Match
FONT_GREEN  = Font(color="276221")
FILL_RED    = PatternFill(start_color="FFC7CE", ...)   # MisMatch
FONT_RED    = Font(color="9C0006")
FILL_ORANGE = PatternFill(start_color="FFEB9C", ...)   # Not Found / Not in FIP
FONT_ORANGE = Font(color="9C6500")
FILL_BLUE   = PatternFill(start_color="91BFE3", ...)   # Known Exception (excepted)
FONT_BLUE   = Font(color="23366F")
```

### Strategy tab colours

Each strategy has its own `TAB_COLOUR` / `TAB_COLOUR_PASTEL` declared on the strategy class.
`_apply_tab_colours()` in `BaseStrategy` applies them — call it at the end of `apply_output_formatting()`.

| Strategy | Main colour | Pastel |
|---|---|---|
| X-Checks | `70AD47` (green) | `D9EAD3` |
| Grouping By | `ED7D31` (orange) | `FCE4D6` |
| Accounting Principles | `8BAFC7` (mid-blue) | `B4C6E7` |
| Conditions | `2167AE` (Zurich blue) | `BDD7EE` |

---

## File upload forms

### Layout

- Outer container: `ttk.Frame` with `padding="15"`, fills the window via `grid(sticky="nsew")`.
- Three rows in the outer frame:
  - Row 0: title label (`Zurich Sans Semibold`, `F_TITLE`)
  - Row 1: horizontal separator
  - Row 2: scrollable canvas (gets `weight=1` so it expands)
  - Row 3: controls frame (fixed height, below the canvas)
- The scrollable canvas holds all file fields; mousewheel scrolling is enabled.
- Three columns inside the fields frame:
  - Col 0: field label (fixed, `LABEL_COL_W` px)
  - Col 1: path label / hint (stretches, `weight=1`)
  - Col 2: Browse button (fixed, ~80 px, flush right with `padx=(2, 8)`)

### Dialog sizing

Dialog width is computed from screen scale:
```python
DIALOG_W = min(usable_w - 2*M, max(700, int(900 * scale)))
```

Dialog height is **content-driven** — measure actual content height after building, then cap at 85% of usable screen height:
```python
ideal_h  = title_h + inner.winfo_reqheight() + controls_h + 30
dialog_h = min(ideal_h, max(400, int(usable_h * 0.85)))
```
Never set a fixed dialog height. Short strategies get compact dialogs; Full Run gets taller ones.

### Field labels

- Required fields: `"Label *"` (asterisk on same line)
- Optional fields: `"Label\n(optional)"` (second line), `wraplength=LABEL_COL_W`

### Sheet name combobox

- Populated from the workbook's own sheet names on file selection — **never allow free text for non-editable fields**.
- State: `"readonly"` unless `field.sheet_editable=True`, in which case `"normal"`.
- Auto-selects `field.default_sheet` if present in the file, otherwise first sheet.

### Hint labels

- Shown below each field in column 1, grey text, `F_BODY` font.
- All hints use the same `wraplength` (computed from dialog width).

### Controls area (below canvas)

Fixed frame, not scrollable. Contents (top to bottom, all centred on a single weighted column):
1. Horizontal separator
2. "Process only differences" checkbox (checked by default)
3. Any extra strategy-specific checkboxes
4. Horizontal separator
5. **Button row** — `btn_frame` packed in the centre:
   - "Return to Selection" (`width=18`, `side="left"`, `padx=(0, 8)`)
   - "Proceed" (`width=18`, `side="left"`, `padx=(8, 0)`, disabled until all required fields filled)
6. Version label (`v{VERSION}`, grey, `F_SMALL`, bottom-left, `sticky="w"`)

### Tooltips

Experimental checkboxes carry tooltips via the `_Tooltip` class (defined in `file_upload_ui.py`).
- Delay: 600ms
- Background: `#DDE4E3` (Dove brand colour)
- Text: Dark Blue `#23366F`
- Wrap: 320px

---

## Progress dialog

- Two buttons, same `btn_frame` pattern as upload forms:
  - "Stop" / "Return to Form" / "Return to Selection" on the left (`side="left"`, `padx=(0, 8)`)
  - "Exit Application" on the right (`side="left"`, `padx=(8, 0)` or similar)
- Log text colour coding:
  - Green: completion lines (`"completed successfully"`)
  - Orange: not-found / not-matched lines
  - Red bold: error lines
  - Green dashes: separator lines

---

## Excel output (comparison workbooks)

### Sheet order convention

1. Source data sheets (EBX Data, FIP Data, etc.)
2. Comparison sheet — always last data sheet; workbook opens on this sheet by default
3. Filtered sheets if applicable (Matched Data, MisMatched Data, Not Found Data)
4. Processing Log — always last

### Header row

- Dark Blue fill (`#23366F`), white bold text, centre-wrap alignment, thin borders, `height=28`.

### Data rows

- Top-wrap alignment, thin borders.
- Comparison column fill: green / red / orange from `BaseStrategy` constants.
- Key / ID column: inherits the worst result colour for that row.

### Processing Log sheet

- Columns: Timestamp, File, Step, Count.
- First entry: app version.
- Includes: files loaded, strategy steps with row counts, output path, expected sensitivity label.

### Sensitivity label

Every output workbook gets the `Internal_Use_Only` MIP label applied via `BaseStrategy.write_excel_output()` → `ExcelLabeler`. Failure is swallowed silently so the run still completes.

### Tab colours

Applied at the end of `apply_output_formatting()` via `_apply_tab_colours(workbook)`. See colour table above. Processing Log tab is always grey `#808080`.

---

## Test plan generators (`docs/`)

Both `generate_fixture_uat.py` and `generate_full_app_uat.py` follow the same structure:

- **Sheet 1 — Overview**: merged wide cells, Zurich Sans, Light Blue label column.
- **Sheet 2 — Test Cases**: columns ID / Area / Test Type / Files+Setup / Steps / Expected Result / Actual Result / Pass-Fail / Tester / Date. Header: Dark Blue fill, white semibold. Logic rows: light-blue conditional format (`D6E4F7`) driven by `$C{row}="Logic"`.
- **Sheet 3 — Sign-off**: Light Blue label column, merged value cells.
- Filenames include a letter suffix (`_a`, `_b`, ...) that increments on each regeneration at the same version — never overwrites a previous output.
- All generated workbooks receive the `Internal_Use_Only` MIP sensitivity label.
