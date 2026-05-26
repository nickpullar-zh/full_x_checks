"""
Generate briefing PDF: Include/Exclude account handling in FIP and EBX
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, PageBreak, KeepTogether
)
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
import os

OUT = os.path.join(os.path.dirname(__file__),
                   'X-Checks_Include_Exclude_Briefing.pdf')

# ── Colour palette ────────────────────────────────────────────────────────────
ZURICH_BLUE   = colors.HexColor('#2167AE')   # Zurich Blue (brand)
LIGHT_BLUE    = colors.HexColor('#91BFE3')   # Light Blue (brand)
HEADER_GREY   = colors.HexColor('#4A4A4A')
ROW_GREY      = colors.HexColor('#F5F5F5')
GREEN_MATCH   = colors.HexColor('#C6EFCE')
RED_MISMATCH  = colors.HexColor('#FFC7CE')
ORANGE_NA     = colors.HexColor('#FFEB9C')
BORDER        = colors.HexColor('#CCCCCC')

# ── Styles ────────────────────────────────────────────────────────────────────
styles = getSampleStyleSheet()

title_style = ParagraphStyle('Title', parent=styles['Normal'],
    fontSize=20, textColor=ZURICH_BLUE, spaceAfter=6,
    fontName='Helvetica-Bold', alignment=TA_LEFT)

subtitle_style = ParagraphStyle('Subtitle', parent=styles['Normal'],
    fontSize=11, textColor=HEADER_GREY, spaceAfter=16,
    fontName='Helvetica', alignment=TA_LEFT)

h1_style = ParagraphStyle('H1', parent=styles['Normal'],
    fontSize=14, textColor=ZURICH_BLUE, spaceBefore=14, spaceAfter=6,
    fontName='Helvetica-Bold', borderPad=4)

h2_style = ParagraphStyle('H2', parent=styles['Normal'],
    fontSize=11, textColor=HEADER_GREY, spaceBefore=10, spaceAfter=4,
    fontName='Helvetica-Bold')

h3_style = ParagraphStyle('H3', parent=styles['Normal'],
    fontSize=10, textColor=HEADER_GREY, spaceBefore=8, spaceAfter=3,
    fontName='Helvetica-BoldOblique')

body_style = ParagraphStyle('Body', parent=styles['Normal'],
    fontSize=9, leading=13, spaceAfter=6,
    fontName='Helvetica', alignment=TA_JUSTIFY)

bullet_style = ParagraphStyle('Bullet', parent=styles['Normal'],
    fontSize=9, leading=13, spaceAfter=3,
    fontName='Helvetica', leftIndent=14, bulletIndent=4)

code_style = ParagraphStyle('Code', parent=styles['Normal'],
    fontSize=8, leading=11, spaceAfter=2,
    fontName='Courier', backColor=colors.HexColor('#F8F8F8'),
    leftIndent=10, borderPad=3)

caption_style = ParagraphStyle('Caption', parent=styles['Normal'],
    fontSize=8, textColor=HEADER_GREY, spaceAfter=8,
    fontName='Helvetica-Oblique', alignment=TA_CENTER)

note_style = ParagraphStyle('Note', parent=styles['Normal'],
    fontSize=8.5, leading=12, spaceAfter=6,
    fontName='Helvetica-Oblique', textColor=HEADER_GREY,
    leftIndent=10, borderPad=4)

# ── Table helpers ─────────────────────────────────────────────────────────────

def header_cell(text):
    return Paragraph(f'<b>{text}</b>', ParagraphStyle('TH', parent=styles['Normal'],
        fontSize=8.5, fontName='Helvetica-Bold', textColor=colors.white, leading=11))

def cell(text, bold=False, colour=None):
    st = ParagraphStyle('TD', parent=styles['Normal'],
        fontSize=8, fontName='Helvetica-Bold' if bold else 'Helvetica',
        leading=11, wordWrap='LTR')
    return Paragraph(str(text), st)

def code_cell(text):
    return Paragraph(f'<font name="Courier" size="7.5">{text}</font>',
                     ParagraphStyle('TC', parent=styles['Normal'], fontSize=7.5, leading=10))

def make_table(headers, rows, col_widths, row_colours=None):
    data = [[header_cell(h) for h in headers]]
    for i, row in enumerate(rows):
        data.append([cell(v) for v in row])

    ts = TableStyle([
        ('BACKGROUND',  (0,0), (-1,0), ZURICH_BLUE),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, ROW_GREY]),
        ('GRID',        (0,0), (-1,-1), 0.4, BORDER),
        ('VALIGN',      (0,0), (-1,-1), 'TOP'),
        ('TOPPADDING',  (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 5),
        ('RIGHTPADDING',(0,0), (-1,-1), 5),
    ])
    if row_colours:
        for row_idx, bg in row_colours.items():
            ts.add('BACKGROUND', (0, row_idx), (-1, row_idx), bg)

    t = Table(data, colWidths=col_widths, repeatRows=1)
    t.setStyle(ts)
    return t

# ── Document sections ─────────────────────────────────────────────────────────

def section_rule():
    return HRFlowable(width='100%', thickness=1, color=ZURICH_BLUE,
                      spaceAfter=6, spaceBefore=4)

def sub_rule():
    return HRFlowable(width='100%', thickness=0.5, color=BORDER,
                      spaceAfter=4, spaceBefore=2)

# ── Build content ─────────────────────────────────────────────────────────────

story = []
W = 17*cm  # usable width

# ── Cover ──────────────────────────────────────────────────────────────────────
story.append(Spacer(1, 2*cm))
story.append(Paragraph('X-Checks Automation', subtitle_style))
story.append(Paragraph('Include &amp; Exclude Account Handling', title_style))
story.append(Paragraph('Representation in FIP Text Output and EBX Excel Workbook', subtitle_style))
story.append(section_rule())
story.append(Paragraph(
    'This briefing documents every mechanism used in the FIP Validation Rule text output '
    'and the EBX Cross Checks publication file to represent included and excluded accounts '
    'or reporting units. It covers observed patterns, current match/mismatch status, and '
    'the implications for the X-Checks comparison process.',
    body_style))
story.append(Spacer(1, 0.4*cm))

# ── 1. Background ──────────────────────────────────────────────────────────────
story.append(Paragraph('1.  Background', h1_style))
story.append(section_rule())
story.append(Paragraph(
    'The X-Checks process compares formulas and variables extracted from two sources: '
    'the EBX "cross checks all" Excel sheet (which defines the rules) and the FIP '
    'Validation Rule text output (which records what the EPM system has actually applied). '
    'Both sources can restrict a check to a subset of reporting units (Included/Excluded RUs) '
    'or exclude specific account types from the calculation. '
    'These restrictions are encoded differently in each source.',
    body_style))
story.append(Paragraph(
    'The EBX workbook uses three dedicated columns: <b>Included RUs</b>, <b>Excluded RUs</b>, '
    'and <b>Exclude Account Type</b>. The FIP text output embeds the restriction directly inside '
    'the variable name within the formula string, using a variety of notations.',
    body_style))

# ── 2. FIP Representation ──────────────────────────────────────────────────────
story.append(PageBreak())
story.append(Paragraph('2.  FIP Text Output — Exclusion Representation', h1_style))
story.append(section_rule())
story.append(Paragraph(
    'FIP encodes exclusions as suffixes appended to the variable name inside VAL_YTD() calls. '
    'Four distinct mechanisms have been identified, with the account-type exclusion appearing '
    'in six different syntactic variants.',
    body_style))

story.append(Paragraph('2.1  Mechanism 1 — EXN_ Account Prefix', h2_style))
story.append(sub_rule())
story.append(Paragraph(
    'Certain FS Account codes carry an <b>EXN_</b> prefix, indicating that they hold '
    'aggregated data for excluded entities. These accounts appear in the variable section '
    'of the FIP block alongside normal accounts. The exclusion is structural — it is baked '
    'into the account code itself rather than being a formula annotation.',
    body_style))
story.append(Paragraph('32 distinct EXN_ codes are present in the test file:', body_style))

exn_data = [
    ['EXN_11100','EXN_11121','EXN_11122','EXN_11131','EXN_11132','EXN_11140'],
    ['EXN_11221','EXN_11222','EXN_12100','EXN_12251','EXN_12252','EXN_12260'],
    ['EXN_13100','EXN_14100','EXN_15100','EXN_15210','EXN_21100','EXN_21220'],
    ['EXN_21220ff','EXN_21231','EXN_21231ff','EXN_21232','EXN_21233','EXN_21236'],
    ['EXN_21240','EXN_21250','EXN_22100','EXN_22209','EXN_22210','EXN_23100'],
    ['EXN_24100','EXN_25100','EXN_25210','—','—','—'],
]
exn_table_data = [[header_cell(c) for c in ['','','','','','']]]
for row in exn_data:
    exn_table_data.append([code_cell(v) for v in row])
exn_ts = TableStyle([
    ('BACKGROUND',  (0,0), (-1,0), ZURICH_BLUE),
    ('GRID',        (0,0), (-1,-1), 0.4, BORDER),
    ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, ROW_GREY]),
    ('TOPPADDING',  (0,0), (-1,-1), 3),
    ('BOTTOMPADDING', (0,0), (-1,-1), 3),
    ('LEFTPADDING', (0,0), (-1,-1), 5),
])
exn_t = Table(exn_table_data, colWidths=[W/6]*6)
exn_t.setStyle(exn_ts)
story.append(exn_t)
story.append(Spacer(1, 0.3*cm))
story.append(Paragraph(
    'Example (LR048_17): FIP variables include EXN_11121, EXN_11122, EXN_12100, '
    'EXN_21100, EXN_22100 etc. alongside the main FS accounts.',
    note_style))

story.append(Paragraph('2.2  Mechanism 2 — Account Type Exclusion Suffix', h2_style))
story.append(sub_rule())
story.append(Paragraph(
    'When an account type is excluded, FIP appends a suffix directly to the variable name '
    'inside the VAL_YTD() call. The same concept appears in six syntactic variants:',
    body_style))

suffix_rows = [
    ['excl.acc.type=2',        'VAL_YTD(IAN_00051excl.acc.type=2)',        'AS447_09, AS448_09'],
    ['excl.acc.type:2',        'VAL_YTD(IAN_00023exl.acc.type:2)',         'AS130_00'],
    ['exl.acc.type2',          'VAL_YTD(LIN_00355ffexl.acc.type2)',        'AL167_00  (typo: "exl" not "excl")'],
    ['excl.acct.type 2',       'VAL_YTD(EXN_21231ff excl.acct.type 2)',   'EBX variables section'],
    ['excl. acc. type = 2',    'VAL_YTD(LIN_00380 excl. acc. type = 2 ToM L07)', 'Formula text'],
    ['excl. acc. type:1,4',    'VAL_YTD(LIN_00380 excl. acc. type:1,4 TOM L09)', 'Multi-type exclusion'],
    ['excl.acc.type2 (no sep)', 'VAL_YTD(LIN_00380 excl.acc.type2 TOM L08)',     'Compressed form'],
]
story.append(make_table(
    ['Suffix Form', 'Example', 'Observed In'],
    suffix_rows,
    [3.5*cm, 7*cm, 6.5*cm]
))
story.append(Spacer(1, 0.2*cm))
story.append(Paragraph(
    'All variants express the same concept: exclude account type 2 (Affiliated). '
    'The inconsistency in spacing, punctuation and abbreviation ("acc" vs "acct", '
    '"excl" vs "exl") originates from the EPM system and must be normalised before comparison.',
    note_style))

story.append(Paragraph('2.3  Mechanism 3 — Posting Level Exclusion (excl PL)', h2_style))
story.append(sub_rule())
story.append(Paragraph(
    'A single observed case uses <b>excl PL</b> notation to exclude a specific posting level:',
    body_style))
story.append(Paragraph('VAL_YTD(23151ffexclPL10) = 0', code_style))
story.append(Paragraph(
    'This appears in X-Check AL109_70. The EBX file has Excluded RUs = CON_30104 for this '
    'check, but there is no direct column mapping between a CON_ code and a posting-level '
    'exclusion. This is an isolated case.',
    note_style))

story.append(Paragraph('2.4  Mechanism 4 — Text Annotations (excl. CH, excluding SXX)', h2_style))
story.append(sub_rule())
story.append(Paragraph(
    'Some formulas carry free-text exclusion annotations within the variable name:',
    body_style))
annot_rows = [
    ['excl. CH / excl CH units', 'Excludes Swiss consolidation units', 'AL166_00, L123_00 context text'],
    ['excluding S02',            'Excludes segment S02',               'VAL_YTD(10102 excluding S02)'],
    ['excluding S31',            'Excludes segment S31',               'VAL_YTD(10102 excluding S31)'],
    ['excl. 2 - Aff.',           'Excludes affiliated (Group 2)',      'SLSTLSAS_11520n-lifeLOBsexcl.2-Aff'],
    ['excl.2 -Aff',              'Same — compressed variant',          'Variable name suffix'],
]
story.append(make_table(
    ['Annotation', 'Meaning', 'Example Context'],
    annot_rows,
    [4.5*cm, 5.5*cm, 7*cm]
))

# ── 3. EBX Representation ──────────────────────────────────────────────────────
story.append(PageBreak())
story.append(Paragraph('3.  EBX Workbook — Exclusion Columns', h1_style))
story.append(section_rule())
story.append(Paragraph(
    'The EBX "cross checks all" sheet has four columns relevant to inclusion/exclusion. '
    'All are populated at row level (one row per account/sub-account combination per X-Check), '
    'though in practice the values are consistent across all rows of the same X-Check.',
    body_style))

story.append(Paragraph('3.1  Included RUs', h2_style))
story.append(sub_rule())
story.append(Paragraph(
    '<b>93 rows across 89 distinct X-Checks</b> carry an Included RUs value. '
    'This specifies which consolidation unit the check applies to — only that RU '
    'is in scope. The 11 distinct values observed are:',
    body_style))
incl_rows = [
    ['CON_LFLOB',  '72 rows', '~40 X-Checks', 'Life France Line of Business'],
    ['CON_ZIC',    '11 rows', '~10 X-Checks', 'Zurich Insurance Company'],
    ['CON_RR',     '6 rows',  '6 X-Checks',   'Reinsurance entity'],
    ['CON_ARG2',   '2 rows',  '2 X-Checks',   'Argentina entity'],
    ['CON_DE',     '12 rows', '12 X-Checks',  'Germany entity'],
    ['CON_ZIP',    '1 row',   '1 X-Check',    'Zurich Insurance Plc'],
    ['CON_CH2',    '2 rows',  '2 X-Checks',   'Swiss entity variant'],
    ['CON_ZIC',    '—',       '—',            '(see above)'],
    ['CON_GI17',   '1 row',   '1 X-Check',    'General Insurance 2017'],
    ['CON_PC&BBA', '1 row',   '1 X-Check',    'P&C / B&A entity'],
    ['CON_201000', '2 rows',  '2 X-Checks',   'Entity code 201000'],
    ['CON_123SII', '1 row',   '1 X-Check',    'SII entity'],
    ['CON_ZICZR',  '1 row',   '1 X-Check',    'ZIC Zurich Reinsurance'],
    ['CON_ZLIC',   '1 row',   '1 X-Check',    'Zurich Life Insurance Company'],
]
story.append(make_table(
    ['CON_ Code', 'Row Count', 'X-Check Count', 'Entity'],
    [r for r in incl_rows if r[1] != '—'],
    [3.5*cm, 2.5*cm, 3.5*cm, 7.5*cm]
))
story.append(Spacer(1, 0.2*cm))
story.append(Paragraph(
    'Effect on FIP formula: None. In every tested case, X-Checks with Included RUs '
    'produce identical formulas in both FIP and EBX. The CON_ code is informational '
    'metadata only — it does not change the formula text.',
    note_style))

story.append(Paragraph('3.2  Excluded RUs', h2_style))
story.append(sub_rule())
story.append(Paragraph(
    '<b>171 rows across 30 distinct X-Checks</b> carry an Excluded RUs value. '
    'This specifies which consolidation unit is out of scope for the check. '
    '25 distinct CON_ codes are observed:',
    body_style))
excl_ru_rows = [
    ['CON_400069', '50 rows', '1 X-Check',  'Match', 'A142_00'],
    ['CON_DIR',    '29 rows', 'Multiple',   'Match', 'Directorate entity'],
    ['CON_ZIP',    '26 rows', 'Multiple',   'Match', 'Zurich Insurance Plc'],
    ['CON_30104',  '10 rows', '1 X-Check',  'MisMatch', 'AL109_70 — FIP adds "exclPL10"'],
    ['CON_ARG2',   '7 rows',  'Multiple',   'Match', 'Argentina entity'],
    ['CON_RR',     '7 rows',  'Multiple',   'Match', 'Reinsurance entity'],
    ['CON_51012',  '6 rows',  'Multiple',   'Match', '—'],
    ['CON_SWISS',  '4 rows',  '1 X-Check',  'Match', 'LS601_00'],
    ['CON_SISTER', '4 rows',  '1 X-Check',  'Match', 'L019_00'],
    ['CON_201001', '4 rows',  '1 X-Check',  'Match', 'LS169_00'],
    ['CON_CH',     '2 rows',  '2 X-Checks', 'Match', 'AL166_00, L123_00'],
    ['CON_FA',     '3 rows',  '1 X-Check',  'Match', 'R007_00'],
    ['CON_DE',     '1 row',   '1 X-Check',  'Match', 'S123_70'],
    ['Others',     '~30 rows','Multiple',   'Match', '12 further CON_ codes'],
]
story.append(make_table(
    ['CON_ Code', 'Rows', 'X-Checks', 'Formula Match', 'Notes'],
    excl_ru_rows,
    [2.8*cm, 1.8*cm, 2.5*cm, 2.5*cm, 7.4*cm],
    row_colours={4: ORANGE_NA}  # AL109_70 row
))
story.append(Spacer(1, 0.2*cm))
story.append(Paragraph(
    'Effect on FIP formula: None in almost all cases. The single exception is AL109_70 '
    '(CON_30104) where FIP appends "exclPL10" — a posting-level exclusion that has no '
    'direct column equivalent in EBX. All other Excluded RUs produce identical formulas.',
    note_style))

story.append(Paragraph('3.3  Exclude Account Type', h2_style))
story.append(sub_rule())
story.append(Paragraph(
    '<b>130 rows across 50 distinct X-Checks</b> carry an Exclude Account Type value. '
    'This is the most impactful column — it directly affects the formula structure '
    'in FIP. Five distinct values are observed:',
    body_style))
excl_acct_rows = [
    ['2 - Affiliated',                    '118', 'MisMatch', 'FIP appends excl.acc.type=2 suffix; may also flip operator from + to −'],
    ['2 - Affiliated, 6 - Unit linked',   '5',   'Match',    'FIP does NOT add notation for this combination — formulas match as-is'],
    ['2 - Affiliated, 9 - IC Difference', '3',   'Match',    'FIP does NOT add notation for this combination — formulas match as-is'],
    ['1 - 3rd, 4 - Linked',               '2',   'TBC',      'Not yet observed in FIP output — needs further investigation'],
    ['1-3rd, 4-Linked, 6-Unit linked',    '2',   'TBC',      'Not yet observed in FIP output — needs further investigation'],
]
story.append(make_table(
    ['EBX Value', 'Row Count', 'Formula Match', 'Notes'],
    excl_acct_rows,
    [5.5*cm, 2.2*cm, 2.5*cm, 6.8*cm],
    row_colours={1: RED_MISMATCH, 2: GREEN_MATCH, 3: GREEN_MATCH, 4: ORANGE_NA, 5: ORANGE_NA}
))

story.append(Paragraph('3.4  Exclude Z-Core', h2_style))
story.append(sub_rule())
story.append(Paragraph(
    'This column is present in the EBX workbook but contains no data in either test file. '
    'No action required.',
    body_style))

# ── 4. Formula Impact Detail ────────────────────────────────────────────────────
story.append(PageBreak())
story.append(Paragraph('4.  Detailed Formula Impact — Exclude Account Type = 2 - Affiliated', h1_style))
story.append(section_rule())
story.append(Paragraph(
    'This is the only case where a formula mismatch is consistently produced by an exclusion '
    'column. Two sub-patterns are observed depending on the formula structure:',
    body_style))

story.append(Paragraph('4.1  Pattern A — Suffix only, no operator change', h2_style))
story.append(sub_rule())
story.append(Paragraph(
    'When the excluded account is the <i>only</i> variable (or all variables share '
    'the same operator), FIP appends the exclusion suffix to the variable name '
    'without changing the operator:',
    body_style))
story.append(Paragraph('EBX:  ABS(VAL_YTD(LIN_00355ff)) >= CONST(1000,\'USD\',\'E\')', code_style))
story.append(Paragraph('FIP:  ABS(VAL_YTD(LIN_00355ffexl.acc.type2)) >= CONST(1000,\'USD\',\'E\')', code_style))
story.append(Paragraph('X-Check: AL167_00  |  Account: LIN_00355^OAN_00261  |  Excl: 2 - Affiliated', note_style))

story.append(Paragraph('4.2  Pattern B — Suffix + operator flip + variable reorder', h2_style))
story.append(sub_rule())
story.append(Paragraph(
    'When the excluded account is one of multiple variables in an addition formula, '
    'FIP restructures the expression: the excluded variable moves to the end, '
    'its operator changes from <b>+</b> to <b>−</b>, and the exclusion suffix is appended:',
    body_style))
story.append(Paragraph('EBX:  ABS(VAL_YTD(IAN_00051) + VAL_YTD(S12301ffToM349)) <= CONST(5,\'USD\',\'E\')', code_style))
story.append(Paragraph('FIP:  ABS(VAL_YTD(S12301ffToM349) - VAL_YTD(IAN_00051excl.acc.type=2)) <= CONST(5,\'USD\',\'E\')', code_style))
story.append(Paragraph('X-Check: AS447_09  |  Excluded account: IAN_00051  |  Excl: 2 - Affiliated', note_style))
story.append(Spacer(1, 0.2*cm))

story.append(Paragraph('Summary of observed mismatching X-Checks:', h3_style))
detail_rows = [
    ['AL167_00', '2 - Affiliated', 'LIN_00355ff',  'Suffix only',            'LIN_00355ffexl.acc.type2',      'Pattern A'],
    ['AS130_00', '2 - Affiliated', 'IAN_00023',    'Suffix + op flip',       'IAN_00023exl.acc.type:2',       'Pattern B'],
    ['AS447_09', '2 - Affiliated', 'IAN_00051',    'Suffix + op flip',       'IAN_00051excl.acc.type=2',      'Pattern B'],
    ['AS448_09', '2 - Affiliated', 'SN_13105',     'Suffix + op flip',       'SN_13105excl.acc.type=2',       'Pattern B'],
]
story.append(make_table(
    ['X-Check', 'EBX Excl. Type', 'Excluded Acct', 'Change Type', 'FIP Variable Name', 'Pattern'],
    detail_rows,
    [2.2*cm, 3*cm, 2.5*cm, 3*cm, 4.5*cm, 1.8*cm]
))

story.append(Paragraph('4.3  FIP Notation Variants for Account Type 2', h2_style))
story.append(sub_rule())
story.append(Paragraph(
    'All of the following have been observed in the FIP text for account type 2 exclusions. '
    'They represent the same concept and must be normalised to a single canonical form:',
    body_style))
norm_rows = [
    ['excl.acc.type=2',      'AS447_09, AS448_09'],
    ['exl.acc.type:2',       'AS130_00  (note: typo "exl" not "excl")'],
    ['exl.acc.type2',        'AL167_00  (note: typo "exl" not "excl")'],
    ['excl.acct.type 2',     'EXN_ variable blocks'],
    ['excl. acc. type = 2',  'LIN_00380 formula text'],
    ['excl.acc.type2',       'LIN_00380 compressed formula text'],
]
story.append(make_table(
    ['FIP Notation', 'Example X-Check / Context'],
    norm_rows,
    [6*cm, 11*cm]
))

# ── 5. Current Match Status ─────────────────────────────────────────────────────
story.append(PageBreak())
story.append(Paragraph('5.  Current Match Status Summary', h1_style))
story.append(section_rule())

status_rows = [
    ['Included RUs',              'All',    '89',  'Match',    'CON_ codes are metadata; no formula change'],
    ['Excluded RUs (most)',       'All',    '29',  'Match',    'CON_ codes are metadata; no formula change'],
    ['Excluded RUs (AL109_70)',   'CON_30104', '1','MisMatch', 'FIP adds "exclPL10" — no EBX column maps to posting level'],
    ['Excl. Acc. Type 2-Aff.',   '2 - Affiliated', '~4', 'MisMatch', 'FIP appends suffix + may flip operator'],
    ['Excl. Acc. Type 2+6',      '2-Aff, 6-UL',  '~5', 'Match',   'FIP does not add notation for this combination'],
    ['Excl. Acc. Type 2+9',      '2-Aff, 9-IC',  '~3', 'Match',   'FIP does not add notation for this combination'],
    ['Excl. Acc. Type 1,4',      '1-3rd, 4-Lnk', 'TBC','Unknown', 'Not yet observed in FIP test data'],
    ['SLST/IFRSN prefix mismatch','n/a',   '~6', 'MisMatch',  'Structural naming difference; separate from exclusion handling'],
]
story.append(make_table(
    ['Exclusion Type', 'EBX Value', 'X-Checks', 'Status', 'Notes'],
    status_rows,
    [4*cm, 3.5*cm, 1.8*cm, 2.2*cm, 5.5*cm],
    row_colours={
        1: GREEN_MATCH, 2: GREEN_MATCH,
        3: RED_MISMATCH, 4: RED_MISMATCH,
        5: GREEN_MATCH, 6: GREEN_MATCH,
        7: ORANGE_NA, 8: RED_MISMATCH
    }
))

# ── 6. Implications ──────────────────────────────────────────────────────────────
story.append(Paragraph('6.  Implications for the Comparison Process', h1_style))
story.append(section_rule())

story.append(Paragraph('6.1  What must be built', h2_style))
story.append(sub_rule())
story.append(Paragraph(
    'To achieve matching formulas for <b>Exclude Account Type = 2 - Affiliated</b>:',
    body_style))
story.append(Paragraph('(a)  FIP Normalisation', h3_style))
story.append(Paragraph(
    'The 6 syntactic variants of account type 2 exclusion in FIP must be normalised '
    'to a single canonical form before comparison. The recommended canonical form is '
    '<b>excl.acc.type=2</b> (the most common form observed).',
    bullet_style))
story.append(Paragraph('(b)  EBX Formula Enhancement', h3_style))
story.append(Paragraph(
    'For accounts with Exclude Account Type = "2 - Affiliated", the EBX formula builder must:',
    bullet_style))
story.append(Paragraph('• Append the canonical exclusion suffix to the variable name', bullet_style))
story.append(Paragraph('• Move the excluded variable to the end of the VAL_YTD expression', bullet_style))
story.append(Paragraph('• Change its operator from + to − (Pattern B cases)', bullet_style))
story.append(Paragraph('• Leave the operator unchanged for Pattern A cases (single variable or ABS structure)', bullet_style))

story.append(Paragraph('6.2  What does NOT need to be built', h2_style))
story.append(sub_rule())
story.append(Paragraph(
    'The following require no code changes as they either already produce matching '
    'formulas or are out of scope:',
    body_style))
story.append(Paragraph('• Included RUs — purely informational; no formula impact', bullet_style))
story.append(Paragraph('• Excluded RUs (all except AL109_70) — purely informational; no formula impact', bullet_style))
story.append(Paragraph('• Excl. Acc. Type 2 + 6 and 2 + 9 — FIP already matches EBX without change', bullet_style))
story.append(Paragraph('• SLST / IFRSN prefix mismatches — structural naming difference; separate investigation required', bullet_style))

story.append(Paragraph('6.3  Open questions', h2_style))
story.append(sub_rule())
story.append(Paragraph(
    '• <b>AL109_70 / excl PL</b>: No EBX column maps to posting-level exclusions. '
    'Is this a gap in the EBX file, or is posting-level exclusion out of scope?',
    bullet_style))
story.append(Paragraph(
    '• <b>Excl. Acc. Type 1-3rd, 4-Linked</b>: Not yet observed in FIP test data. '
    'What notation does FIP use for account types 1 and 4?',
    bullet_style))
story.append(Paragraph(
    '• <b>Pattern A vs Pattern B determination</b>: The rule for when to flip the '
    'operator (Pattern B) vs leave it unchanged (Pattern A) is not yet fully defined. '
    'Further examples are needed to confirm whether it is determined by the Operator '
    '(X-Check Term) column value or the formula structure.',
    bullet_style))

# ── 7. Data source note ──────────────────────────────────────────────────────────
story.append(Paragraph('7.  Data Sources', h1_style))
story.append(section_rule())
src_rows = [
    ['EBX workbook',    '20251205 EPM X-Checks - Original - Copy.xlsx',  'Sheet: cross checks all'],
    ['FIP text file',   '20251205 FIP X-Checks - Original.txt',           'Full Validation Rule export'],
    ['Comparison output','20260515_095332_X-Checks Comparison.xlsx',      'Generated by new X-Checks app'],
]
story.append(make_table(
    ['Source', 'Filename', 'Notes'],
    src_rows,
    [3*cm, 8.5*cm, 5.5*cm]
))
story.append(Spacer(1, 0.3*cm))
story.append(Paragraph(
    'Analysis date: 2026-05-15  |  Branch: v0.3-X-Checks  |  '
    'Generated by X-Checks automation tooling',
    caption_style))

# ── Build PDF ──────────────────────────────────────────────────────────────────
doc = SimpleDocTemplate(
    OUT,
    pagesize=A4,
    leftMargin=2*cm, rightMargin=2*cm,
    topMargin=2*cm,  bottomMargin=2*cm,
    title='X-Checks Include/Exclude Briefing',
    author='X-Checks Automation',
)

def on_page(canvas, doc):
    canvas.saveState()
    # Header bar
    canvas.setFillColor(ZURICH_BLUE)
    canvas.rect(2*cm, A4[1]-1.4*cm, A4[0]-4*cm, 0.6*cm, fill=1, stroke=0)
    canvas.setFont('Helvetica-Bold', 8)
    canvas.setFillColor(colors.white)
    canvas.drawString(2.2*cm, A4[1]-1.1*cm, 'X-Checks Automation — Include/Exclude Briefing')
    canvas.drawRightString(A4[0]-2*cm, A4[1]-1.1*cm, f'Page {doc.page}')
    # Footer
    canvas.setFont('Helvetica', 7)
    canvas.setFillColor(HEADER_GREY)
    canvas.drawString(2*cm, 1.2*cm, 'INTERNAL — Zurich Insurance  |  2026-05-15')
    canvas.drawRightString(A4[0]-2*cm, 1.2*cm, 'v0.3-X-Checks')
    canvas.restoreState()

doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
print(f'PDF written to: {OUT}')
