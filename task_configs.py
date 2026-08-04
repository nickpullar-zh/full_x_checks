"""
Per-strategy file-upload configurations — all strategies combined for v1.0.
"""
from file_upload_config import UploadTaskConfig, FileFieldConfig, SectionConfig  # noqa: F401


COLLECT_LIVE_X_CHECKS_UPLOAD_CONFIG = UploadTaskConfig(
    task_name="Collect Live X-Checks",
    window_title="Collect Live X-Checks",
    requires_output_directory=True,
    file_fields=[
        FileFieldConfig(
            label="X-Checks Publication File",
            file_types=[("Excel Files", "*.xlsx")],
            description="The X-Checks Publication file with the 'cross checks all' sheet",
            default_sheet="cross checks all",
        ),
    ],
)

X_CHECKS_UPLOAD_CONFIG = UploadTaskConfig(
    task_name="X-Checks",
    window_title="X-Check Files",
    requires_output_directory=True,
    file_fields=[
        FileFieldConfig(
            label="X-Checks Publication File",
            file_types=[("Excel Files", "*.xlsx")],
            description="The X-Checks Publication file with the 'cross checks all' sheet",
            default_sheet="cross checks all"
        ),
        FileFieldConfig(
            label="FIP File",
            file_types=[("Text Files", "*.txt")],
            description="Data from the 'Validation Rule' in FIP Consolidation Workbench"
        ),
        FileFieldConfig(
            label="GCoA Publication File",
            file_types=[("Excel Files", "*.xlsx")],
            required=False,
            description="The X-Checks Publication file with the 'GCoA Base account table' sheet. If the file is not uploaded, no QU values can be returned",
            default_sheet="GCoA Base account table"
        ),
        FileFieldConfig(
            label="Known Exception List",
            file_types=[("Excel Files", "*.xlsx")],
            required=False,
            description="Spreadsheet with an 'X-Checks' sheet listing X-Check Numbers where EBX/FIP differences are expected and documented. If not uploaded, no exception flagging is applied.",
            default_sheet="X-Checks",
            sheet_editable=True,
        ),
    ],
    checkboxes=[
        {
            "key":     "apply_version_spanning",
            "label":   "Apply Version Spanning Validation (experimental)",
            "default": False,
            "tooltip": (
                "What it does: adds version numbers or GAAP framework prefixes to "
                "variable names in the EBX formula, matching how FIP expresses "
                "cross-version comparisons (e.g. 12602v100 vs 12602v800, or "
                "IFRSNS11930RA vs SLST15541ff).\n\n"
                "Why it is experimental: this rule has not yet been validated by the "
                "X-Checks team. Once confirmed correct it will become part of the "
                "standard output and this option will be removed."
            ),
        },
        {
            "key":     "apply_prior_year_balance",
            "label":   "Apply Prior Year Balance Formula (experimental)",
            "default": False,
            "tooltip": (
                "What it does: accounts flagged 'Ending Balance Prior Year' in the "
                "EBX file are expressed as P_VAL_PER(variable,'0','1') instead of "
                "VAL_YTD(variable), matching the FIP formula for prior-year opening "
                "balance checks. Affects X-Checks L003_00 and L019_00.\n\n"
                "Why it is experimental: this rule has not yet been validated by the "
                "X-Checks team. Once confirmed correct it will become part of the "
                "standard output and this option will be removed."
            ),
        },
    ]
)


GROUPING_BY_UPLOAD_CONFIG = UploadTaskConfig(
    task_name="Grouping By",
    window_title="Grouping By Files",
    requires_output_directory=True,
    file_fields=[
        FileFieldConfig(
            label="X-Checks Publication File",
            file_types=[("Excel Files", "*.xlsx")],
            description="The X-Checks Publication file with the 'cross checks all' sheet",
            default_sheet="cross checks all",
        ),
        FileFieldConfig(
            label="FIP File (ZQ9_VALFLDGR)",
            file_types=[("Excel Files", "*.xlsx")],
            description="Raw ZQ9_VALFLDGR export. Expected columns: ValidRule, Long Text, Field name.",
            default_sheet="Sheet1",
        ),
        FileFieldConfig(
            label="Mapping File",
            file_types=[("CSV / Text Files", "*.csv *.txt"), ("All Files", "*.*")],
            description="Mapping file in CSV or TXT format (FIP Data, EBX item)",
            default_sheet="Sheet1",
        ),
        FileFieldConfig(
            label="Known Exception List",
            file_types=[("Excel Files", "*.xlsx")],
            required=False,
            description="Spreadsheet with a 'Grouping By' sheet listing EBX Keys where differences are expected and documented. If not uploaded, no exception flagging is applied.",
            default_sheet="Grouping By",
            sheet_editable=True,
        ),
    ],
)

ACCOUNTING_PRINCIPLES_UPLOAD_CONFIG = UploadTaskConfig(
    task_name="Accounting Principles",
    window_title="Accounting Principles Files",
    requires_output_directory=True,
    file_fields=[
        FileFieldConfig(
            label="X-Checks Publication File",
            file_types=[("Excel Files", "*.xlsx")],
            description="EBX file with the 'cross checks all' sheet.",
            default_sheet="cross checks all",
            header_signals=["X-Check No.", "Status", "Type of change"],
        ),
        FileFieldConfig(
            label="FIP File (VALMSG)",
            file_types=[("Excel Files", "*.xlsx")],
            description="Raw ZQ9_VALMSG export. Expected columns: MethC, MK, Medium Text, ValidRule, Long Text, UCFV20G-TRUE_BRANCH, Message class, Msg., MT, Message Text. The Key column is calculated after import.",
            default_sheet="FIP Methods Rules and Condition",
        ),
        FileFieldConfig(
            label="Validation Methods File",
            file_types=[("Excel Files", "*.xlsx")],
            description="The 'validation methods.xlsx' template — sheet 'Validation Methods', "
                        "with Validation Events in row 1 and the current period block in rows 4–6.",
            default_sheet="Validation Methods",
        ),
        FileFieldConfig(
            label="Known Exception List",
            file_types=[("Excel Files", "*.xlsx")],
            required=False,
            description="Spreadsheet with an 'Accounting Principles' sheet listing rows where differences are expected and documented. If not uploaded, no exception flagging is applied.",
            default_sheet="Accounting Principles",
            sheet_editable=True,
        ),
    ],
)

CONDITIONS_UPLOAD_CONFIG = UploadTaskConfig(
    task_name="Conditions",
    window_title="Conditions Files",
    requires_output_directory=True,
    file_fields=[
        FileFieldConfig(
            label="X-Checks Publication File",
            file_types=[("Excel Files", "*.xlsx")],
            description="The X-Checks Publication file with the 'cross checks all' sheet",
            default_sheet="cross checks all",
        ),
        FileFieldConfig(
            label="FIP File (ZQ9_VALMETH)",
            file_types=[("Excel Files", "*.xlsx")],
            description="Raw ZQ9_VALMETH export. Expected columns: MethC, MK, Medium Text, ValidRule, Medium Text, UCFV20G-TRUE_BRANCH, ValidRule, Medium Text. Keys are calculated after import.",
            default_sheet="FIP Conditions",
        ),
        FileFieldConfig(
            label="Known Exception List",
            file_types=[("Excel Files", "*.xlsx")],
            required=False,
            description="Spreadsheet with a 'Conditions' sheet listing EBX/FIP pairs where differences are expected and documented. If not uploaded, no exception flagging is applied.",
            default_sheet="Conditions",
            sheet_editable=True,
        ),
    ],
)


def _build_full_run_config(registry: dict) -> UploadTaskConfig:
    """
    Build the Full Run UploadTaskConfig with an explicit ordered layout:
      X-Checks Publication File
      --- X-Checks ---
      FIP File
      GCoA Publication File
      --- Grouping By ---
      FIP File (ZQ9_VALFLDGR)
      Mapping File
      --- Accounting Principles ---
      FIP File (VALMSG)
      Validation Methods File
      --- Conditions ---
      FIP File (ZQ9_VALMETH)
      ---
      Known Exception List
    """
    # Build a lookup of label → FileFieldConfig from all registered strategies
    field_map: dict = {}
    for task_name, (config, _) in registry.items():
        if task_name in ("Full Run", "Collect Live X-Checks"):
            continue
        for field in config.file_fields:
            if isinstance(field, FileFieldConfig) and field.label not in field_map:
                field_map[field.label] = field

    def _f(label):
        return field_map[label]

    merged = [
        _f("X-Checks Publication File"),
        SectionConfig(title="X-Checks"),
        _f("FIP File"),
        _f("GCoA Publication File"),
        SectionConfig(title="Grouping By"),
        _f("FIP File (ZQ9_VALFLDGR)"),
        _f("Mapping File"),
        SectionConfig(title="Accounting Principles"),
        _f("FIP File (VALMSG)"),
        _f("Validation Methods File"),
        SectionConfig(title="Conditions"),
        _f("FIP File (ZQ9_VALMETH)"),
        SectionConfig(),   # plain separator before KEL
        FileFieldConfig(
            label="Known Exception List",
            file_types=[("Excel Files", "*.xlsx")],
            required=False,
            description=_f("Known Exception List").description,
            default_sheet="(per strategy)",
            sheet_note=(
                "Sheet name is set automatically per strategy "
                "(X-Checks / Grouping By / Accounting Principles / Conditions)."
            ),
        ),
    ]

    return UploadTaskConfig(
        task_name="Full Run",
        window_title="Full Run — All Strategies",
        requires_output_directory=True,
        file_fields=merged,
    )
