"""
Per-strategy file-upload configurations — all strategies combined for v1.0.
"""
from file_upload_config import UploadTaskConfig, FileFieldConfig  # noqa: F401


GROUPING_BY_UPLOAD_CONFIG = UploadTaskConfig(
    task_name="Grouping By",
    window_title="Grouping By Files",
    requires_output_directory=True,
    file_fields=[
        FileFieldConfig(
            label="FIP File (ZQ9_VALFLDGR)",
            file_types=[("Excel Files", "*.xlsx")],
            description="FIP download from ZQ9_VALFLDGR",
            default_sheet="Sheet1",
        ),
        FileFieldConfig(
            label="X-Checks Publication File",
            file_types=[("Excel Files", "*.xlsx")],
            description="The X-Checks Publication file with the 'cross checks all' sheet",
            default_sheet="cross checks all",
        ),
        FileFieldConfig(
            label="Mapping File",
            file_types=[("CSV Files", "*.csv"), ("All Files", "*.*")],
            description="Mapping file in CSV format (FIP Data, EBX item)",
            default_sheet="Sheet1",
        ),
    ],
)

ACCOUNTING_PRINCIPLES_UPLOAD_CONFIG = UploadTaskConfig(
    task_name="Accounting Principles",
    window_title="Accounting Principles Files",
    requires_output_directory=True,
    file_fields=[
        FileFieldConfig(
            label="Validation Methods File",
            file_types=[("Excel Files", "*.xlsx")],
            description="The 'validation methods.xlsx' template — sheet 'Validation Methods', "
                        "with Validation Events in row 1 and the current period block in rows 4–6.",
            default_sheet="Validation Methods",
        ),
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
            description="VALMSG export with the 'FIP Methods Rules and Condition' sheet "
                        "(rows keyed by '<Method>|<X-Check No.>' and an MT column of W/E).",
            default_sheet="FIP Methods Rules and Condition",
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
            label="FIP File",
            file_types=[("Excel Files", "*.xlsx")],
            description="FIP download from ZQ9_VALMETH (sheet: FIP Conditions)",
            default_sheet="FIP Conditions",
        ),
    ],
)


def _build_full_run_config(registry: dict) -> UploadTaskConfig:
    """
    Build a merged UploadTaskConfig from all registered strategies, deduplicating
    file fields by label. Called after the registry is fully populated.
    """
    seen: set = set()
    merged: list = []
    for task_name, (config, _) in registry.items():
        if task_name == "Full Run":
            continue
        for field in config.file_fields:
            if field.label not in seen:
                seen.add(field.label)
                merged.append(field)
    return UploadTaskConfig(
        task_name="Full Run",
        window_title="Full Run — All Strategies",
        requires_output_directory=True,
        file_fields=merged,
    )
