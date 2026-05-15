from file_upload_config import UploadTaskConfig, FileFieldConfig

# --- X-Checks Task ---
X_CHECKS_UPLOAD_CONFIG = UploadTaskConfig(
    task_name="X-Checks",
    window_title="X-Check Files",
    requires_output_directory=True,
    file_fields=[
        FileFieldConfig(
            label="FIP file",
            file_types=[("Text Files", "*.txt")],
            description="Data from the 'Validation Rule' in FIP Consolidation Workbench"
        ),
        FileFieldConfig(
            label="X-Checks Publication File",
            file_types=[("Excel Files", "*.xlsx")],
            description="The X-Checks Publication file with the 'cross checks all' sheet",
            default_sheet="cross checks all"
        ),
        FileFieldConfig(
            label="GCoA Publication File",
            file_types=[("Excel Files", "*.xlsx")],
            required=False, # If the file is not uploaded, no QU values can be returned
            description="The X-Checks Publication file with the 'GCoA Base account table' sheet\n  If the file is not uploaded, no QU values can be returned",
            default_sheet="GCoA Base account table"
        ),
        FileFieldConfig(
            label="Known Exception List",
            file_types=[("Excel Files", "*.xlsx")],
            required=False,
            description="Spreadsheet with a 'Known Exceptions' sheet listing X-Check Numbers where EBX/FIP differences are expected and documented. If not uploaded, no exception flagging is applied.",
            default_sheet="Known Exceptions"
        ),
    ]
)

# --- Accounting Principles Task ---
ACCOUNTING_PRINCIPLES_UPLOAD_CONFIG = UploadTaskConfig(
    task_name="X-Check Accounting Principles",
    window_title="X-Check Accounting Principles Files",
    requires_output_directory=True,
    file_fields=[
        FileFieldConfig(
            label="FIP File (ZQ9_VALMSG)",
            file_types=[("Excel Files", "*.xlsx")],
            description="Data from ZQ9_VALMSG in FIP",
            default_sheet="Sheet1"
        ),
        FileFieldConfig(
            label="X-Checks Publication File",
            file_types=[("Excel Files", "*.xlsx")],
            description="The X-Checks Publication file with the 'cross checks all' sheet",
            default_sheet="cross checks all"
        ),
        FileFieldConfig(
            label="Validation Methods",
            file_types=[("Excel Files", "*.xlsx")],
            description="The Validation Methods Excel file with the sheet 'Validation Methods'",
            default_sheet="Validation Methods"
        ),
    ]
)

# --- Conditions Task ---
CONDITIONS_UPLOAD_CONFIG = UploadTaskConfig(
    task_name="X-Check Conditions",
    window_title="X-Check Conditions Files",
    requires_output_directory=True,
    file_fields=[
        FileFieldConfig(
            label="FIP File (ZQ9_VALMETH)",
            file_types=[("Excel Files", "*.xlsx")],
            description="Data from ZQ9_VALMETH in FIP",
            default_sheet="Sheet1"
        ),
        FileFieldConfig(
            label="X-Checks Publication File",
            file_types=[("Excel Files", "*.xlsx")],
            description="The X-Checks Publication file with the 'cross checks all' sheet",
            default_sheet="cross checks all"
        ),
    ]
)

# --- Grouping By Task ---
GROUPING_BY_UPLOAD_CONFIG = UploadTaskConfig(
    task_name="X-Check Grouping By",
    window_title="X-Check Grouping By Files",
    requires_output_directory=True,
    file_fields=[
        FileFieldConfig(
            label="FIP File (ZQ9_VALFLDGR)",
            file_types=[("Excel Files", "*.xlsx")],
            description="Data from ZQ9_VALFLDGR in FIP",
            default_sheet="Sheet1"
        ),
        FileFieldConfig(
            label="X-Checks Publication File",
            file_types=[("Excel Files", "*.xlsx")],
            description="The X-Checks Publication file with the 'cross checks all' sheet",
            default_sheet="cross checks all",
            required_columns=["X-Check No.","Reference  X-Check (Condition)","Grouping By"]
        ),
        FileFieldConfig(
            label="Mapping File",
            file_types=[("Text Files", "*.txt")],
            description="Mapping file in CSV format"
        ),
    ]
)