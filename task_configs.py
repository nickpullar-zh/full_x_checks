"""
Per-strategy file-upload configurations.

This branch (v0.5-Accounting-Principles) registers exactly one strategy:
Accounting Principles. Other strategies live on their own branches.
"""
from file_upload_config import UploadTaskConfig, FileFieldConfig


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
            description="EBX file containing both the 'cross checks all' sheet and the "
                        "'FIP Methods Rules and Condition' sheet (the VALMSG dump). "
                        "The strategy reads both sheets from this single file.",
            default_sheet="cross checks all",
        ),
    ],
)
