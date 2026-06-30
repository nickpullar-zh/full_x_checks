"""
Per-strategy file-upload configurations.

This file lives on the infrastructure (`main`) branch with no entries — each
strategy branch (e.g. v0.3-X-Checks, v0.5-Accounting-Principles) adds its own
UploadTaskConfig and registers it in task_registry.py.
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
