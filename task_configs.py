"""
Per-strategy file-upload configurations.

This file lives on the infrastructure (`main`) branch with no entries — each
strategy branch (e.g. v0.3-X-Checks, v0.5-Accounting-Principles) adds its own
UploadTaskConfig and registers it in task_registry.py.
"""
from file_upload_config import UploadTaskConfig, FileFieldConfig  # noqa: F401


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
