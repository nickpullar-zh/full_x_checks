"""
Per-strategy file-upload configurations.

This file lives on the infrastructure (`main`) branch with no entries — each
strategy branch (e.g. v0.3-X-Checks, v0.5-Accounting-Principles) adds its own
UploadTaskConfig and registers it in task_registry.py.
"""
from file_upload_config import UploadTaskConfig, FileFieldConfig  # noqa: F401

# Each strategy branch adds its own config here, e.g.:
#
# X_CHECKS_UPLOAD_CONFIG = UploadTaskConfig(
#     task_name="X-Checks",
#     window_title="X-Check Files",
#     requires_output_directory=True,
#     file_fields=[...],
# )
