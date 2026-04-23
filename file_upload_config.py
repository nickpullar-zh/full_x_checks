from dataclasses import dataclass, field
from typing import List, Tuple, Optional


@dataclass
class FileFieldConfig:
    """Configuration for a single file input field."""
    label: str                              # Display label for the file input
    file_types: List[Tuple[str, str]]       # e.g. [("Excel files", "*.xlsx")]
    required: bool = True
    description: str = ""                   # Helper text shown below the field
    default_sheet: str = "Sheet1"
    required_columns: Optional[list[str]] = None  # None means keep all columns

    @property
    def show_sheet(self) -> bool:
        """Only show sheet name input for Excel files."""
        return any(
            ext.lower().endswith(".xlsx") or ext.lower().endswith(".xls")
            for _, ext in self.file_types
        )

@dataclass
class UploadTaskConfig:
    """Configuration for a complete upload task (one use case)."""
    task_name: str                          # e.g. "X-Checks Publication File"
    file_fields: List[FileFieldConfig] = field(default_factory=list)
    window_title: str = "File Upload"
    requires_output_directory: bool = True  # Most tasks will need this