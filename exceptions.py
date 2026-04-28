"""
Custom exceptions for file loading and processing errors.
These allow the data layer to signal problems without
depending on Tkinter or any other UI framework.
"""

class FileLoadError(Exception):
    """Raised when a file cannot be opened or read."""
    pass

class SheetNotFoundError(Exception):
    """Raised when a required sheet is missing from an Excel file."""
    pass

class MissingColumnsError(Exception):
    """Raised when required columns are not found in a file."""
    pass

class UnsupportedFileTypeError(Exception):
    """Raised when a file has an unrecognised extension."""
    pass