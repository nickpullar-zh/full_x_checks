from abc import ABC, abstractmethod
from tkinter import messagebox
import pandas as pd
import os


class BaseStrategy(ABC):
    """
    Handles everything that is ALWAYS the same:
    - Loading files into memory
    - Calling the use-case-specific processing
    """

    def execute(self, files: dict, output_directory: str):
        """
        Entry point called by main.py.
        Loads all files then hands off to the subclass.
        """
        print("Loading files into memory...")
        loaded_files = self._load_files(files["files"], files["sheet_names"])

        if loaded_files is None: #← Abort cleanly if a file was locked
            return

        print("Files loaded successfully:")
        for label, data in loaded_files.items():
            print(f"  {label}: {type(data)}")

        # Hand off to the specific use case
        self.process(loaded_files, files, output_directory)

    def _load_files(self, files: dict, sheet_names:dict) -> dict:
        """
        Reads each file into memory based on its extension.
        Returns a dict of {label: DataFrame/content}
        """
        loaded = {}
        for label, path in files.items():
            if path is None:
                continue
            try:
                ext = os.path.splitext(path)[1].lower()  # ← Normalise to lowercase

                if ext == ".xlsx":
                    sheet = sheet_names.get(label, "Sheet1")  # ← Get sheet name
                    loaded[label] = pd.read_excel(path, sheet_name=sheet)
                elif ext == ".csv":
                    loaded[label] = pd.read_csv(path)
                elif ext == ".txt":
                    with open(path, "r") as f:
                        loaded[label] = f.read()
                else:
                    raise ValueError(f"Unsupported file type for: {path}")
            except PermissionError:
                messagebox.showerror(
                    "File In Use",
                    f"'{os.path.basename(path)}' is currently open in another application.\n\n"
                    f"Please close it and try again.\n\n"
                    f"The application will now stop."
                )
                return None  # ← Stop loading, return nothing

            except ValueError as e:
                if "Worksheet" in str(e) or "sheet" in str(e).lower():
                    messagebox.showerror(
                        "Sheet Not Found",
                        f"Could not find sheet '{sheet_names.get(label)}' "
                        f"in '{os.path.basename(path)}'.\n\n"
                        f"Please check the sheet name and try again."
                    )
                else:
                    messagebox.showerror(
                        "Unsupported File Type",
                        f"'{os.path.basename(path)}' is not a supported file type.\n\n"
                        f"Please select an xlsx, csv or txt file."
                    )
                    return None
        return loaded

    @abstractmethod
    def process(self, loaded_files: dict, files: dict, output_directory: str):
        """
        Subclasses implement THIS — not execute().
        By the time this is called, all files are already in memory.
        loaded_files = {label: DataFrame/content}
        files = {label: original file path}
        """
        pass