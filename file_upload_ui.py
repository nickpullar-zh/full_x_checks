import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Optional, Dict
from file_upload_config import UploadTaskConfig

class FileUploadUI:
    """
    Dynamically builds a file upload dialog from an UploadTaskConfig.
    Handles any number of file fields, optional fields, and an optional
    output directory picker — all driven by configuration.
    """

    def __init__(self, config: UploadTaskConfig, parent: tk.Tk):
        self.config = config
        self.file_paths: Dict[str, tk.StringVar] = {}
        self.sheet_names: Dict[str, tk.StringVar] = {}
        self.sheet_entries: Dict[str, ttk.Entry] = {}
        self.path_labels: Dict[str, ttk.Label] = {}
        self.sheet_labels: Dict[str, ttk.Label] = {}
        self.output_directory = ""
        self.output_label = None
        self.result: Optional[Dict] = None
        self.parent = parent
        self.process_only_differences = tk.BooleanVar(value=False)
        self.extra_checkboxes: dict = {}

        self.root = tk.Toplevel(parent)  # ← Toplevel not Tk()
        self.root.title(config.window_title)
        self.root.resizable(False, False)
        self.root.grab_set()  # ← Modal
        self._build_ui()
        self._set_position()  # ← Position logic
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ==========================================
    # UI Construction
    # ==========================================

    def _on_close(self):
        """
        Called when the user clicks the red X button.
        Destroys the window and exits the application cleanly.
        """
        self.root.destroy()

    def _check_ready(self):
        """
        Enables the Proceed button only when all required
        fields have been filled.
        """
        # Check all required file fields
        for field in self.config.file_fields:
            if field.required and not self.file_paths[field.label].get():
                self.submit_btn.config(state="disabled")
                return

        # Check output directory if required
        if self.config.requires_output_directory and not self.output_directory:
            self.submit_btn.config(state="disabled")
            return

        # All required fields filled — enable the button
        self.submit_btn.config(state="normal")

    def _set_position(self):
        """
        Positions the dialog at the same top-left as the parent window.
        Clamps to usable screen boundaries (respects taskbar) if needed.
        NOTE: Taskbar-aware positioning only works on Windows.
        """
        self.root.update_idletasks()

        parent_x = self.parent.winfo_x()
        parent_y = self.parent.winfo_y()
        window_width = self.root.winfo_reqwidth()
        window_height = self.root.winfo_reqheight()
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # Use Windows API for taskbar-aware positioning, fall back to screen size
        if os.name == 'nt':
            import ctypes
            import ctypes.wintypes

            work_area = ctypes.wintypes.RECT()
            ctypes.windll.user32.SystemParametersInfoW(48, 0, ctypes.byref(work_area), 0)
            usable_height = work_area.bottom  # ← Top of taskbar
            usable_width = work_area.right    # ← Right boundary (accounts for side taskbars too)
        else:
            usable_height = screen_height
            usable_width = screen_width

        desired_x = parent_x
        desired_y = parent_y

        if desired_x + window_width > usable_width:
            desired_x = usable_width - window_width

        # ← Clamp against usable height, not full screen height
        if desired_y + window_height > usable_height:
            desired_y = usable_height - window_height

        desired_x = max(0, desired_x)
        desired_y = max(0, desired_y)

        self.root.geometry(f"+{desired_x}+{desired_y}")

    def _build_ui(self):
        """Dynamically builds UI rows from config."""
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.grid(row=0, column=0, sticky="nsew")

        # ==========================================
        # Constants
        # ==========================================
        HINT_WRAP_LENGTH = 800  # ← Max width before hint text wraps
        LABEL_FALLBACK_WIDTH = 400  # ← Fallback if no hints exist

        current_row = 0

        # --- Title ---
        ttk.Label(
            main_frame,
            text=self.config.task_name,
            font=("Helvetica", 14, "bold")
        ).grid(row=current_row, column=0, columnspan=3, pady=(0, 15))
        current_row += 1

        ttk.Separator(main_frame, orient="horizontal").grid(
            row=current_row, column=0, columnspan=3, sticky="ew", pady=5
        )
        current_row += 1

        # ==========================================
        # Pass 1 — Build all rows, store hint label
        # references so we can measure them later
        # ==========================================
        hint_labels = []

        for field in self.config.file_fields:
            path_var = tk.StringVar()
            self.file_paths[field.label] = path_var

            # --- Field label ---
            label_text = (
                f"{field.label} *" if field.required
                else f"{field.label} (optional)"
            )
            ttk.Label(
                main_frame,
                text=label_text,
                wraplength=150,
                justify="left",
                anchor="w"
            ).grid(row=current_row, column=0, padx=5, pady=(8, 0), sticky="w")

            # --- Path display label (placeholder for now) ---
            path_label = ttk.Label(
                main_frame,
                text="No file selected",
                foreground="grey",
                justify="left",
                anchor="w"
            )
            path_label.grid(row=current_row, column=1, padx=5, pady=(8, 0), sticky="w")
            self.path_labels[field.label] = path_label

            # --- Browse button ---
            ttk.Button(
                main_frame,
                text="Browse...",
                command=lambda f=field, v=path_var: self._browse_file(f, v)
            ).grid(row=current_row, column=2, padx=5, pady=(8, 0))

            current_row += 1

            # --- Sheet name row ---
            if field.show_sheet:
                sheet_var = tk.StringVar(value=field.default_sheet)  # ← Default value
                self.sheet_names[field.label] = sheet_var

                sheet_label = ttk.Label(
                    main_frame,
                    text="Sheet name:",
                    font=("Helvetica", 8),
                    foreground="grey"
                )
                sheet_label.grid(row=current_row, column=0, padx=5, pady=(0, 4), sticky="e")
                self.sheet_labels[field.label] = sheet_label  # ← Store reference

                sheet_entry = ttk.Entry(
                    main_frame,
                    textvariable=sheet_var,
                    width=30,
                    font=("Helvetica", 8),
                    state="disabled"
                )
                sheet_entry.grid(row=current_row, column=1, padx=5, pady=(0, 4), sticky="w")
                self.sheet_entries[field.label] = sheet_entry

                current_row += 1

            # --- Hint label ---
            hint_text = f"  {field.description}" if field.description else ""
            hint_label = ttk.Label(
                main_frame,
                text=hint_text,
                foreground="black",
                font=("Helvetica", 8),
                wraplength=HINT_WRAP_LENGTH,  # ← Max width before wrapping
                justify="left"
            )
            hint_label.grid(row=current_row, column=1, sticky="w", pady=(0, 4))
            hint_labels.append(hint_label)  # ← Store for measuring later

            current_row += 1

        ttk.Separator(main_frame, orient="horizontal").grid(
            row=current_row, column=0, columnspan=3, sticky="ew", pady=5
        )
        current_row += 1

        # --- Output Directory ---
        if self.config.requires_output_directory:
            ttk.Label(
                main_frame,
                text="Output Directory *",
                wraplength=150,
                justify="left",
                anchor="w"
            ).grid(row=current_row, column=0, padx=5, pady=(8, 0), sticky="w")

            self.output_label = ttk.Label(
                main_frame,
                text="No directory selected",
                foreground="grey",
                justify="left",
                anchor="w"
            )
            self.output_label.grid(row=current_row, column=1, padx=5, pady=(8, 0), sticky="w")

            ttk.Button(
                main_frame,
                text="Browse...",
                command=self._browse_directory
            ).grid(row=current_row, column=2, padx=5, pady=(8, 0))

            current_row += 1

            # --- Output directory hint ---
            output_hint_label = ttk.Label(
                main_frame,
                text="  Folder where output files will be saved",
                foreground="black",
                font=("Helvetica", 8),
                wraplength=HINT_WRAP_LENGTH,  # ← Max width before wrapping
                justify="left"
            )
            output_hint_label.grid(row=current_row, column=1, sticky="w", pady=(0, 4))
            hint_labels.append(output_hint_label)  # ← Include in measurement

            current_row += 1

        ttk.Separator(main_frame, orient="horizontal").grid(
            row=current_row, column=0, columnspan=3, sticky="ew", pady=5
        )
        current_row += 1

        # --- Process Only Differences Checkbox ---
        ttk.Checkbutton(
            main_frame,
            text="Process only differences",
            variable=self.process_only_differences
        ).grid(row=current_row, column=0, columnspan=3, pady=(8, 4))
        current_row += 1

        # --- Config-driven checkboxes (e.g. experimental options) ---
        for cb in self.config.checkboxes:
            var = tk.BooleanVar(value=cb.get("default", False))
            self.extra_checkboxes[cb["key"]] = var
            ttk.Checkbutton(
                main_frame,
                text=cb["label"],
                variable=var
            ).grid(row=current_row, column=0, columnspan=3, pady=(2, 2))
            current_row += 1

        ttk.Separator(main_frame, orient="horizontal").grid(
            row=current_row, column=0, columnspan=3, sticky="ew", pady=5
        )
        current_row += 1

        # --- Proceed Button ---
        self.submit_btn = ttk.Button(
            main_frame,
            text="Proceed",
            command=self._on_submit
        )
        self.submit_btn.grid(row=current_row, column=0, columnspan=3, pady=15)
        self.submit_btn.config(state="disabled")

        # ==========================================
        # Pass 2 — Measure hint labels AFTER render
        # and apply the widest as the uniform width
        # for all path/directory display labels
        # ==========================================
        self.root.update_idletasks()  # ← Force Tkinter to render

        # Find the widest hint label
        max_hint_width = max(
            (label.winfo_width() for label in hint_labels if label.winfo_width() > 1),
            default=LABEL_FALLBACK_WIDTH
        )

        # Apply uniform width to all path display labels
        for path_label in self.path_labels.values():
            path_label.config(wraplength=max_hint_width)

        # Apply same width to output directory label
        if self.output_label:
            self.output_label.config(wraplength=max_hint_width)


    # ==========================================
    # Event Handlers
    # ==========================================

    def _browse_file(self, field, path_var: tk.StringVar):
        """Opens a file picker for a specific field."""
        filepath = filedialog.askopenfilename(
            title=f"Select {field.label}",
            filetypes=field.file_types + [("All Files", "*.*")]
        )
        if filepath:
            path_var.set(filepath)
            # Update the label directly — no StringVar/Entry issues
            self.path_labels[field.label].config(
                text=os.path.basename(filepath),  # ← Show filename only, not full path
                foreground="black"  # ← Change from grey to black
            )
            if field.show_sheet:
                self.sheet_labels[field.label].config(foreground="black")
                self.sheet_entries[field.label].config(state="normal")  # ← Enable on file select
            self._check_ready()

    def _browse_directory(self):
        """Opens a directory picker for the output location."""
        directory = filedialog.askdirectory(
            title="Select Output Directory"
        )
        if directory:
            self.output_directory = directory  # ← Store as plain string
            if self.output_label is not None:
                self.output_label.config(
                    text=directory,  # ← Show full path for directory
                    foreground="black"  # ← Change from grey to black
                )
            self._check_ready()

    def _on_submit(self):
        
        from datetime import datetime

        """All validation already handled by _check_ready() — just package results."""
        
        # Set the global timestamp once here
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        #app_state.process_only_differences = self.process_only_differences.get()

        self.result = {
            "files": {
                label: var.get() or None
                for label, var in self.file_paths.items()
            },
            "sheet_names": {
                label: var.get() or "Sheet1"
                for label, var in self.sheet_names.items()
            },
            "output_directory": self.output_directory or None,
            "timestamp": timestamp,  # Reference the local timestamp here
            "process_only_differences": self.process_only_differences.get()
        }
        for key, var in self.extra_checkboxes.items():
            self.result[key] = var.get()
        self.root.destroy()

    # ==========================================
    # Public Interface
    # ==========================================

    def run(self) -> Optional[Dict]:
        """
        Launch the UI and return a dict when complete:
        {
            "files": {"Label": "path/to/file" or None},
            "output_directory": "path/to/dir" or None
        }
        Returns None if the user closes the window without submitting.
        """
        self.root.wait_window()
        return self.result