import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


def _exe_dir() -> str:
    """Return the folder the EXE (or script) lives in."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


# Tracks the last folder a file dialog visited. Starts at the EXE folder;
# updated whenever the user picks a file or folder so the next dialog opens there.
_last_dir: str = _exe_dir()


def _get_initial_dir() -> str:
    return _last_dir


def _set_last_dir(path: str) -> None:
    """Update the remembered directory when a FILE is selected. Folder picks are ignored."""
    global _last_dir
    if os.path.isfile(path):
        _last_dir = os.path.dirname(path)
from typing import Optional, Dict
from file_upload_config import UploadTaskConfig, FileFieldConfig, SectionConfig


class _Tooltip:
    """Shows a floating tooltip window when the user hovers over a widget."""

    DELAY_MS = 600   # ms before tooltip appears
    WRAP_PX  = 320   # max tooltip width before wrapping

    def __init__(self, widget: tk.Widget, text: str):
        self._widget  = widget
        self._text    = text
        self._win: Optional[tk.Toplevel] = None
        self._after_id = None
        widget.bind("<Enter>", self._on_enter)
        widget.bind("<Leave>", self._on_leave)

    def _on_enter(self, event=None):
        self._after_id = self._widget.after(self.DELAY_MS, self._show)

    def _on_leave(self, event=None):
        if self._after_id:
            self._widget.after_cancel(self._after_id)
            self._after_id = None
        self._hide()

    def _show(self):
        if self._win:
            return
        x = self._widget.winfo_rootx() + 20
        y = self._widget.winfo_rooty() + self._widget.winfo_height() + 4
        self._win = tk.Toplevel(self._widget)
        self._win.wm_overrideredirect(True)   # no title bar or borders
        self._win.wm_geometry(f"+{x}+{y}")
        tk.Label(
            self._win,
            text=self._text,
            justify="left",
            background="#DDE4E3",   # Dove (brand)
            foreground="#23366F",   # Dark Blue (accessible on Dove per brand guidelines)
            relief="solid",
            borderwidth=1,
            wraplength=self.WRAP_PX,
            padx=6,
            pady=4,
        ).pack()

    def _hide(self):
        if self._win:
            self._win.destroy()
            self._win = None


class FileUploadUI:
    """
    Dynamically builds a file upload dialog from an UploadTaskConfig.
    Handles any number of file fields, optional fields, and an optional
    output directory picker — all driven by configuration.
    """

    def __init__(self, config: UploadTaskConfig, parent: tk.Tk, prefill: Optional[Dict] = None):
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
        self.process_only_differences = tk.BooleanVar(value=True)  # v0.4: default ON
        self.extra_checkboxes: dict = {}

        self.root = tk.Toplevel(parent)  # ← Toplevel not Tk()
        self.root.title(config.window_title)
        self.root.resizable(False, False)
        self.root.grab_set()  # ← Modal
        self._build_ui()
        if prefill:
            self._apply_prefill(prefill)
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
        # Check all required file fields (skip SectionConfig dividers)
        for field in self.config.file_fields:
            if not isinstance(field, FileFieldConfig):
                continue
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
        Positions the dialog at the same top-left as the parent window,
        keeping a _SCREEN_MARGIN px gap from every usable-screen edge.
        """
        self.root.update_idletasks()

        M = self._SCREEN_MARGIN
        parent_x = self.parent.winfo_x()
        parent_y = self.parent.winfo_y()
        window_width  = self.root.winfo_reqwidth()
        window_height = self.root.winfo_reqheight()

        if os.name == 'nt':
            import ctypes, ctypes.wintypes
            wa = ctypes.wintypes.RECT()
            ctypes.windll.user32.SystemParametersInfoW(48, 0, ctypes.byref(wa), 0)
            usable_top    = wa.top
            usable_left   = wa.left
            usable_bottom = wa.bottom
            usable_right  = wa.right
        else:
            usable_top  = usable_left = 0
            usable_right  = self.root.winfo_screenwidth()
            usable_bottom = self.root.winfo_screenheight()

        # Clamp with margin: prefer parent position, stay within usable area
        max_x = usable_right  - window_width  - M
        max_y = usable_bottom - window_height - M
        min_x = usable_left  + M
        min_y = usable_top   + M

        desired_x = max(min_x, min(parent_x, max_x))
        desired_y = max(min_y, min(parent_y, max_y))

        self.root.geometry(f"+{desired_x}+{desired_y}")

    # Minimum margin (px) to keep between the dialog edge and the screen edge.
    _SCREEN_MARGIN = 20

    def _build_ui(self):
        """
        Builds the upload form inside a scrollable canvas so it always fits
        on screen regardless of how many fields the strategy requires.
        The title, scrollable fields area, and controls (checkboxes / Proceed)
        are laid out in a fixed outer frame; only the fields area scrolls.
        """
        HINT_WRAP_LENGTH    = 533
        LABEL_FALLBACK_WIDTH = 267

        # ── Outer container fills the window ──────────────────────────────────
        outer = ttk.Frame(self.root, padding="15")
        outer.grid(row=0, column=0, sticky="nsew")
        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)
        outer.rowconfigure(1, weight=1)   # row 1 = scrollable area
        outer.columnconfigure(0, weight=1)

        # ── Title ─────────────────────────────────────────────────────────────
        title_frame = ttk.Frame(outer)
        title_frame.grid(row=0, column=0, sticky="ew", pady=(0, 4))
        ttk.Label(
            title_frame,
            text=self.config.task_name,
            font=("Zurich Sans Semibold", 14)
        ).pack()
        ttk.Separator(outer, orient="horizontal").grid(
            row=0, column=0, sticky="ew", pady=(44, 0)
        )

        # ── Scrollable canvas for file fields ─────────────────────────────────
        canvas_frame = ttk.Frame(outer)
        canvas_frame.grid(row=1, column=0, sticky="nsew")
        canvas_frame.rowconfigure(0, weight=1)
        canvas_frame.columnconfigure(0, weight=1)

        canvas = tk.Canvas(canvas_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical",
                                  command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        # Inner frame rendered inside the canvas
        inner = ttk.Frame(canvas)
        canvas_window = canvas.create_window((0, 0), window=inner, anchor="nw")

        def _on_inner_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_configure(event):
            canvas.itemconfig(canvas_window, width=event.width)

        inner.bind("<Configure>", _on_inner_configure)
        canvas.bind("<Configure>", _on_canvas_configure)

        # Mousewheel scrolling (Windows)
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        # ── Build fields into inner frame ─────────────────────────────────────
        main_frame = inner   # field-building code uses main_frame below

        # ── Build all fields into the scrollable inner frame ──────────────────
        hint_labels: list = []
        current_row = 0

        for field in self.config.file_fields:

            # --- Section divider / title ---
            if isinstance(field, SectionConfig):
                ttk.Separator(main_frame, orient="horizontal").grid(
                    row=current_row, column=0, columnspan=3, sticky="ew", pady=(10, 2)
                )
                current_row += 1
                if field.title:
                    ttk.Label(
                        main_frame,
                        text=field.title,
                        font=("Zurich Sans Semibold", 10),
                    ).grid(row=current_row, column=0, columnspan=3, sticky="w",
                           padx=5, pady=(0, 4))
                    current_row += 1
                continue

            path_var = tk.StringVar()
            self.file_paths[field.label] = path_var

            label_text = (
                f"{field.label} *" if field.required
                else f"{field.label}\n(optional)"
            )
            ttk.Label(
                main_frame,
                text=label_text,
                wraplength=180,
                justify="left",
                anchor="w"
            ).grid(row=current_row, column=0, padx=5, pady=(8, 0), sticky="w")

            path_label = ttk.Label(
                main_frame,
                text="No file selected",
                foreground="grey",
                justify="left",
                anchor="w"
            )
            path_label.grid(row=current_row, column=1, padx=5, pady=(8, 0), sticky="w")
            self.path_labels[field.label] = path_label

            ttk.Button(
                main_frame,
                text="Browse...",
                command=lambda f=field, v=path_var: self._browse_file(f, v)
            ).grid(row=current_row, column=2, padx=5, pady=(8, 0))
            current_row += 1

            # --- Sheet name row ---
            if field.show_sheet:
                sheet_var = tk.StringVar(value=field.default_sheet)
                self.sheet_names[field.label] = sheet_var

                sheet_label = ttk.Label(
                    main_frame,
                    text="Sheet name:",
                    font=("Zurich Sans", 9),
                    foreground="grey"
                )
                sheet_label.grid(row=current_row, column=0, padx=5,
                                 pady=(0, 4), sticky="e")
                self.sheet_labels[field.label] = sheet_label

                sheet_combo = ttk.Combobox(
                    main_frame,
                    textvariable=sheet_var,
                    width=28,
                    font=("Zurich Sans", 9),
                    state="disabled"
                )
                sheet_combo.grid(row=current_row, column=1, padx=5,
                                 pady=(0, 4), sticky="w")
                self.sheet_entries[field.label] = sheet_combo
                current_row += 1

                if field.sheet_note:
                    ttk.Label(
                        main_frame,
                        text=f"  {field.sheet_note}",
                        foreground="grey",
                        font=("Zurich Sans", 9),
                        wraplength=HINT_WRAP_LENGTH,
                        justify="left",
                    ).grid(row=current_row, column=1, sticky="w", pady=(0, 2))
                    current_row += 1

            # --- Hint label ---
            hint_text = f"  {field.description}" if field.description else ""
            hint_label = ttk.Label(
                main_frame,
                text=hint_text,
                foreground="black",
                font=("Zurich Sans", 9),
                wraplength=HINT_WRAP_LENGTH,
                justify="left"
            )
            hint_label.grid(row=current_row, column=1, sticky="w", pady=(0, 4))
            hint_labels.append(hint_label)
            current_row += 1

        # ── Output Directory (inside scrollable area) ─────────────────────────
        ttk.Separator(main_frame, orient="horizontal").grid(
            row=current_row, column=0, columnspan=3, sticky="ew", pady=5
        )
        current_row += 1

        if self.config.requires_output_directory:
            ttk.Label(
                main_frame,
                text="Output Directory *",
                wraplength=180,
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
            self.output_label.grid(row=current_row, column=1, padx=5,
                                   pady=(8, 0), sticky="w")

            ttk.Button(
                main_frame,
                text="Browse...",
                command=self._browse_directory
            ).grid(row=current_row, column=2, padx=5, pady=(8, 0))
            current_row += 1

            output_hint_label = ttk.Label(
                main_frame,
                text="  Folder where output files will be saved",
                foreground="black",
                font=("Zurich Sans", 9),
                wraplength=HINT_WRAP_LENGTH,
                justify="left"
            )
            output_hint_label.grid(row=current_row, column=1, sticky="w", pady=(0, 4))
            hint_labels.append(output_hint_label)
            current_row += 1

        # ── Controls row (fixed below the canvas) ─────────────────────────────
        controls = ttk.Frame(outer, padding=(0, 8, 0, 0))
        controls.grid(row=2, column=0, sticky="ew")

        ttk.Separator(controls, orient="horizontal").grid(
            row=0, column=0, columnspan=3, sticky="ew", pady=(0, 5)
        )

        # Process Only Differences Checkbox
        ttk.Checkbutton(
            controls,
            text="Process only differences",
            variable=self.process_only_differences
        ).grid(row=1, column=0, columnspan=3, pady=(4, 2))
        ctrl_row = 2

        for cb in self.config.checkboxes:
            var = tk.BooleanVar(value=cb.get("default", False))
            self.extra_checkboxes[cb["key"]] = var
            btn = ttk.Checkbutton(controls, text=cb["label"], variable=var)
            btn.grid(row=ctrl_row, column=0, columnspan=3, pady=(2, 2))
            if cb.get("tooltip"):
                _Tooltip(btn, cb["tooltip"])
            ctrl_row += 1

        ttk.Separator(controls, orient="horizontal").grid(
            row=ctrl_row, column=0, columnspan=3, sticky="ew", pady=5
        )
        ctrl_row += 1

        self.submit_btn = ttk.Button(
            controls, text="Proceed", command=self._on_submit
        )
        self.submit_btn.grid(row=ctrl_row, column=0, columnspan=3, pady=(8, 4))
        self.submit_btn.config(state="disabled")
        ctrl_row += 1

        from version import __version__
        tk.Label(
            controls,
            text=f"v{__version__}",
            font=("Zurich Sans", 8),
            foreground="#999999",
        ).grid(row=ctrl_row, column=0, columnspan=3, sticky="w", pady=(4, 0))

        # ── Size canvas to usable screen, respecting margin ───────────────────
        self.root.update_idletasks()
        M = self._SCREEN_MARGIN
        if os.name == 'nt':
            import ctypes, ctypes.wintypes
            wa = ctypes.wintypes.RECT()
            ctypes.windll.user32.SystemParametersInfoW(48, 0, ctypes.byref(wa), 0)
            usable_h = wa.bottom - wa.top
            usable_w = wa.right  - wa.left
        else:
            usable_h = self.root.winfo_screenheight()
            usable_w = self.root.winfo_screenwidth()

        # Height the canvas may occupy = screen minus margins minus non-scroll rows
        controls_h = controls.winfo_reqheight()
        title_h    = 60  # title label + separator
        max_canvas_h = usable_h - 2 * M - controls_h - title_h
        # Natural height of the inner content
        content_h  = inner.winfo_reqheight()
        canvas_h   = min(content_h, max(200, max_canvas_h))
        canvas.configure(height=canvas_h)

        # ── Pass 2 — uniform hint/path label widths ───────────────────────────
        max_hint_width = max(
            (label.winfo_width() for label in hint_labels if label.winfo_width() > 1),
            default=LABEL_FALLBACK_WIDTH
        )
        for path_label in self.path_labels.values():
            path_label.config(wraplength=max_hint_width)
        if self.output_label:
            self.output_label.config(wraplength=max_hint_width)
        for hint_label in hint_labels:
            hint_label.config(wraplength=max_hint_width)

    def _apply_prefill(self, prefill: dict):
        """Restore a previous run's inputs into the form fields."""
        # File paths
        for label, path in (prefill.get("files") or {}).items():
            if path and label in self.file_paths:
                self.file_paths[label].set(path)
                self.path_labels[label].config(
                    text=os.path.basename(path), foreground="black"
                )
                for field in self.config.file_fields:
                    if field.label == label and field.show_sheet:
                        combo = self.sheet_entries[label]
                        if not field.sheet_note:
                            self.sheet_labels[label].config(foreground="black")
                            try:
                                import pandas as pd
                                sheet_names = pd.ExcelFile(path).sheet_names
                                combo["values"] = sheet_names
                            except Exception:
                                pass
                            combo.config(state="normal" if field.sheet_editable else "readonly")

        # Sheet names
        for label, sheet in (prefill.get("sheet_names") or {}).items():
            if sheet and label in self.sheet_names:
                self.sheet_names[label].set(sheet)

        # Output directory
        output_dir = prefill.get("output_directory")
        if output_dir:
            self.output_directory = output_dir
            if self.output_label is not None:
                self.output_label.config(text=output_dir, foreground="black")

        # Checkboxes
        self.process_only_differences.set(prefill.get("process_only_differences", True))
        for key, var in self.extra_checkboxes.items():
            if key in prefill:
                var.set(prefill[key])

        self._check_ready()

    # ==========================================
    # Event Handlers
    # ==========================================

    def _browse_file(self, field, path_var: tk.StringVar):
        """Opens a file picker for a specific field."""
        filepath = filedialog.askopenfilename(
            title=f"Select {field.label}",
            filetypes=field.file_types + [("All Files", "*.*")],
            initialdir=_get_initial_dir(),
        )
        if filepath:
            _set_last_dir(filepath)
            path_var.set(filepath)
            self.path_labels[field.label].config(
                text=os.path.basename(filepath),
                foreground="black"
            )
            if field.show_sheet:
                combo = self.sheet_entries[field.label]
                if field.sheet_note:
                    # Sheet is fixed — never populate or enable the combobox
                    pass
                else:
                    self.sheet_labels[field.label].config(foreground="black")
                    # Read sheet names from the workbook
                    try:
                        import pandas as pd
                        sheet_names = pd.ExcelFile(filepath).sheet_names
                    except Exception:
                        sheet_names = []
                    if sheet_names:
                        combo["values"] = sheet_names
                        # Use default if present, else first sheet
                        default = field.default_sheet
                        selected = default if default in sheet_names else sheet_names[0]
                        self.sheet_names[field.label].set(selected)
                    # Editable fields allow free-text entry; others are read-only
                    combo.config(state="normal" if field.sheet_editable else "readonly")
            self._check_ready()

    def _browse_directory(self):
        """Opens a directory picker for the output location."""
        directory = filedialog.askdirectory(
            title="Select Output Directory",
            initialdir=_get_initial_dir(),
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