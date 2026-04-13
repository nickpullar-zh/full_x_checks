import tkinter as tk
from tkinter import ttk
from file_upload_ui import FileUploadUI
from task_registry import TASK_REGISTRY

# This is v0.2-Grouping_by

class TaskSelectorUI:
    """
    Simple launcher window — user picks a task,
    then the appropriate upload dialog is shown.
    """

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("X_Check Application")
        self.root.resizable(False, False)
        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _on_close(self):
        """
        Called when the user clicks the red X on the launcher.
        Exits the application cleanly.
        """
        self.root.destroy()
        exit()

    def _on_task_selected(self, *args):
        """
        Called whenever the dropdown selection changes.
        Enables the Start button only when a task is selected.
        """
        if self.task_var.get():
            self.start_btn.config(state="normal")
        else:
            self.start_btn.config(state="disabled")

    def _build_ui(self):
        frame = ttk.Frame(self.root, padding="20")
        frame.grid(row=0, column=0)

        ttk.Label(
            frame,
            text="X-Check Application",
            font=("Helvetica", 16, "bold")
        ).grid(row=0, column=0, columnspan=2, pady=(0, 20))

        ttk.Label(
            frame,
            text="Select a task to begin:"
        ).grid(row=1, column=0, sticky="w", pady=5)

        self.task_var = tk.StringVar()
        self.task_var.trace_add("write", self._on_task_selected)
        task_dropdown = ttk.Combobox(
            frame,
            textvariable=self.task_var,
            values=list(TASK_REGISTRY.keys()),
            state="readonly",
            width=30
        )
        task_dropdown.grid(row=1, column=1, padx=10, pady=5)

        print("v0.2-Grouping_By")

        self.start_btn = ttk.Button(
            frame,
            text="Start",
            command=self._on_start
        )
        self.start_btn.grid(row=2, column=0, columnspan=2, pady=15)
        self.start_btn.config(state="disabled")  # ← Disabled by default

    def _on_start(self):
        task_name = self.task_var.get()
        if not task_name:
            return

        config, strategy_class = TASK_REGISTRY[task_name]

        # Hide launcher while upload dialog is open
        self.root.withdraw()

        files = FileUploadUI(config, self.root).run()

        if files:
            strategy = strategy_class()
            strategy.execute(files, files["output_directory"])  # ← Pass full files dict

        self.root.destroy()  # ← Closes launcher and exits
        # Return to launcher after task completes
        #self.root.deiconify()

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    TaskSelectorUI().run()