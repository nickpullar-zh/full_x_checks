"""
Full Run strategy — runs all registered strategies sequentially, combining
every strategy's output sheets into a single workbook with colour-coded tabs.

The shared Processing Log is built from a single self.log list that each
sub-strategy appends to, producing one continuous log across all strategies.

Tab colours per strategy (Zurich brand palette):
  Conditions  — Zurich Blue  #2167AE
"""

import traceback
from collections import OrderedDict

from strategies.base_strategy import BaseStrategy


# Assign a tab fill colour to each strategy name.
STRATEGY_COLOURS = {
    "Conditions":            "2167AE",  # Zurich Blue
    "Accounting Principles": "23366F",  # Dark Blue
    "X-Checks":              "70AD47",  # Green
    "Grouping By":           "ED7D31",  # Orange
}

STRATEGY_COLOURS_PASTEL = {
    "Conditions":            "BDD7EE",
    "Accounting Principles": "8BAFC7",
    "X-Checks":              "D9EAD3",
    "Grouping By":           "FCE4D6",
}

_FALLBACK_COLOURS = ["9DC3E6", "A9D18E", "FFD966", "BDD7EE", "F4B183"]

_SHEET_PREFIXES = {
    "Conditions":            "Cond",
    "Accounting Principles": "AP",
    "X-Checks":              "XC",
    "Grouping By":           "GB",
}


class FullRun(BaseStrategy):
    """
    Runs every registered strategy (except itself) in sequence and combines
    all output sheets into a single workbook with per-strategy tab colours.
    """

    def process(self, loaded_files: dict, files: dict) -> bool:
        from task_registry import TASK_REGISTRY

        all_sheets: OrderedDict = OrderedDict()
        strategy_sheet_names: dict[str, list[str]] = {}
        strategy_instances: dict[str, object] = {}
        fallback_idx = 0

        # Strategies that produce no output sheets and should not run in Full Run
        _EXCLUDED = {"Full Run", "Collect Live X-Checks"}

        for task_name, (config, strategy_factory) in TASK_REGISTRY.items():
            if task_name in _EXCLUDED:
                continue

            self.log_step(self.log, "Full Run", f"— Starting: {task_name} —", 0)

            # Build per-strategy subsets of files and loaded_files.
            from file_upload_config import FileFieldConfig as _FFC
            strategy_labels = {f.label for f in config.file_fields if isinstance(f, _FFC)}

            strategy_files = dict(files)
            strategy_files["files"] = {
                label: path
                for label, path in files["files"].items()
                if label in strategy_labels
            }
            strategy_files["sheet_names"] = {
                label: sheet
                for label, sheet in files["sheet_names"].items()
                if label in strategy_labels
            }
            strategy_loaded = {
                label: data
                for label, data in loaded_files.items()
                if label in strategy_labels
            }

            # Instantiate and wire up the sub-strategy.
            strategy = strategy_factory(config)
            strategy.log = self.log
            strategy._progress_dialog = self._progress_dialog
            strategy._stop_event = self._stop_event
            # Share the cached COM labeler — only one Excel session for the run.
            strategy._sensitivity_labeler = self._sensitivity_labeler

            # Capture sheets rather than writing a separate file.
            captured_sheets: OrderedDict = OrderedDict()

            def _capture(output_path, sheets, log, summaries=None,
                         _captured=captured_sheets, _name=task_name, _s=strategy):
                _captured.update(sheets)
                _s.log_step(log, "Full Run",
                            f"{_name}: captured {len(sheets)} sheet(s)", len(sheets))

            strategy.write_excel_output = _capture
            # Suppress per-strategy sensitivity labelling; applied once at the end.
            strategy._apply_sensitivity_label = lambda path: None

            try:
                result = strategy.process(strategy_loaded, strategy_files)
            except StopIteration:
                raise
            except Exception as exc:
                self.log_step(self.log, task_name, f"Exception: {exc}", 0)
                self.log_step(self.log, task_name, traceback.format_exc(), 0)
                self.log_step(self.log, "Full Run",
                              f"Aborting Full Run — {task_name} encountered an error.", 0)
                return False

            if result is False:
                self.log_step(self.log, "Full Run",
                              f"Aborting Full Run — {task_name} failed.", 0)
                return False

            if not captured_sheets:
                self.log_step(self.log, "Full Run",
                              f"{task_name}: no sheets produced — skipping", 0)
                continue

            prefix = _SHEET_PREFIXES.get(task_name, task_name[:4])
            tab_names: list[str] = []

            for sheet_name, df in captured_sheets.items():
                combined = _unique_name(f"{prefix} — {sheet_name}", all_sheets)
                all_sheets[combined] = df
                tab_names.append(combined)

            strategy_sheet_names[task_name] = tab_names
            strategy_instances[task_name] = strategy
            self.log_step(self.log, "Full Run",
                          f"{task_name}: complete ({len(tab_names)} sheet(s))", 0)

        if not all_sheets:
            self.log_step(self.log, "Full Run",
                          "No output produced by any strategy.", 0)
            return False

        output_path = self.build_output_path(
            files["output_directory"], "Full Run", files["timestamp"]
        )
        self.log_step(self.log, "Full Run",
                      f"Writing combined workbook ({len(all_sheets)} sheets)", len(all_sheets))

        self._strategy_sheet_names = strategy_sheet_names
        self._strategy_instances = strategy_instances
        self.write_excel_output(output_path, all_sheets, self.log)
        return True

    def apply_output_formatting(self, workbook):
        # 1. Delegate cell-level formatting to each strategy via a prefix shim.
        # Each strategy's apply_output_formatting checks for e.g. "Comparison" by
        # name — the shim maps unprefixed names to the real prefixed sheet names so
        # the strategy never needs to know about prefixes.
        # Tab colours set by the strategy are overwritten in step 2 below.
        strategy_instances = getattr(self, "_strategy_instances", {})
        for task_name, strategy in strategy_instances.items():
            prefix = _SHEET_PREFIXES.get(task_name, task_name[:4])
            shim = _PrefixedWorkbook(workbook, prefix)
            strategy.apply_output_formatting(shim)

        # 2. Tab colours for each strategy group + Processing Log (overrides any
        #    colours set by the strategy delegates above).
        strategy_sheet_names = getattr(self, "_strategy_sheet_names", {})
        fallback_idx = 0
        for task_name, sheet_names in strategy_sheet_names.items():
            colour = STRATEGY_COLOURS.get(task_name)
            pastel = STRATEGY_COLOURS_PASTEL.get(task_name)
            if colour is None:
                colour = _FALLBACK_COLOURS[fallback_idx % len(_FALLBACK_COLOURS)]
                fallback_idx += 1
            for name in sheet_names:
                if name in workbook.sheetnames:
                    if "Comparison" in name:
                        workbook[name].sheet_properties.tabColor = colour
                    else:
                        workbook[name].sheet_properties.tabColor = pastel or colour
        if "Processing Log" in workbook.sheetnames:
            workbook["Processing Log"].sheet_properties.tabColor = "808080"


class _PrefixedWorkbook:
    """
    Shim that wraps a real openpyxl Workbook and makes prefixed sheet names
    (e.g. "XC — Comparison") accessible under their unprefixed base names
    (e.g. "Comparison"). Allows strategy apply_output_formatting methods to
    work unchanged whether called standalone or from Full Run.
    """

    def __init__(self, workbook, prefix: str):
        self._wb = workbook
        self._prefix = prefix + " — "

    def __getitem__(self, sheet_name: str):
        prefixed = self._prefix + sheet_name
        if prefixed in self._wb.sheetnames:
            return self._wb[prefixed]
        if sheet_name in self._wb.sheetnames:
            return self._wb[sheet_name]
        raise KeyError(sheet_name)

    @property
    def sheetnames(self):
        return [
            name[len(self._prefix):] if name.startswith(self._prefix) else name
            for name in self._wb.sheetnames
        ]


def _unique_name(name: str, existing: dict, max_len: int = 31) -> str:
    """Ensure name is unique within existing and within Excel's 31-char sheet limit."""
    candidate = name[:max_len]
    if candidate not in existing:
        return candidate
    for i in range(2, 100):
        suffix = f" ({i})"
        candidate = f"{name[:max_len - len(suffix)]}{suffix}"
        if candidate not in existing:
            return candidate
    return name[:max_len]
