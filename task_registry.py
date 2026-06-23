"""
Maps display name in the Application Selector dropdown → (UploadTaskConfig, strategy factory).

This file lives on the infrastructure (`main`) branch with an empty registry —
each strategy branch (e.g. v0.3-X-Checks, v0.5-Accounting-Principles) adds its
own entries here and a matching config in task_configs.py.
"""
import importlib

from task_configs import CONDITIONS_UPLOAD_CONFIG

# PyInstaller dependency hints — never executed at runtime, but scanned by the
# static analyser so strategy modules are bundled even though they are
# loaded lazily via _lazy(). Each strategy branch adds a matching `if False`
# import alongside its TASK_REGISTRY entry.
if False:
    from strategies.conditions import Conditions  # noqa: F401


def _lazy(module: str, cls: str):
    """Returns a callable that imports and instantiates the strategy class on demand.
    Defers heavy library imports (pandas, openpyxl, numpy) until the user clicks Start."""
    def _factory(*args, **kwargs):
        return getattr(importlib.import_module(module), cls)(*args, **kwargs)
    return _factory


# Registry maps display name → (UI config, processing strategy factory).
# To add a new strategy: append one line and add the matching config in
# task_configs.py and the matching `if False:` import above.
TASK_REGISTRY: dict = {
    "Conditions": (CONDITIONS_UPLOAD_CONFIG, _lazy("strategies.conditions", "Conditions")),
}
