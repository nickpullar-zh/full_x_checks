"""
Maps display name in the Application Selector dropdown → (UploadTaskConfig, strategy factory).

All strategies registered here for v1.0. Full Run is registered last so
_build_full_run_config() sees all other entries.
"""
import importlib

from task_configs import (
    GROUPING_BY_UPLOAD_CONFIG,
    ACCOUNTING_PRINCIPLES_UPLOAD_CONFIG,
    CONDITIONS_UPLOAD_CONFIG,
    _build_full_run_config,
)

# PyInstaller dependency hints — never executed at runtime, but scanned by the
# static analyser so strategy modules are bundled even though they are
# loaded lazily via _lazy().
if False:
    from strategies.grouping_by import GroupingBy                          # noqa: F401
    from strategies.accounting_principles import AccountingPrinciples      # noqa: F401
    from strategies.conditions import Conditions                           # noqa: F401
    from strategies.full_run import FullRun                                # noqa: F401


def _lazy(module: str, cls: str):
    """Returns a callable that imports and instantiates the strategy class on demand.
    Defers heavy library imports (pandas, openpyxl, numpy) until the user clicks Start."""
    def _factory(*args, **kwargs):
        return getattr(importlib.import_module(module), cls)(*args, **kwargs)
    return _factory


TASK_REGISTRY: dict = {
    "Grouping By": (
        GROUPING_BY_UPLOAD_CONFIG,
        _lazy("strategies.grouping_by", "GroupingBy"),
    ),
    "Accounting Principles": (
        ACCOUNTING_PRINCIPLES_UPLOAD_CONFIG,
        _lazy("strategies.accounting_principles", "AccountingPrinciples"),
    ),
    "Conditions": (
        CONDITIONS_UPLOAD_CONFIG,
        _lazy("strategies.conditions", "Conditions"),
    ),
}

# Full Run is registered last so _build_full_run_config sees all other entries.
TASK_REGISTRY["Full Run"] = (
    _build_full_run_config(TASK_REGISTRY),
    _lazy("strategies.full_run", "FullRun"),
)
