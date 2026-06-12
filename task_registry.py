import importlib

from task_configs import (
    X_CHECKS_UPLOAD_CONFIG,
    ACCOUNTING_PRINCIPLES_UPLOAD_CONFIG,
    CONDITIONS_UPLOAD_CONFIG,
    GROUPING_BY_UPLOAD_CONFIG,
)

# PyInstaller dependency hints — never executed at runtime, but scanned by
# PyInstaller's static analyser so that pandas/openpyxl/numpy are bundled.
if False:
    from strategies.x_checks import XChecks
    from strategies.grouping_by import GroupingBy


def _lazy(module: str, cls: str):
    """Returns a callable that imports and instantiates the strategy class on demand.
    Defers heavy library imports (pandas, openpyxl, numpy) until the user clicks Start."""
    def _factory(*args, **kwargs):
        return getattr(importlib.import_module(module), cls)(*args, **kwargs)
    return _factory


# Registry maps display name → (UI config, processing strategy factory)
# To add a new use case: add one line here and a new config in task_configs.py
TASK_REGISTRY = {
    "X-Checks":           (X_CHECKS_UPLOAD_CONFIG,   _lazy('strategies.x_checks',    'XChecks')),
    #"X-Checks Accounting Principles": (ACCOUNTING_PRINCIPLES_UPLOAD_CONFIG, _lazy('strategies.accounting_principles', 'AccountingPrinciples')),
    #"X-Checks Conditions":            (CONDITIONS_UPLOAD_CONFIG,            _lazy('strategies.conditions',            'Conditions')),
    "X-Checks Grouping By": (GROUPING_BY_UPLOAD_CONFIG, _lazy('strategies.grouping_by', 'GroupingBy')),
}