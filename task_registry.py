"""
Maps display name in the Application Selector dropdown → (UploadTaskConfig, strategy factory).

This branch (v0.5-Accounting-Principles) registers exactly one strategy.
"""
import importlib

from task_configs import ACCOUNTING_PRINCIPLES_UPLOAD_CONFIG

# PyInstaller dependency hints — never executed at runtime, but scanned by the
# static analyser so strategy modules are bundled even though they are
# loaded lazily via _lazy().
if False:
    from strategies.accounting_principles import AccountingPrinciples  # noqa: F401


def _lazy(module: str, cls: str):
    """Returns a callable that imports and instantiates the strategy class on demand.
    Defers heavy library imports (pandas, openpyxl, numpy) until the user clicks Start."""
    def _factory(*args, **kwargs):
        return getattr(importlib.import_module(module), cls)(*args, **kwargs)
    return _factory


TASK_REGISTRY: dict = {
    "Accounting Principles": (
        ACCOUNTING_PRINCIPLES_UPLOAD_CONFIG,
        _lazy("strategies.accounting_principles", "AccountingPrinciples"),
    ),
}
