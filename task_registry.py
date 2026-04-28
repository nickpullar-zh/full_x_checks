from task_configs import (
    X_CHECKS_UPLOAD_CONFIG,
    ACCOUNTING_PRINCIPLES_UPLOAD_CONFIG,
    CONDITIONS_UPLOAD_CONFIG,
    GROUPING_BY_UPLOAD_CONFIG,
)
from strategies.x_checks import XChecks
from strategies.accounting_principles import AccountingPrinciples
from strategies.conditions import Conditions
from strategies.grouping_by import GroupingBy

# Registry maps display name → (UI config, processing strategy)
# To add a new use case: add one line here and a new config in task_configs.py
TASK_REGISTRY = {
    "X-Checks":     (X_CHECKS_UPLOAD_CONFIG,  XChecks),
    "X-Checks Accounting Principles": (ACCOUNTING_PRINCIPLES_UPLOAD_CONFIG,  AccountingPrinciples),
    "X-Checks Conditions":     (CONDITIONS_UPLOAD_CONFIG,   Conditions),
    "X-Checks Grouping By":     (GROUPING_BY_UPLOAD_CONFIG,   GroupingBy),
}