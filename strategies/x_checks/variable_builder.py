"""
Variable Builder

Takes a list of raw variable inputs (fs_accounts + movement_types)
and produces a standardised, formatted variables string.

Used by both fip_extraction.py and ebx_extraction.py to ensure
identical output formatting regardless of source.

Input:
    [
        {'fs_accounts': ['A246', 'A247'], 'movement_types': ['ToM1', 'ToM2']},
        ...
    ]

Output:
    'Name:A246ffToM1ff;FS Account:A246^A247;Movement Types:ToM1^ToM2|...'
"""


def build_variables_string(raw_variables: list[dict]) -> str:
    """
    Main entry point. Converts raw variable data into a standardised
    pipe-delimited string of variable definitions.

    Args:
        raw_variables: List of dicts, each with 'fs_accounts' and
                       'movement_types' keys containing lists of strings.

    Returns:
        Pipe-delimited string of formatted variable definitions.
    """
    parts = []
    for var in raw_variables:
        parts.append(_build_variable_entry(var['fs_accounts'], var['movement_types']))
    return '|'.join(parts)


def _build_variable_entry(fs_accounts: list[str], movement_types: list[str]) -> str:
    """
    Builds a single variable entry string from fs_accounts and movement_types.

    Format:
        Name:<name>;FS Account:<accts>;Movement Types:<types>

    Naming rules:
        - Base name  : first account (sorted)
        - ff suffix  : added if more than one account
        - ToM suffix : added if movement types exist (using first sorted movement type)
        - ff suffix  : added after ToM if more than one movement type
    """
    name        = _build_variable_name(fs_accounts, movement_types)
    accts_str   = _build_accounts_string(fs_accounts)
    types_str   = _build_movement_types_string(movement_types)

    return f"Name:{name};FS Account:{accts_str};Movement Types:{types_str}"


def _build_variable_name(fs_accounts: list[str], movement_types: list[str]) -> str:
    """
    Derives the variable name from accounts and movement types.

    Rules (preserved from original EBXExtraction1.py create_variable()):
        - Start with first sorted account
        - Append 'ff' if more than one account
        - Append 'ToM' + first sorted movement type if movement types exist
        - Append 'ff' if more than one movement type
    """
    clean_accounts = [a for a in fs_accounts if a]
    if not clean_accounts:
        return '<blank>'
    sorted_accounts = sorted(clean_accounts)
    name = sorted_accounts[0][:-2] if sorted_accounts[0].endswith('.0') else sorted_accounts[0]

    if len(clean_accounts) > 1:
        name += 'ff'

    clean_types = [t for t in movement_types if t and t != 'nan']
    if clean_types:
        sorted_types = sorted(clean_types)
        t = sorted_types[0]
        name += 'ToM' + (t[:-2] if t.endswith('.0') else t)
        if len(sorted_types) > 1:
            name += 'ff'

    return name


def _build_accounts_string(fs_accounts: list[str]) -> str:
    """Returns sorted, caret-delimited account string, or '<blank>' if empty."""
    clean = [a for a in fs_accounts if a and a != 'nan']
    if not clean:
        return '<blank>'
    return '^'.join(sorted(clean))


def _build_movement_types_string(movement_types: list[str]) -> str:
    """Returns sorted, caret-delimited movement types string, or '<blank>' if empty."""
    clean = [t for t in movement_types if t and t != 'nan']
    if not clean:
        return '<blank>'
    return '^'.join(sorted(clean))