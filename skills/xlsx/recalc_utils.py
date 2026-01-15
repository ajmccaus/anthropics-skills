#!/usr/bin/env python3
"""
Utilities for xlsx recalc operations.
"""

import re
from typing import Optional, Any, Dict, Generator

# Compiled regex for Excel error detection
EXCEL_ERROR_PATTERN = re.compile(r'#(?:VALUE!|DIV/0!|REF!|NAME\?|NULL!|NUM!|N/A)')


def scan_cell_for_errors(value: Any) -> Optional[str]:
    """Return the first Excel error found in value, or None."""
    if value is None or not isinstance(value, str) or len(value) < 4:
        return None
    match = EXCEL_ERROR_PATTERN.search(value)
    return match.group() if match else None


def get_worksheet_bounds(ws) -> Optional[Dict[str, int]]:
    """Get actual data bounds of worksheet. Returns None if empty."""
    if ws.min_row is None or ws.max_row is None:
        return None
    return {
        'min_row': ws.min_row,
        'max_row': ws.max_row,
        'min_col': ws.min_col or 1,
        'max_col': ws.max_col or 1,
    }


def iterate_data_cells(ws) -> Generator:
    """Iterate over cells in worksheet's actual data range."""
    bounds = get_worksheet_bounds(ws)
    if bounds is None:
        return
    for row in ws.iter_rows(
        min_row=bounds['min_row'],
        max_row=bounds['max_row'],
        min_col=bounds['min_col'],
        max_col=bounds['max_col']
    ):
        for cell in row:
            yield cell


def is_formula(value: Any) -> bool:
    """Check if value is a formula (starts with =)."""
    return isinstance(value, str) and len(value) > 0 and value[0] == '='
