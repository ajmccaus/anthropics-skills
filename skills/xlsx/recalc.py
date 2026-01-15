#!/usr/bin/env python3
"""
Excel Formula Recalculation Script
Recalculates all formulas in an Excel file using LibreOffice

Cross-platform support: Windows, macOS, Linux
"""

import json
import os
import subprocess
import sys
from pathlib import Path

from openpyxl import load_workbook

# Import cross-platform utilities
from platform_utils import (
    get_libreoffice_macro_dir,
    get_soffice_command,
    run_with_timeout,
)
from recalc_utils import (
    iterate_data_cells,
    scan_cell_for_errors,
    is_formula,
    get_worksheet_bounds,
)


def setup_libreoffice_macro():
    """Setup LibreOffice macro for recalculation if not already configured."""
    try:
        macro_dir = get_libreoffice_macro_dir()
    except EnvironmentError as e:
        print(f"Warning: {e}", file=sys.stderr)
        return False

    macro_file = os.path.join(macro_dir, 'Module1.xba')

    if os.path.exists(macro_file):
        with open(macro_file, 'r') as f:
            if 'RecalculateAndSave' in f.read():
                return True

    # Find soffice command
    soffice = get_soffice_command()
    if not soffice:
        print("Warning: LibreOffice not found. Please install LibreOffice.", file=sys.stderr)
        return False

    if not os.path.exists(macro_dir):
        try:
            run_with_timeout([soffice, '--headless', '--terminate_after_init'], timeout=10)
        except subprocess.TimeoutExpired:
            pass  # Timeout is acceptable for initialization
        os.makedirs(macro_dir, exist_ok=True)

    macro_content = '''<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">
    Sub RecalculateAndSave()
      ThisComponent.calculateAll()
      ThisComponent.store()
      ThisComponent.close(True)
    End Sub
</script:module>'''

    try:
        with open(macro_file, 'w') as f:
            f.write(macro_content)
        return True
    except Exception:
        return False


def recalc(filename, timeout=30):
    """
    Recalculate formulas in Excel file and report any errors.

    Cross-platform: Works on Windows, macOS, and Linux.

    Args:
        filename: Path to Excel file
        timeout: Maximum time to wait for recalculation (seconds)

    Returns:
        dict with error locations and counts
    """
    if not Path(filename).exists():
        return {'error': f'File {filename} does not exist'}

    abs_path = str(Path(filename).absolute())

    if not setup_libreoffice_macro():
        return {'error': 'Failed to setup LibreOffice macro. Ensure LibreOffice is installed.'}

    # Find soffice command (cross-platform)
    soffice = get_soffice_command()
    if not soffice:
        return {'error': 'LibreOffice not found. Please install LibreOffice and ensure it is in PATH.'}

    cmd = [
        soffice, '--headless', '--norestore',
        'vnd.sun.star.script:Standard.Module1.RecalculateAndSave?language=Basic&location=application',
        abs_path
    ]

    # Use cross-platform timeout via subprocess (works on all platforms)
    try:
        result = run_with_timeout(cmd, timeout=timeout)
    except subprocess.TimeoutExpired:
        return {'error': f'Recalculation timed out after {timeout} seconds'}

    if result.returncode != 0:
        error_msg = result.stderr or 'Unknown error during recalculation'
        if 'Module1' in error_msg or 'RecalculateAndSave' not in error_msg:
            return {'error': 'LibreOffice macro not configured properly'}
        else:
            return {'error': error_msg}

    # Check for Excel errors - using optimized single-pass analysis
    try:
        # Load workbook once with data_only=False to get both formulas and values
        wb = load_workbook(filename, data_only=False)

        error_details = {}
        total_errors = 0
        formula_count = 0

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]

            # Skip empty sheets efficiently
            if get_worksheet_bounds(ws) is None:
                continue

            # Use optimized bounded iteration
            for cell in iterate_data_cells(ws):
                value = cell.value

                # Check for formula (O(1) operation)
                if is_formula(value):
                    formula_count += 1

                # Check for errors using compiled regex
                error = scan_cell_for_errors(value)
                if error:
                    if error not in error_details:
                        error_details[error] = []
                    error_details[error].append(f"{sheet_name}!{cell.coordinate}")
                    total_errors += 1

        wb.close()

        # Build result summary
        result = {
            'status': 'success' if total_errors == 0 else 'errors_found',
            'total_errors': total_errors,
            'total_formulas': formula_count,
            'error_summary': {}
        }

        # Add non-empty error categories
        for err_type, locations in error_details.items():
            result['error_summary'][err_type] = {
                'count': len(locations),
                'locations': locations[:20]  # Show up to 20 locations
            }

        return result

    except Exception as e:
        return {'error': str(e)}


def main():
    if len(sys.argv) < 2:
        print("Usage: python recalc.py <excel_file> [timeout_seconds]")
        print("\nRecalculates all formulas in an Excel file using LibreOffice")
        print("\nReturns JSON with error details:")
        print("  - status: 'success' or 'errors_found'")
        print("  - total_errors: Total number of Excel errors found")
        print("  - total_formulas: Number of formulas in the file")
        print("  - error_summary: Breakdown by error type with locations")
        print("    - #VALUE!, #DIV/0!, #REF!, #NAME?, #NULL!, #NUM!, #N/A")
        sys.exit(1)
    
    filename = sys.argv[1]
    timeout = int(sys.argv[2]) if len(sys.argv) > 2 else 30
    
    result = recalc(filename, timeout)
    print(json.dumps(result, indent=2))


if __name__ == '__main__':
    main()