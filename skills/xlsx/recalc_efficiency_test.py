#!/usr/bin/env python3
"""
Tests for xlsx recalc efficiency improvements.
Run with: python -m pytest recalc_efficiency_test.py -v
"""

import unittest
from unittest.mock import MagicMock, patch
import re


class TestErrorPatternMatching(unittest.TestCase):
    """Test efficient Excel error detection patterns."""

    def test_compiled_regex_matches_all_errors(self):
        """Test that compiled regex matches all Excel error types."""
        from recalc_utils import EXCEL_ERROR_PATTERN

        errors = ['#VALUE!', '#DIV/0!', '#REF!', '#NAME?', '#NULL!', '#NUM!', '#N/A']
        for error in errors:
            self.assertIsNotNone(
                EXCEL_ERROR_PATTERN.search(error),
                f"Pattern should match {error}"
            )

    def test_compiled_regex_matches_errors_in_context(self):
        """Test regex matches errors within cell content."""
        from recalc_utils import EXCEL_ERROR_PATTERN

        test_cases = [
            ('Cell contains #VALUE! error', '#VALUE!'),
            ('Result: #DIV/0!', '#DIV/0!'),
            ('#REF! at start', '#REF!'),
            ('Multiple #NAME? and #NULL! errors', '#NAME?'),  # First match
        ]

        for content, expected in test_cases:
            match = EXCEL_ERROR_PATTERN.search(content)
            self.assertIsNotNone(match, f"Should match error in: {content}")
            self.assertEqual(match.group(), expected)

    def test_compiled_regex_no_false_positives(self):
        """Test regex doesn't match non-error content."""
        from recalc_utils import EXCEL_ERROR_PATTERN

        non_errors = [
            'Normal text',
            '12345',
            '#INVALID',  # Not a real Excel error
            'DIV/0',     # Missing #
            'VALUE!',    # Missing #
            '#value!',   # Wrong case (Excel errors are uppercase)
        ]

        for content in non_errors:
            self.assertIsNone(
                EXCEL_ERROR_PATTERN.search(content),
                f"Should not match: {content}"
            )

    def test_find_all_errors_in_cell(self):
        """Test finding all errors in a cell with multiple errors."""
        from recalc_utils import find_errors_in_value

        content = 'Errors: #VALUE! and #REF! found'
        errors = find_errors_in_value(content)
        self.assertEqual(len(errors), 2)
        self.assertIn('#VALUE!', errors)
        self.assertIn('#REF!', errors)


class TestCellIterationBounds(unittest.TestCase):
    """Test efficient cell iteration with bounds."""

    def test_get_worksheet_bounds(self):
        """Test extracting actual data bounds from worksheet."""
        from recalc_utils import get_worksheet_bounds

        # Mock worksheet
        mock_ws = MagicMock()
        mock_ws.min_row = 1
        mock_ws.max_row = 100
        mock_ws.min_col = 1
        mock_ws.max_col = 26

        bounds = get_worksheet_bounds(mock_ws)
        self.assertEqual(bounds['min_row'], 1)
        self.assertEqual(bounds['max_row'], 100)
        self.assertEqual(bounds['min_col'], 1)
        self.assertEqual(bounds['max_col'], 26)

    def test_get_worksheet_bounds_empty_sheet(self):
        """Test bounds detection for empty worksheet."""
        from recalc_utils import get_worksheet_bounds

        mock_ws = MagicMock()
        mock_ws.min_row = None
        mock_ws.max_row = None
        mock_ws.min_col = None
        mock_ws.max_col = None

        bounds = get_worksheet_bounds(mock_ws)
        self.assertIsNone(bounds)

    def test_iterate_cells_with_bounds(self):
        """Test that cell iteration uses bounds."""
        from recalc_utils import iterate_data_cells

        mock_ws = MagicMock()
        mock_ws.min_row = 1
        mock_ws.max_row = 10
        mock_ws.min_col = 1
        mock_ws.max_col = 5

        # Create mock cells
        mock_cells = []
        for row in range(1, 11):
            row_cells = []
            for col in range(1, 6):
                cell = MagicMock()
                cell.value = f'R{row}C{col}'
                cell.coordinate = f'{chr(64+col)}{row}'
                row_cells.append(cell)
            mock_cells.append(row_cells)

        mock_ws.iter_rows.return_value = mock_cells

        cells = list(iterate_data_cells(mock_ws))

        # Verify iter_rows was called with bounds
        mock_ws.iter_rows.assert_called_once()
        call_kwargs = mock_ws.iter_rows.call_args[1]
        self.assertEqual(call_kwargs['min_row'], 1)
        self.assertEqual(call_kwargs['max_row'], 10)
        self.assertEqual(call_kwargs['min_col'], 1)
        self.assertEqual(call_kwargs['max_col'], 5)


class TestWorkbookLoadingOptimization(unittest.TestCase):
    """Test optimized workbook loading strategies."""

    def test_single_load_extracts_both_values_and_formulas(self):
        """Test that we can extract both values and formulas in single load."""
        from recalc_utils import analyze_workbook_single_pass

        # This tests the concept - actual implementation may vary
        mock_wb = MagicMock()
        mock_ws = MagicMock()
        mock_ws.title = 'Sheet1'
        mock_wb.sheetnames = ['Sheet1']
        mock_wb.__getitem__ = MagicMock(return_value=mock_ws)

        # Create cells with formulas
        cells = []
        cell1 = MagicMock()
        cell1.value = '=SUM(A1:A10)'
        cell1.coordinate = 'B1'
        cells.append([cell1])

        cell2 = MagicMock()
        cell2.value = '#REF!'
        cell2.coordinate = 'B2'
        cells.append([cell2])

        mock_ws.iter_rows.return_value = cells
        mock_ws.min_row = 1
        mock_ws.max_row = 2
        mock_ws.min_col = 1
        mock_ws.max_col = 2

        result = analyze_workbook_single_pass(mock_wb)

        self.assertIn('formula_count', result)
        self.assertIn('errors', result)

    def test_read_only_mode_for_large_files(self):
        """Test that read_only mode is used for analysis."""
        from recalc_utils import should_use_read_only_mode

        # Files over threshold should use read_only
        self.assertTrue(should_use_read_only_mode(file_size_mb=50))
        self.assertTrue(should_use_read_only_mode(file_size_mb=10))

        # Small files can use normal mode
        self.assertFalse(should_use_read_only_mode(file_size_mb=1))


class TestErrorScanOptimization(unittest.TestCase):
    """Test optimized error scanning."""

    def test_early_exit_on_error_found(self):
        """Test that scanning can exit early when error is found."""
        from recalc_utils import scan_cell_for_errors

        # Cell with error should return immediately
        result = scan_cell_for_errors('#VALUE!')
        self.assertEqual(result, '#VALUE!')

    def test_none_value_skipped(self):
        """Test that None values are skipped efficiently."""
        from recalc_utils import scan_cell_for_errors

        result = scan_cell_for_errors(None)
        self.assertIsNone(result)

    def test_numeric_value_skipped(self):
        """Test that numeric values are skipped (can't contain errors)."""
        from recalc_utils import scan_cell_for_errors

        result = scan_cell_for_errors(12345)
        self.assertIsNone(result)

        result = scan_cell_for_errors(3.14159)
        self.assertIsNone(result)

    def test_short_string_optimization(self):
        """Test that very short strings are handled efficiently."""
        from recalc_utils import scan_cell_for_errors

        # Strings shorter than shortest error (#N/A = 4 chars) can't contain errors
        result = scan_cell_for_errors('abc')
        self.assertIsNone(result)


class TestFormulaDetection(unittest.TestCase):
    """Test efficient formula detection."""

    def test_is_formula(self):
        """Test formula detection."""
        from recalc_utils import is_formula

        self.assertTrue(is_formula('=SUM(A1:A10)'))
        self.assertTrue(is_formula('=A1+B1'))
        self.assertTrue(is_formula('=IF(A1>0,1,0)'))

        self.assertFalse(is_formula('Normal text'))
        self.assertFalse(is_formula('12345'))
        self.assertFalse(is_formula(None))
        self.assertFalse(is_formula(123))
        self.assertFalse(is_formula(''))

    def test_formula_startswith_optimization(self):
        """Test that formula check uses efficient startswith."""
        from recalc_utils import is_formula

        # Should be O(1), not O(n)
        long_text = 'A' * 1000000  # 1 million chars
        self.assertFalse(is_formula(long_text))  # Should be instant


if __name__ == '__main__':
    unittest.main()
