#!/usr/bin/env python3
"""
Tests for cross-platform utilities for xlsx skill.
Run with: python -m pytest platform_utils_test.py -v
"""

import os
import sys
import unittest
from pathlib import Path
from unittest.mock import patch, MagicMock


class TestLibreOfficePaths(unittest.TestCase):
    """Test LibreOffice path detection across platforms."""

    def test_get_macro_dir_windows(self):
        """Test macro directory detection on Windows."""
        from platform_utils import get_libreoffice_macro_dir

        with patch('platform.system', return_value='Windows'):
            with patch.dict(os.environ, {'APPDATA': r'C:\Users\TestUser\AppData\Roaming'}):
                result = get_libreoffice_macro_dir()
                self.assertIn('AppData', result)
                self.assertIn('Roaming', result)
                self.assertIn('LibreOffice', result)
                self.assertIn('Standard', result)

    def test_get_macro_dir_darwin(self):
        """Test macro directory detection on macOS."""
        from platform_utils import get_libreoffice_macro_dir

        with patch('platform.system', return_value='Darwin'):
            with patch('os.path.expanduser', return_value='/Users/testuser/Library/Application Support/LibreOffice/4/user/basic/Standard'):
                result = get_libreoffice_macro_dir()
                self.assertIn('Library', result)
                self.assertIn('Application Support', result)
                self.assertIn('LibreOffice', result)

    def test_get_macro_dir_linux(self):
        """Test macro directory detection on Linux."""
        from platform_utils import get_libreoffice_macro_dir

        with patch('platform.system', return_value='Linux'):
            with patch('os.path.expanduser', return_value='/home/testuser/.config/libreoffice/4/user/basic/Standard'):
                result = get_libreoffice_macro_dir()
                self.assertIn('.config', result)
                self.assertIn('libreoffice', result)


class TestSofficeCommand(unittest.TestCase):
    """Test soffice command/path detection across platforms."""

    def test_get_soffice_command_in_path(self):
        """Test when soffice is available in PATH."""
        from platform_utils import get_soffice_command

        with patch('shutil.which', return_value='/usr/bin/soffice'):
            result = get_soffice_command()
            self.assertEqual(result, '/usr/bin/soffice')

    def test_get_soffice_command_windows_program_files(self):
        """Test Windows Program Files detection."""
        from platform_utils import get_soffice_command

        with patch('platform.system', return_value='Windows'):
            with patch('shutil.which', return_value=None):  # Not in PATH
                with patch('os.path.exists') as mock_exists:
                    # Simulate finding LibreOffice in Program Files
                    def exists_side_effect(path):
                        return 'Program Files\\LibreOffice' in path and 'soffice.exe' in path
                    mock_exists.side_effect = exists_side_effect

                    result = get_soffice_command()
                    self.assertIsNotNone(result)
                    self.assertIn('soffice', result.lower())

    def test_get_soffice_command_windows_not_found(self):
        """Test Windows when LibreOffice is not installed."""
        from platform_utils import get_soffice_command

        with patch('platform.system', return_value='Windows'):
            with patch('shutil.which', return_value=None):
                with patch('os.path.exists', return_value=False):
                    result = get_soffice_command()
                    self.assertIsNone(result)

    def test_get_soffice_command_linux_not_found(self):
        """Test Linux when LibreOffice is not installed."""
        from platform_utils import get_soffice_command

        with patch('platform.system', return_value='Linux'):
            with patch('shutil.which', return_value=None):
                result = get_soffice_command()
                self.assertIsNone(result)


class TestSubprocessTimeout(unittest.TestCase):
    """Test cross-platform subprocess timeout handling."""

    def test_run_with_timeout_success(self):
        """Test successful command execution within timeout."""
        from platform_utils import run_with_timeout

        # Use a simple cross-platform command
        result = run_with_timeout(['python', '-c', 'print("hello")'], timeout=5)
        self.assertEqual(result.returncode, 0)
        self.assertIn('hello', result.stdout)

    def test_run_with_timeout_expires(self):
        """Test that timeout is enforced."""
        from platform_utils import run_with_timeout
        import subprocess

        # Command that would take longer than timeout
        with self.assertRaises(subprocess.TimeoutExpired):
            run_with_timeout(['python', '-c', 'import time; time.sleep(10)'], timeout=1)

    def test_run_with_timeout_windows(self):
        """Test timeout works on Windows (mocked)."""
        from platform_utils import run_with_timeout

        with patch('platform.system', return_value='Windows'):
            with patch('subprocess.run') as mock_run:
                mock_run.return_value = MagicMock(returncode=0, stdout='ok', stderr='')
                result = run_with_timeout(['soffice', '--help'], timeout=30)

                # Verify subprocess.run was called with timeout parameter
                mock_run.assert_called_once()
                call_kwargs = mock_run.call_args[1]
                self.assertEqual(call_kwargs['timeout'], 30)


class TestExcelComAvailability(unittest.TestCase):
    """Test Excel COM automation availability detection (Windows only)."""

    def test_is_excel_available_windows_with_excel(self):
        """Test detection when Excel is installed on Windows."""
        from platform_utils import is_excel_com_available

        with patch('platform.system', return_value='Windows'):
            # Mock successful win32com import and Excel detection
            with patch.dict(sys.modules, {'win32com': MagicMock(), 'win32com.client': MagicMock()}):
                with patch('platform_utils._check_excel_installed', return_value=True):
                    result = is_excel_com_available()
                    self.assertTrue(result)

    def test_is_excel_available_windows_without_excel(self):
        """Test detection when Excel is not installed on Windows."""
        from platform_utils import is_excel_com_available

        with patch('platform.system', return_value='Windows'):
            with patch.dict(sys.modules, {'win32com': MagicMock(), 'win32com.client': MagicMock()}):
                with patch('platform_utils._check_excel_installed', return_value=False):
                    result = is_excel_com_available()
                    self.assertFalse(result)

    def test_is_excel_available_non_windows(self):
        """Test that Excel COM is not available on non-Windows platforms."""
        from platform_utils import is_excel_com_available

        with patch('platform.system', return_value='Linux'):
            result = is_excel_com_available()
            self.assertFalse(result)

        with patch('platform.system', return_value='Darwin'):
            result = is_excel_com_available()
            self.assertFalse(result)


class TestPathValidation(unittest.TestCase):
    """Test path length and character validation for Windows."""

    def test_validate_path_length_ok(self):
        """Test that normal paths pass validation."""
        from platform_utils import validate_path_for_platform

        result = validate_path_for_platform('/home/user/documents/file.xlsx')
        self.assertTrue(result['valid'])

    def test_validate_path_length_windows_too_long(self):
        """Test that long paths are flagged on Windows."""
        from platform_utils import validate_path_for_platform

        with patch('platform.system', return_value='Windows'):
            long_path = 'C:\\' + 'a' * 300 + '\\file.xlsx'
            result = validate_path_for_platform(long_path)
            self.assertFalse(result['valid'])
            self.assertIn('length', result['error'].lower())

    def test_validate_path_length_windows_with_prefix(self):
        """Test that long paths with \\?\\ prefix are allowed on Windows."""
        from platform_utils import validate_path_for_platform

        with patch('platform.system', return_value='Windows'):
            long_path = '\\\\?\\C:\\' + 'a' * 300 + '\\file.xlsx'
            result = validate_path_for_platform(long_path)
            self.assertTrue(result['valid'])


if __name__ == '__main__':
    unittest.main()
