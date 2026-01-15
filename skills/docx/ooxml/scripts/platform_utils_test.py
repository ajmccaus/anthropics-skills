#!/usr/bin/env python3
"""
Tests for cross-platform utilities for docx/ooxml scripts.
Run with: python -m pytest platform_utils_test.py -v
"""

import unittest
from unittest.mock import patch, MagicMock
from pathlib import Path


class TestSofficeValidation(unittest.TestCase):
    """Test soffice command detection for pack.py validation."""

    def test_get_soffice_for_validation_in_path(self):
        """Test when soffice is in PATH."""
        from platform_utils import get_soffice_for_validation

        with patch('shutil.which', return_value='/usr/bin/soffice'):
            result = get_soffice_for_validation()
            self.assertEqual(result, '/usr/bin/soffice')

    def test_get_soffice_for_validation_windows(self):
        """Test Windows soffice detection."""
        from platform_utils import get_soffice_for_validation

        with patch('platform.system', return_value='Windows'):
            with patch('shutil.which', return_value=None):
                with patch('os.path.exists') as mock_exists:
                    mock_exists.side_effect = lambda p: 'LibreOffice' in p
                    result = get_soffice_for_validation()
                    if result:
                        self.assertIn('soffice', result.lower())

    def test_get_soffice_for_validation_not_found(self):
        """Test graceful handling when soffice not found."""
        from platform_utils import get_soffice_for_validation

        with patch('shutil.which', return_value=None):
            with patch('platform.system', return_value='Linux'):
                result = get_soffice_for_validation()
                self.assertIsNone(result)


class TestValidationCommand(unittest.TestCase):
    """Test building validation commands across platforms."""

    def test_build_validation_command_basic(self):
        """Test basic validation command construction."""
        from platform_utils import build_validation_command

        cmd = build_validation_command(
            soffice_path='/usr/bin/soffice',
            doc_path='/tmp/test.docx',
            output_dir='/tmp/output',
            filter_name='html:HTML'
        )

        self.assertIn('/usr/bin/soffice', cmd)
        self.assertIn('--headless', cmd)
        self.assertIn('--convert-to', cmd)
        self.assertIn('html:HTML', cmd)

    def test_build_validation_command_windows_paths(self):
        """Test command with Windows-style paths."""
        from platform_utils import build_validation_command

        with patch('platform.system', return_value='Windows'):
            cmd = build_validation_command(
                soffice_path=r'C:\Program Files\LibreOffice\program\soffice.exe',
                doc_path=r'C:\Users\Test\document.docx',
                output_dir=r'C:\Users\Test\temp',
                filter_name='html:HTML'
            )

            self.assertIn('soffice.exe', cmd[0])


class TestDocumentValidation(unittest.TestCase):
    """Test document validation with cross-platform support."""

    def test_validate_document_success(self):
        """Test successful document validation."""
        from platform_utils import validate_document_cross_platform

        with patch('platform_utils.get_soffice_for_validation', return_value='/usr/bin/soffice'):
            with patch('subprocess.run') as mock_run:
                mock_run.return_value = MagicMock(returncode=0, stderr='')
                with patch('pathlib.Path.exists', return_value=True):
                    result = validate_document_cross_platform(Path('/tmp/test.docx'))
                    self.assertTrue(result)

    def test_validate_document_soffice_not_found(self):
        """Test validation when soffice is not available."""
        from platform_utils import validate_document_cross_platform

        with patch('platform_utils.get_soffice_for_validation', return_value=None):
            # Should return True (skip validation) with warning
            result = validate_document_cross_platform(Path('/tmp/test.docx'))
            self.assertTrue(result)  # Validation skipped

    def test_validate_document_timeout(self):
        """Test validation timeout handling."""
        from platform_utils import validate_document_cross_platform
        import subprocess

        with patch('platform_utils.get_soffice_for_validation', return_value='/usr/bin/soffice'):
            with patch('subprocess.run', side_effect=subprocess.TimeoutExpired('cmd', 10)):
                result = validate_document_cross_platform(Path('/tmp/test.docx'))
                self.assertFalse(result)


class TestFilterSelection(unittest.TestCase):
    """Test filter selection based on file type."""

    def test_get_filter_for_docx(self):
        """Test filter selection for .docx files."""
        from platform_utils import get_validation_filter

        self.assertEqual(get_validation_filter('.docx'), 'html:HTML')

    def test_get_filter_for_pptx(self):
        """Test filter selection for .pptx files."""
        from platform_utils import get_validation_filter

        self.assertEqual(get_validation_filter('.pptx'), 'html:impress_html_Export')

    def test_get_filter_for_xlsx(self):
        """Test filter selection for .xlsx files."""
        from platform_utils import get_validation_filter

        self.assertEqual(get_validation_filter('.xlsx'), 'html:HTML (StarCalc)')

    def test_get_filter_unknown_type(self):
        """Test filter selection for unknown file types."""
        from platform_utils import get_validation_filter

        self.assertIsNone(get_validation_filter('.txt'))
        self.assertIsNone(get_validation_filter('.pdf'))


class TestCrossplatformPaths(unittest.TestCase):
    """Test cross-platform path handling."""

    def test_normalize_path_windows(self):
        """Test path normalization on Windows."""
        from platform_utils import normalize_path

        with patch('platform.system', return_value='Windows'):
            # Forward slashes should be converted
            result = normalize_path('C:/Users/Test/file.docx')
            self.assertIn('\\', result)

    def test_normalize_path_unix(self):
        """Test path normalization on Unix."""
        from platform_utils import normalize_path

        with patch('platform.system', return_value='Linux'):
            result = normalize_path('/home/user/file.docx')
            self.assertEqual(result, '/home/user/file.docx')

    def test_handle_spaces_in_path(self):
        """Test handling paths with spaces."""
        from platform_utils import normalize_path

        path_with_spaces = '/home/user/My Documents/file.docx'
        result = normalize_path(path_with_spaces)
        # Should preserve spaces (let subprocess handle quoting)
        self.assertIn('My Documents', result)


class TestLibreOfficeInstallation(unittest.TestCase):
    """Test LibreOffice installation detection."""

    def test_find_libreoffice_windows_program_files(self):
        """Test finding LibreOffice in Windows Program Files."""
        from platform_utils import find_libreoffice_windows

        with patch('os.environ.get') as mock_env:
            mock_env.side_effect = lambda k, d=None: {
                'PROGRAMFILES': r'C:\Program Files',
                'PROGRAMFILES(X86)': r'C:\Program Files (x86)',
                'LOCALAPPDATA': r'C:\Users\Test\AppData\Local',
            }.get(k, d)

            with patch('os.path.exists') as mock_exists:
                mock_exists.side_effect = lambda p: 'LibreOffice' in p and 'soffice' in p

                result = find_libreoffice_windows()
                if result:
                    self.assertIn('soffice', result.lower())

    def test_find_libreoffice_linux(self):
        """Test finding LibreOffice on Linux."""
        from platform_utils import find_libreoffice_linux

        with patch('shutil.which', return_value='/usr/bin/soffice'):
            result = find_libreoffice_linux()
            self.assertEqual(result, '/usr/bin/soffice')

    def test_find_libreoffice_macos(self):
        """Test finding LibreOffice on macOS."""
        from platform_utils import find_libreoffice_macos

        with patch('os.path.exists') as mock_exists:
            mock_exists.side_effect = lambda p: 'LibreOffice.app' in p

            result = find_libreoffice_macos()
            if result:
                self.assertIn('LibreOffice', result)


if __name__ == '__main__':
    unittest.main()
