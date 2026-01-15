#!/usr/bin/env python3
"""
Cross-platform utilities for xlsx skill.
Handles LibreOffice and Excel detection for formula recalculation.
"""

import os
import platform
import shutil
import subprocess
from typing import Optional, Dict, Any


def get_libreoffice_macro_dir() -> str:
    """Get the LibreOffice macro directory for the current platform."""
    system = platform.system()

    if system == 'Windows':
        appdata = os.environ.get('APPDATA')
        if not appdata:
            raise EnvironmentError("APPDATA environment variable not set")
        return os.path.join(appdata, 'LibreOffice', '4', 'user', 'basic', 'Standard')
    elif system == 'Darwin':
        return os.path.expanduser(
            '~/Library/Application Support/LibreOffice/4/user/basic/Standard'
        )
    else:
        return os.path.expanduser('~/.config/libreoffice/4/user/basic/Standard')


def get_soffice_command() -> Optional[str]:
    """Find the soffice executable. Returns None if not found."""
    soffice_path = shutil.which('soffice')
    if soffice_path:
        return soffice_path

    system = platform.system()
    if system == 'Windows':
        for base in [os.environ.get('PROGRAMFILES', r'C:\Program Files'),
                     os.environ.get('PROGRAMFILES(X86)', r'C:\Program Files (x86)')]:
            path = os.path.join(base, 'LibreOffice', 'program', 'soffice.exe')
            if os.path.exists(path):
                return path
    elif system == 'Darwin':
        for path in ['/Applications/LibreOffice.app/Contents/MacOS/soffice',
                     os.path.expanduser('~/Applications/LibreOffice.app/Contents/MacOS/soffice')]:
            if os.path.exists(path):
                return path
    return None


def is_excel_com_available() -> bool:
    """Check if Excel COM automation is available (Windows only)."""
    if platform.system() != 'Windows':
        return False
    try:
        import win32com.client
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Quit()
        return True
    except Exception:
        return False


def recalc_with_excel_com(filename: str) -> Dict[str, Any]:
    """Recalculate Excel file using Excel COM (Windows only)."""
    if platform.system() != 'Windows':
        return {'success': False, 'error': 'Excel COM only available on Windows'}

    try:
        import win32com.client
        import pythoncom
    except ImportError:
        return {'success': False, 'error': 'pywin32 not installed'}

    excel = None
    try:
        pythoncom.CoInitialize()
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(filename)
        excel.CalculateFull()
        wb.Save()
        wb.Close(SaveChanges=True)
        return {'success': True}
    except Exception as e:
        return {'success': False, 'error': str(e)}
    finally:
        if excel:
            try:
                excel.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
