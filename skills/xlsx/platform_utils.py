#!/usr/bin/env python3
"""
Cross-platform utilities for xlsx skill.
Handles LibreOffice detection, subprocess timeout, and Windows compatibility.
"""

import os
import platform
import shutil
import subprocess
from pathlib import Path
from typing import Optional, Dict, Any, List


# Windows path length limit (without \\?\ prefix)
WINDOWS_MAX_PATH = 260


def get_libreoffice_macro_dir() -> str:
    """
    Get the LibreOffice macro directory for the current platform.

    Returns:
        Path to the LibreOffice user macro directory.

    Raises:
        EnvironmentError: If required environment variables are not set on Windows.
    """
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

    else:  # Linux and other Unix-like systems
        return os.path.expanduser('~/.config/libreoffice/4/user/basic/Standard')


def get_soffice_command() -> Optional[str]:
    """
    Find the soffice command/executable for the current platform.

    Checks PATH first, then platform-specific installation locations.

    Returns:
        Path to soffice executable, or None if not found.
    """
    # First, check if soffice is in PATH
    soffice_path = shutil.which('soffice')
    if soffice_path:
        return soffice_path

    system = platform.system()

    if system == 'Windows':
        return _find_soffice_windows()
    elif system == 'Darwin':
        return _find_soffice_macos()
    else:
        # On Linux, if not in PATH, it's not installed
        return None


def _find_soffice_windows() -> Optional[str]:
    """Find soffice on Windows by checking common installation paths."""
    # Common Windows installation paths
    program_files = os.environ.get('PROGRAMFILES', r'C:\Program Files')
    program_files_x86 = os.environ.get('PROGRAMFILES(X86)', r'C:\Program Files (x86)')
    local_appdata = os.environ.get('LOCALAPPDATA', '')

    search_paths = [
        os.path.join(program_files, 'LibreOffice', 'program', 'soffice.exe'),
        os.path.join(program_files_x86, 'LibreOffice', 'program', 'soffice.exe'),
    ]

    if local_appdata:
        search_paths.append(
            os.path.join(local_appdata, 'Programs', 'LibreOffice', 'program', 'soffice.exe')
        )

    for path in search_paths:
        if os.path.exists(path):
            return path

    return None


def _find_soffice_macos() -> Optional[str]:
    """Find soffice on macOS by checking common installation paths."""
    app_paths = [
        '/Applications/LibreOffice.app/Contents/MacOS/soffice',
        os.path.expanduser('~/Applications/LibreOffice.app/Contents/MacOS/soffice'),
    ]

    for path in app_paths:
        if os.path.exists(path):
            return path

    return None


def run_with_timeout(
    cmd: List[str],
    timeout: int,
    capture_output: bool = True,
    text: bool = True,
    **kwargs
) -> subprocess.CompletedProcess:
    """
    Run a subprocess with timeout, cross-platform.

    Uses Python's subprocess timeout parameter which works on all platforms.

    Args:
        cmd: Command and arguments as list.
        timeout: Timeout in seconds.
        capture_output: Whether to capture stdout/stderr.
        text: Whether to decode output as text.
        **kwargs: Additional arguments passed to subprocess.run.

    Returns:
        CompletedProcess instance.

    Raises:
        subprocess.TimeoutExpired: If command times out.
    """
    return subprocess.run(
        cmd,
        timeout=timeout,
        capture_output=capture_output,
        text=text,
        **kwargs
    )


def is_excel_com_available() -> bool:
    """
    Check if Excel COM automation is available (Windows only).

    Returns:
        True if Excel COM is available and Excel is installed.
    """
    if platform.system() != 'Windows':
        return False

    try:
        import win32com.client
        return _check_excel_installed()
    except ImportError:
        return False


def _check_excel_installed() -> bool:
    """Check if Microsoft Excel is installed and accessible via COM."""
    try:
        import win32com.client
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Quit()
        return True
    except Exception:
        return False


def validate_path_for_platform(path: str) -> Dict[str, Any]:
    """
    Validate a file path for the current platform.

    Checks for Windows path length limits and other platform-specific issues.

    Args:
        path: The file path to validate.

    Returns:
        Dict with 'valid' (bool) and optional 'error' (str) keys.
    """
    result = {'valid': True}

    if platform.system() == 'Windows':
        # Check for \\?\ prefix which allows long paths
        if path.startswith('\\\\?\\'):
            return result

        # Check path length
        if len(path) > WINDOWS_MAX_PATH:
            result['valid'] = False
            result['error'] = (
                f"Path length ({len(path)}) exceeds Windows limit ({WINDOWS_MAX_PATH}). "
                f"Consider using \\\\?\\ prefix for long paths."
            )

    return result


def get_temp_dir() -> Path:
    """
    Get a suitable temporary directory for the current platform.

    Returns:
        Path to temporary directory.
    """
    import tempfile
    return Path(tempfile.gettempdir())


def ensure_short_temp_path() -> Path:
    """
    Ensure a short path for temporary files (helpful on Windows).

    Returns:
        Path to a short temporary directory.
    """
    if platform.system() == 'Windows':
        # Use a shorter path on Windows to avoid path length issues
        short_temp = Path('C:\\Temp')
        if short_temp.exists() or _try_create_dir(short_temp):
            return short_temp

    return get_temp_dir()


def _try_create_dir(path: Path) -> bool:
    """Try to create a directory, return False on failure."""
    try:
        path.mkdir(parents=True, exist_ok=True)
        return True
    except (OSError, PermissionError):
        return False
