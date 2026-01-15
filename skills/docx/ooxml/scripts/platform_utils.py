#!/usr/bin/env python3
"""
Cross-platform utilities for docx/ooxml scripts.
Handles LibreOffice detection and validation across Windows, macOS, and Linux.
"""

import os
import platform
import shutil
import subprocess
from pathlib import Path
from typing import Optional, List


def get_soffice_for_validation() -> Optional[str]:
    """
    Find soffice executable for document validation.

    Returns:
        Path to soffice executable, or None if not found.
    """
    # First check PATH
    soffice_path = shutil.which('soffice')
    if soffice_path:
        return soffice_path

    system = platform.system()

    if system == 'Windows':
        return find_libreoffice_windows()
    elif system == 'Darwin':
        return find_libreoffice_macos()
    else:
        return find_libreoffice_linux()


def find_libreoffice_windows() -> Optional[str]:
    """Find LibreOffice installation on Windows."""
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


def find_libreoffice_linux() -> Optional[str]:
    """Find LibreOffice installation on Linux."""
    # On Linux, rely on PATH
    return shutil.which('soffice')


def find_libreoffice_macos() -> Optional[str]:
    """Find LibreOffice installation on macOS."""
    app_paths = [
        '/Applications/LibreOffice.app/Contents/MacOS/soffice',
        os.path.expanduser('~/Applications/LibreOffice.app/Contents/MacOS/soffice'),
    ]

    for path in app_paths:
        if os.path.exists(path):
            return path

    # Also check PATH
    return shutil.which('soffice')


def get_validation_filter(extension: str) -> Optional[str]:
    """
    Get the LibreOffice export filter for a file type.

    Args:
        extension: File extension including dot (e.g., '.docx').

    Returns:
        Filter name string, or None for unsupported types.
    """
    filters = {
        '.docx': 'html:HTML',
        '.pptx': 'html:impress_html_Export',
        '.xlsx': 'html:HTML (StarCalc)',
    }
    return filters.get(extension.lower())


def build_validation_command(
    soffice_path: str,
    doc_path: str,
    output_dir: str,
    filter_name: str
) -> List[str]:
    """
    Build the soffice validation command.

    Args:
        soffice_path: Path to soffice executable.
        doc_path: Path to document to validate.
        output_dir: Directory for output files.
        filter_name: LibreOffice export filter name.

    Returns:
        Command as list of strings.
    """
    return [
        soffice_path,
        '--headless',
        '--convert-to',
        filter_name,
        '--outdir',
        output_dir,
        doc_path,
    ]


def validate_document_cross_platform(
    doc_path: Path,
    timeout: int = 10
) -> bool:
    """
    Validate a document using LibreOffice, cross-platform.

    Args:
        doc_path: Path to the document to validate.
        timeout: Timeout in seconds.

    Returns:
        True if validation passed or was skipped, False if validation failed.
    """
    import tempfile

    soffice = get_soffice_for_validation()
    if not soffice:
        print("Warning: soffice not found. Skipping validation.", flush=True)
        return True  # Skip validation

    filter_name = get_validation_filter(doc_path.suffix)
    if not filter_name:
        print(f"Warning: No filter for {doc_path.suffix}. Skipping validation.", flush=True)
        return True

    with tempfile.TemporaryDirectory() as temp_dir:
        cmd = build_validation_command(
            soffice_path=soffice,
            doc_path=str(doc_path),
            output_dir=temp_dir,
            filter_name=filter_name
        )

        try:
            result = subprocess.run(
                cmd,
                capture_output=True,
                timeout=timeout,
                text=True,
            )

            # Check if output file was created
            expected_output = Path(temp_dir) / f"{doc_path.stem}.html"
            if not expected_output.exists():
                error_msg = result.stderr.strip() or "Document validation failed"
                print(f"Validation error: {error_msg}", flush=True)
                return False

            return True

        except subprocess.TimeoutExpired:
            print("Validation error: Timeout during conversion", flush=True)
            return False
        except FileNotFoundError:
            print("Warning: soffice not found. Skipping validation.", flush=True)
            return True
        except Exception as e:
            print(f"Validation error: {e}", flush=True)
            return False


def normalize_path(path: str) -> str:
    """
    Normalize a path for the current platform.

    Args:
        path: The path to normalize.

    Returns:
        Normalized path string.
    """
    if platform.system() == 'Windows':
        # Convert forward slashes to backslashes on Windows
        return path.replace('/', '\\')
    return path
