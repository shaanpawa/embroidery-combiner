"""
File validation for embroidery design files.
Checks for corruption, empty files, format issues, and design sanity.
"""

import os
import struct
from typing import Callable, List, Optional

import pyembroidery

from app.config import MAX_FILE_SIZE_MB


class ValidationResult:
    """Result of validating a single file."""

    def __init__(self, path: str):
        self.path = path
        self.filename = os.path.basename(path)
        self.valid = False
        self.errors: List[str] = []
        self.warnings: List[str] = []
        self.stitch_count = 0
        self.color_count = 0
        self.width_mm = 0.0
        self.height_mm = 0.0
        self.file_size = 0

    def add_error(self, msg: str) -> None:
        self.errors.append(msg)
        self.valid = False

    def add_warning(self, msg: str) -> None:
        self.warnings.append(msg)

    @property
    def status(self) -> str:
        if self.errors:
            return "Error"
        if self.warnings:
            return "Warning"
        return "OK"

    @property
    def summary(self) -> str:
        if self.errors:
            return self.errors[0]
        if self.warnings:
            return self.warnings[0]
        if self.stitch_count > 0:
            return f"{self.stitch_count} stitches, {self.width_mm:.1f}x{self.height_mm:.1f}mm"
        return "Valid (NGS)"


def is_valid_for_combining(result: ValidationResult) -> bool:
    """Check if a validated file can be used in combining."""
    return result.valid and not result.errors


def validate_file(path: str) -> ValidationResult:
    """Validate a single embroidery file."""
    result = ValidationResult(path)

    if not os.path.exists(path):
        result.add_error("File not found")
        return result

    if not os.access(path, os.R_OK):
        result.add_error("File not readable (permission denied)")
        return result

    result.file_size = os.path.getsize(path)
    if result.file_size == 0:
        result.add_error("File is empty (0 bytes)")
        return result

    if result.file_size > MAX_FILE_SIZE_MB * 1024 * 1024:
        result.add_warning(f"Very large file ({result.file_size / (1024*1024):.1f} MB)")

    ext = os.path.splitext(path)[1].lower()
    if ext == '.ngs':
        return _validate_ngs(path, result)
    elif ext == '.dst':
        return _validate_dst(path, result)
    else:
        result.add_error(f"Unsupported format: {ext}")
        return result


def _validate_ngs(path: str, result: ValidationResult) -> ValidationResult:
    """Validate an NGS file by checking OLE2 structure."""
    # Check OLE2 magic bytes
    try:
        with open(path, 'rb') as f:
            magic = f.read(8)
        if magic != b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
            result.add_error("Not a valid NGS file (corrupted or wrong format)")
            return result
    except IOError as e:
        result.add_error(f"Cannot read file: {e}")
        return result

    try:
        import olefile
    except ImportError:
        # Can't deep-validate without olefile, but magic bytes passed
        result.valid = True
        result.add_warning("Install olefile for deeper NGS validation")
        return result

    try:
        ole = olefile.OleFileIO(path)
    except Exception as e:
        result.add_error(f"Corrupted file container: {e}")
        return result

    try:
        streams = ['/'.join(s) for s in ole.listdir()]

        if 'design.vvt' not in streams:
            result.add_error("Missing stitch data — file is incomplete")
            return result

        vvt_data = ole.openstream('design.vvt').read()
        if len(vvt_data) < 12:
            result.add_error("Stitch data too small — file is corrupted")
            return result

        word0, word1, word2 = struct.unpack_from('<III', vvt_data, 0)
        expected_data_size = len(vvt_data) - 12
        if word1 != expected_data_size:
            result.add_warning("Header size mismatch — file may be corrupted")

        if 'index.tmp' not in streams:
            result.add_warning("Missing index data — file may be incomplete")

        if 'stats.dat' not in streams:
            result.add_warning("Missing design statistics")

        result.stitch_count = -1  # Can't extract from NGS
        result.valid = True

    finally:
        ole.close()

    return result


def _validate_dst(path: str, result: ValidationResult) -> ValidationResult:
    """Validate a DST file using pyembroidery."""
    try:
        with open(path, 'rb') as f:
            header = f.read(3)
        if header != b'LA:':
            result.add_warning("Non-standard DST header")
    except IOError as e:
        result.add_error(f"Cannot read file: {e}")
        return result

    try:
        pattern = pyembroidery.read(path)
    except Exception as e:
        result.add_error(f"Failed to parse: {e}")
        return result

    if pattern is None:
        result.add_error("Could not read file — corrupted or unsupported")
        return result

    result.stitch_count = len(pattern.stitches)
    result.color_count = len(pattern.threadlist)

    if result.stitch_count == 0:
        result.add_error("No stitches — empty or corrupted design")
        return result

    try:
        ext = pattern.extents()
        if ext is None or len(ext) < 4:
            result.add_error("Cannot determine design dimensions")
            return result

        result.width_mm = (ext[2] - ext[0]) / 10.0
        result.height_mm = (ext[3] - ext[1]) / 10.0

        if result.width_mm <= 0 or result.height_mm <= 0:
            result.add_error("Zero or negative dimensions — corrupted stitch data")
            return result

        if result.width_mm > 1000 or result.height_mm > 2000:
            result.add_warning(
                f"Very large design ({result.width_mm:.0f}x{result.height_mm:.0f}mm)"
            )

        if result.stitch_count < 3:
            result.add_warning(f"Only {result.stitch_count} stitches — possibly incomplete")

    except Exception as e:
        result.add_warning(f"Could not determine dimensions: {e}")

    result.valid = True
    return result


def validate_batch(
    file_paths: List[str],
    progress_callback: Optional[Callable] = None,
) -> List[ValidationResult]:
    """
    Validate a list of files.

    Args:
        file_paths: Paths to validate.
        progress_callback: Called as progress_callback(current, total, result).

    Returns:
        List of ValidationResult, one per file.
    """
    results = []
    for i, path in enumerate(file_paths):
        result = validate_file(path)
        results.append(result)
        if progress_callback:
            progress_callback(i + 1, len(file_paths), result)
    return results
