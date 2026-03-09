"""
Core combining logic for embroidery design files.
Reads multiple DST files and combines them vertically with configurable gap.
"""

import errno
import os
from typing import Callable, List, Optional

import pyembroidery


class CombineError(Exception):
    """Raised when combining fails."""
    pass


def combine_designs(
    dst_files: List[str],
    gap_mm: float = 3.0,
    progress_callback: Optional[Callable] = None,
) -> pyembroidery.EmbPattern:
    """
    Combine multiple DST files into one, stacking vertically with a gap.

    Args:
        dst_files: List of file paths to DST files, in order.
        gap_mm: Gap between designs in millimeters.
        progress_callback: Called as progress_callback(current, total).

    Returns:
        Combined EmbPattern object.

    Raises:
        CombineError: If combining fails for any reason.
    """
    if not dst_files:
        raise CombineError("No files to combine")

    gap = gap_mm * 10  # pyembroidery uses 1/10mm units

    try:
        combined = pyembroidery.read(dst_files[0])
    except Exception as e:
        raise CombineError(f"Failed to read {os.path.basename(dst_files[0])}: {e}")

    if combined is None or len(combined.stitches) == 0:
        raise CombineError(f"Cannot read or empty: {os.path.basename(dst_files[0])}")

    c_ext = combined.extents()
    if c_ext is None or (c_ext[0] == c_ext[2] and c_ext[1] == c_ext[3]):
        raise CombineError(f"Zero-dimension design: {os.path.basename(dst_files[0])}")

    if len(dst_files) == 1:
        return combined

    # Strip END from first design so subsequent designs can be appended.
    # pyembroidery's DST writer stops at the first END command, which would
    # cause all designs after the first to be silently dropped from the output.
    combined.stitches = [s for s in combined.stitches if s[2] != pyembroidery.END]

    for i, path in enumerate(dst_files[1:], 1):
        try:
            design = pyembroidery.read(path)
        except Exception as e:
            raise CombineError(f"Failed to read {os.path.basename(path)}: {e}")

        if design is None or len(design.stitches) == 0:
            raise CombineError(f"Cannot read or empty: {os.path.basename(path)}")

        c_ext = combined.extents()
        d_ext = design.extents()

        if c_ext is None:
            raise CombineError("Cannot determine dimensions of combined design")
        if d_ext is None:
            raise CombineError(f"Cannot determine dimensions of {os.path.basename(path)}")

        # Stack: new design below combined + gap
        y_offset = c_ext[3] - d_ext[1] + gap
        design.translate(0, y_offset)

        # Strip END from this design before appending
        design.stitches = [s for s in design.stitches if s[2] != pyembroidery.END]

        # TRIM before next design so machine cuts thread
        combined.add_command(pyembroidery.TRIM)
        combined.add_pattern(design)

        if progress_callback:
            progress_callback(i, len(dst_files) - 1)

    # Add final END command
    combined.add_command(pyembroidery.END)

    return combined


def save_combined(
    pattern: pyembroidery.EmbPattern,
    output_path: str,
    overwrite: bool = False,
) -> str:
    """
    Save combined pattern to file.

    Args:
        pattern: The combined pattern.
        output_path: Where to save.
        overwrite: If False, raises FileExistsError when output exists.

    Returns:
        The output path.

    Raises:
        FileExistsError: If output exists and overwrite is False.
        CombineError: On disk full, permission denied, or other write errors.
    """
    if os.path.exists(output_path) and not overwrite:
        raise FileExistsError(f"Output file already exists: {output_path}")

    try:
        pyembroidery.write(pattern, output_path)
    except OSError as e:
        if e.errno == errno.ENOSPC:
            raise CombineError("Disk full — cannot save output file")
        if e.errno == errno.EACCES:
            raise CombineError(f"Permission denied: {output_path}")
        raise CombineError(f"Failed to save: {e}")

    # Verify the output was written
    if not os.path.exists(output_path):
        raise CombineError("Output file was not created")

    if os.path.getsize(output_path) == 0:
        raise CombineError("Output file is empty — save failed")

    return output_path


def validate_combined_output(output_path: str) -> dict:
    """
    Read back the combined output and return summary info.
    Used for post-combine verification.
    """
    try:
        pattern = pyembroidery.read(output_path)
    except Exception:
        return {"valid": False, "error": "Cannot read output file"}

    if pattern is None:
        return {"valid": False, "error": "Output file is unreadable"}

    stitch_count = len(pattern.stitches)
    if stitch_count == 0:
        return {"valid": False, "error": "Output has no stitches"}

    ext = pattern.extents()
    if ext is None:
        return {"valid": False, "error": "Cannot determine output dimensions"}

    return {
        "valid": True,
        "stitch_count": stitch_count,
        "color_count": len(pattern.threadlist),
        "width_mm": (ext[2] - ext[0]) / 10.0,
        "height_mm": (ext[3] - ext[1]) / 10.0,
    }
