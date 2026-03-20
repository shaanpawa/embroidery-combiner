"""
Core combining logic for embroidery design files.
Reads multiple DST files and combines them vertically with configurable gap.
Supports single-column and two-column layouts.
"""

import errno
import os
from typing import Callable, List, Optional

import pyembroidery


class CombineError(Exception):
    """Raised when combining fails."""
    pass


def _strip_trailing_footer(stitches):
    """Strip trailing TRIM + JUMPs + END from a design's stitch list.

    Individual DST files end with: [stitching] → TRIM → JUMPs → END
    When combining, this footer must be removed so designs join cleanly
    with just a COLOR_CHANGE between them. Truncate at the last STITCH.
    """
    last_stitch = -1
    for i in range(len(stitches) - 1, -1, -1):
        if stitches[i][2] == pyembroidery.STITCH:
            last_stitch = i
            break
    if last_stitch < 0:
        return [s for s in stitches if s[2] != pyembroidery.END]
    return stitches[:last_stitch + 1]


def _strip_extra_color_changes(pattern: pyembroidery.EmbPattern, num_designs: int) -> None:
    """Remove duplicate/spurious COLOR_CHANGE commands if any exist.

    Each individual name DST has exactly 1 internal COLOR_CHANGE (red→green).
    Between designs, add_pattern() inserts 1 COLOR_CHANGE which is the correct
    needle switch (green→red for the next name). So the expected count is:
      N internal + (N-1) between = 2N-1 total COLOR_CHANGEs.

    This function is a safety net: if add_pattern() inserts extra CCs beyond
    what's expected (e.g., consecutive CCs with no stitches between them),
    remove the duplicates.
    """
    if num_designs <= 1:
        return

    # Find all COLOR_CHANGE indices
    cc_indices = [i for i, s in enumerate(pattern.stitches) if s[2] == pyembroidery.COLOR_CHANGE]

    # Remove consecutive COLOR_CHANGEs (keep only the first of any run)
    to_remove = set()
    for i in range(1, len(cc_indices)):
        # Check if there are any STITCH commands between consecutive CCs
        has_stitch = False
        for j in range(cc_indices[i-1] + 1, cc_indices[i]):
            if pattern.stitches[j][2] == pyembroidery.STITCH:
                has_stitch = True
                break
        if not has_stitch:
            to_remove.add(cc_indices[i])  # Duplicate CC, remove it

    if to_remove:
        pattern.stitches = [s for i, s in enumerate(pattern.stitches) if i not in to_remove]


def _read_design(path: str) -> pyembroidery.EmbPattern:
    """Read a DST file and validate it has stitches and dimensions."""
    try:
        design = pyembroidery.read(path)
    except Exception as e:
        raise CombineError(f"Failed to read {os.path.basename(path)}: {e}")

    if design is None or len(design.stitches) == 0:
        raise CombineError(f"Cannot read or empty: {os.path.basename(path)}")

    ext = design.extents()
    if ext is None or (ext[0] == ext[2] and ext[1] == ext[3]):
        raise CombineError(f"Zero-dimension design: {os.path.basename(path)}")

    return design


def _stack_vertical(
    dst_files: List[str],
    gap: float,
    progress_callback: Optional[Callable] = None,
    progress_offset: int = 0,
    progress_total: int = 0,
) -> pyembroidery.EmbPattern:
    """
    Stack multiple DST files vertically with a gap (in 1/10mm units).
    Returns a combined pattern with no END command (caller adds it).
    """
    if not dst_files:
        raise CombineError("No files to stack")

    combined = _read_design(dst_files[0])
    combined.stitches = _strip_trailing_footer(combined.stitches)

    for i, path in enumerate(dst_files[1:], 1):
        design = _read_design(path)

        c_ext = combined.extents()
        d_ext = design.extents()

        if c_ext is None:
            raise CombineError("Cannot determine dimensions of combined design")
        if d_ext is None:
            raise CombineError(f"Cannot determine dimensions of {os.path.basename(path)}")

        y_offset = c_ext[3] - d_ext[1] + gap
        design.translate(0, y_offset)
        design.stitches = _strip_trailing_footer(design.stitches)

        # Between designs: COLOR_CHANGE (needle switch green→red) + JUMPs to next position.
        # NO TRIM — matches reference combo files. Machine should NOT cut thread between names.
        combined.add_command(pyembroidery.COLOR_CHANGE)
        combined.add_pattern(design)

        if progress_callback:
            progress_callback(progress_offset + i, progress_total)

    return combined


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

    gap_mm = max(0.0, min(50.0, gap_mm))
    gap = gap_mm * 10  # pyembroidery uses 1/10mm units

    if len(dst_files) == 1:
        return _read_design(dst_files[0])

    total = len(dst_files) - 1
    combined = _stack_vertical(
        dst_files, gap,
        progress_callback=progress_callback,
        progress_offset=0,
        progress_total=total,
    )
    _strip_extra_color_changes(combined, len(dst_files))
    combined.add_command(pyembroidery.END)
    return combined


def combine_designs_two_column(
    left_files: List[str],
    right_files: List[str],
    gap_mm: float = 3.0,
    column_gap_mm: float = 10.0,
    progress_callback: Optional[Callable] = None,
) -> pyembroidery.EmbPattern:
    """
    Combine designs in a two-column layout.

    Left column (slots 1-10) stacked vertically, right column (slots 11-20)
    stacked vertically and offset horizontally. Both columns start at the same y.

    Args:
        left_files: DST file paths for the left column (max 10).
        right_files: DST file paths for the right column (max 10).
        gap_mm: Vertical gap between designs in millimeters.
        column_gap_mm: Horizontal gap between columns in millimeters.
        progress_callback: Called as progress_callback(current, total).

    Returns:
        Combined EmbPattern object.

    Raises:
        CombineError: If combining fails.
    """
    if not left_files and not right_files:
        raise CombineError("No files to combine")

    gap_mm = max(0.0, min(50.0, gap_mm))
    column_gap_mm = max(0.0, min(100.0, column_gap_mm))
    gap = gap_mm * 10
    column_gap = column_gap_mm * 10

    total_files = len(left_files) + len(right_files)

    # Left column only
    if not right_files:
        if len(left_files) == 1:
            return _read_design(left_files[0])
        combined = _stack_vertical(
            left_files, gap, progress_callback,
            progress_offset=0, progress_total=total_files - 1,
        )
        _strip_extra_color_changes(combined, len(left_files) if not right_files else len(right_files))
        combined.add_command(pyembroidery.END)
        return combined

    # Right column only
    if not left_files:
        if len(right_files) == 1:
            return _read_design(right_files[0])
        combined = _stack_vertical(
            right_files, gap, progress_callback,
            progress_offset=0, progress_total=total_files - 1,
        )
        _strip_extra_color_changes(combined, len(left_files) if not right_files else len(right_files))
        combined.add_command(pyembroidery.END)
        return combined

    # Both columns
    left_progress = len(left_files) - 1
    left_pattern = _stack_vertical(
        left_files, gap, progress_callback,
        progress_offset=0, progress_total=total_files - 1,
    )

    right_pattern = _stack_vertical(
        right_files, gap, progress_callback,
        progress_offset=left_progress, progress_total=total_files - 1,
    )

    # Position right column next to left column
    l_ext = left_pattern.extents()
    r_ext = right_pattern.extents()

    if l_ext is None or r_ext is None:
        raise CombineError("Cannot determine column dimensions")

    # Horizontal offset: right edge of left column + gap - left edge of right column
    x_offset = l_ext[2] - r_ext[0] + column_gap
    # Vertical alignment: both columns start at the same y
    y_offset = l_ext[1] - r_ext[1]

    right_pattern.translate(x_offset, y_offset)
    right_pattern.stitches = _strip_trailing_footer(right_pattern.stitches)

    # Merge right column into left — COLOR_CHANGE (needle switch) + JUMPs, no TRIM
    left_pattern.add_command(pyembroidery.COLOR_CHANGE)
    left_pattern.add_pattern(right_pattern)
    _strip_extra_color_changes(left_pattern, len(left_files) + len(right_files))
    left_pattern.add_command(pyembroidery.END)

    return left_pattern


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


def validate_pattern_in_memory(pattern: pyembroidery.EmbPattern) -> dict:
    """Compute validation stats from an in-memory pattern (no disk re-read)."""
    if pattern is None or len(pattern.stitches) == 0:
        return {"valid": False, "error": "Empty pattern"}

    ext = pattern.extents()
    if ext is None:
        return {"valid": False, "error": "Cannot determine dimensions"}

    cc_count = sum(1 for s in pattern.stitches if s[2] == pyembroidery.COLOR_CHANGE)

    return {
        "valid": True,
        "stitch_count": len(pattern.stitches),
        "color_count": cc_count,
        "width_mm": (ext[2] - ext[0]) / 10.0,
        "height_mm": (ext[3] - ext[1]) / 10.0,
    }


