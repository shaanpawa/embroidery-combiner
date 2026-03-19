"""
Pipeline orchestrator: maps Excel combo files to DST files and exports combined DSTs.
"""

import os
from dataclasses import dataclass, field
from typing import Callable, List, Optional, Tuple

from app.core.combiner import (
    CombineError, combine_designs_two_column, save_combined,
    validate_combined_output,
)
from app.core.excel_parser import ComboFile


@dataclass
class ExportResult:
    combo: ComboFile
    success: bool
    output_path: str = ""
    error: str = ""
    validation: dict = field(default_factory=dict)


def resolve_dst_files(
    combo: ComboFile,
    dst_folder: str,
) -> Tuple[List[str], List[int]]:
    """
    Map combo slots to DST file paths.

    Returns:
        (found_paths, missing_programs) - found_paths is ordered by slot position,
        missing_programs lists program numbers without a matching DST file.
    """
    # Build a case-insensitive lookup of files in the folder
    try:
        folder_files = {f.lower(): f for f in os.listdir(dst_folder)}
    except OSError:
        folder_files = {}

    found = []
    missing = []
    for entry in combo.slots:
        # Try both lowercase and uppercase extensions
        filename_lower = f"{entry.program}.dst"
        actual_name = folder_files.get(filename_lower)
        if actual_name:
            found.append(os.path.join(dst_folder, actual_name))
        else:
            missing.append(entry.program)
            found.append(None)
    return found, missing


def check_combo_ready(combo: ComboFile, dst_folder: str) -> Tuple[bool, List[int]]:
    """Check if all DST files exist for a combo. Returns (ready, missing_programs)."""
    _, missing = resolve_dst_files(combo, dst_folder)
    return len(missing) == 0, missing


def export_combo(
    combo: ComboFile,
    dst_folder: str,
    output_folder: str,
    gap_mm: float = 3.0,
    column_gap_mm: float = 10.0,
    overwrite: bool = False,
) -> ExportResult:
    """Generate one combo DST file."""
    paths, missing = resolve_dst_files(combo, dst_folder)

    if missing:
        return ExportResult(
            combo=combo,
            success=False,
            error=f"Missing DST files for programs: {missing}",
        )

    left_paths = paths[:10]
    right_paths = paths[10:]

    # Filter out None (shouldn't happen if missing is empty, but be safe)
    left_paths = [p for p in left_paths if p is not None]
    right_paths = [p for p in right_paths if p is not None]

    output_path = os.path.join(output_folder, combo.filename)

    try:
        pattern = combine_designs_two_column(
            left_paths, right_paths,
            gap_mm=gap_mm,
            column_gap_mm=column_gap_mm,
        )
        save_combined(pattern, output_path, overwrite=overwrite)
        validation = validate_combined_output(output_path)
        return ExportResult(
            combo=combo, success=True,
            output_path=output_path, validation=validation,
        )
    except (CombineError, FileExistsError) as e:
        return ExportResult(combo=combo, success=False, error=str(e))


def export_all(
    combos: List[ComboFile],
    dst_folder: str,
    output_folder: str,
    gap_mm: float = 3.0,
    column_gap_mm: float = 10.0,
    overwrite: bool = False,
    progress_callback: Optional[Callable] = None,
) -> List[ExportResult]:
    """Export multiple combo files. Returns results per file."""
    results = []
    for i, combo in enumerate(combos):
        result = export_combo(combo, dst_folder, output_folder, gap_mm, column_gap_mm, overwrite)
        results.append(result)
        if progress_callback:
            progress_callback(i + 1, len(combos))
    return results
