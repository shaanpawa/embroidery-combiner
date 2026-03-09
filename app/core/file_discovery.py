"""
Folder scanning, file sorting, and sequence analysis for embroidery files.
"""

import os
import re
from dataclasses import dataclass, field
from typing import List, Optional

from app.config import SUPPORTED_EXTENSIONS


@dataclass
class DiscoveredFile:
    path: str
    filename: str
    extension: str
    number: Optional[int]
    size_bytes: int
    included: bool = True

    @property
    def size_display(self) -> str:
        if self.size_bytes < 1024:
            return f"{self.size_bytes} B"
        elif self.size_bytes < 1024 * 1024:
            return f"{self.size_bytes / 1024:.1f} KB"
        else:
            return f"{self.size_bytes / (1024 * 1024):.1f} MB"


@dataclass
class DiscoveryResult:
    files: List[DiscoveredFile] = field(default_factory=list)
    ngs_count: int = 0
    dst_count: int = 0
    skipped_files: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)

    @property
    def has_mixed_formats(self) -> bool:
        return self.ngs_count > 0 and self.dst_count > 0

    @property
    def needs_conversion(self) -> bool:
        return self.ngs_count > 0

    @property
    def total_files(self) -> int:
        return len(self.files)


def is_range_filename(filename: str) -> bool:
    """Check if filename looks like a combined output (e.g., '216-225.dst')."""
    name = os.path.splitext(filename)[0]
    return bool(re.match(r'^\d+-\d+$', name))


def extract_number(filename: str) -> Optional[int]:
    """Extract the numeric part from a filename like '216.dst' or 'design_216.dst'."""
    name = os.path.splitext(filename)[0]
    match = re.search(r'(\d+)', name)
    if match:
        return int(match.group(1))
    return None


def _sort_key(f: DiscoveredFile):
    """Sort numbered files first (by number), then unnumbered alphabetically."""
    if f.number is not None:
        return (0, f.number, f.filename)
    return (1, 0, f.filename)


def discover_folder(folder_path: str) -> DiscoveryResult:
    """
    Scan a folder for embroidery files.

    Returns DiscoveryResult with sorted files, counts, and warnings.
    """
    result = DiscoveryResult()

    if not os.path.isdir(folder_path):
        result.warnings.append(f"Not a valid folder: {folder_path}")
        return result

    try:
        entries = os.listdir(folder_path)
    except OSError as e:
        result.warnings.append(f"Cannot read folder: {e}")
        return result

    for entry in entries:
        full_path = os.path.join(folder_path, entry)
        if not os.path.isfile(full_path):
            continue

        ext = os.path.splitext(entry)[1].lower()
        if ext not in SUPPORTED_EXTENSIONS:
            result.skipped_files.append(entry)
            continue

        # Skip files that look like previous combined output (e.g., "216-225.dst")
        if is_range_filename(entry):
            result.skipped_files.append(entry)
            result.warnings.append(f"Skipped '{entry}' (looks like a combined output)")
            continue

        try:
            size = os.path.getsize(full_path)
        except OSError:
            size = 0

        number = extract_number(entry)

        discovered = DiscoveredFile(
            path=full_path,
            filename=entry,
            extension=ext,
            number=number,
            size_bytes=size,
        )

        if ext == '.ngs':
            result.ngs_count += 1
        elif ext == '.dst':
            result.dst_count += 1

        result.files.append(discovered)

    result.files.sort(key=_sort_key)

    # Detect sequence issues
    result.warnings.extend(check_sequence_gaps(result.files))
    result.warnings.extend(check_duplicates(result.files))

    if result.total_files == 1:
        result.warnings.append("Only 1 file — nothing to combine")

    if result.has_mixed_formats:
        result.warnings.append(
            f"Mixed formats: {result.ngs_count} NGS + {result.dst_count} DST files"
        )

    if result.needs_conversion:
        result.warnings.append(
            f"{result.ngs_count} NGS file(s) will be converted to DST"
        )

    return result


def check_sequence_gaps(files: List[DiscoveredFile]) -> List[str]:
    """Check for gaps in numbered file sequence."""
    numbers = sorted(f.number for f in files if f.number is not None)
    if len(numbers) < 2:
        return []

    expected = set(range(numbers[0], numbers[-1] + 1))
    missing = sorted(expected - set(numbers))
    if missing:
        return [f"Missing in sequence: {', '.join(str(n) for n in missing)}"]
    return []


def check_duplicates(files: List[DiscoveredFile]) -> List[str]:
    """Check for duplicate numbers."""
    seen = {}
    warnings = []
    for f in files:
        if f.number is None:
            continue
        if f.number in seen:
            warnings.append(
                f"Duplicate number {f.number}: {seen[f.number]} and {f.filename}"
            )
        seen[f.number] = f.filename
    return warnings


def generate_output_name(files: List[DiscoveredFile], extension: str = '.dst') -> str:
    """Generate output filename from number range, e.g. '216-225.dst'."""
    numbers = sorted(f.number for f in files if f.included and f.number is not None)
    if not numbers:
        return f"combined{extension}"
    if len(numbers) == 1:
        return f"{numbers[0]}{extension}"
    return f"{numbers[0]}-{numbers[-1]}{extension}"
