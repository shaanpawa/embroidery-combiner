"""
Excel order parser for embroidery combo workflow.
Reads an order Excel, groups names by (machine_program, com_no),
expands quantities, and splits into combo files of max 20 slots.

Supports auto-detection of column positions from headers.
"""

from collections import defaultdict
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from openpyxl import load_workbook

# Default column indices (A=0, F=5, G=6, H=7, O=14, P=15)
DEFAULT_COLUMN_MAP = {
    "program": 0,
    "name_line1": 5,
    "name_line2": 6,
    "quantity": 7,
    "com_no": 14,
    "machine_program": 15,
}

# Header patterns for auto-detection (lowercase). Order matters — first match wins.
HEADER_PATTERNS = {
    "program": ["program", "prog", "prg", "file number", "dst"],
    "name_line1": ["row 1", "name 1", "first name", "embroidery name", "name"],
    "name_line2": ["row 2", "name 2", "title", "subtitle", "organisation", "organization"],
    "quantity": ["quantity", "qty", "amount", "count", "pcs"],
    "com_no": ["com no", "combo no", "combo number", "combo", "com"],
    "machine_program": ["machine program", "machine prog", "machine"],
}

# Fields where "M" alone is a valid header (special case — too short for fuzzy match)
M_HEADER_FIELD = "machine_program"


@dataclass
class NameEntry:
    program: int
    name_line1: str
    name_line2: str
    quantity: int
    com_no: str
    machine_program: str


@dataclass
class ComboGroup:
    machine_program: str
    com_no: str
    entries: List[NameEntry] = field(default_factory=list)

    @property
    def group_key(self) -> Tuple[str, str]:
        return (self.machine_program, self.com_no)

    @property
    def total_slots(self) -> int:
        return sum(e.quantity for e in self.entries)


@dataclass
class ComboFile:
    machine_program: str
    com_no: str
    part_number: int
    total_parts: int
    slots: List[NameEntry] = field(default_factory=list)

    @property
    def left_column(self) -> List[NameEntry]:
        return self.slots[:10]

    @property
    def right_column(self) -> List[NameEntry]:
        return self.slots[10:]

    @property
    def filename(self) -> str:
        return f"{self.machine_program}_Com{self.com_no}_{self.part_number}of{self.total_parts}.dst"


@dataclass
class ParseResult:
    entries: List[NameEntry] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)


@dataclass
class DetectResult:
    headers: List[str]
    preview_rows: List[List]
    detected_mapping: Dict[str, int]
    confidence: str  # "high" or "low"


def detect_columns(path: str, preview_count: int = 5) -> DetectResult:
    """Read Excel headers and auto-detect column mapping.

    Returns headers, sample rows, detected mapping, and confidence level.
    """
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active

    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    if not rows:
        return DetectResult([], [], dict(DEFAULT_COLUMN_MAP), "low")

    # Row 1 = headers
    raw_headers = list(rows[0]) if rows else []
    headers = [str(h or "").strip() for h in raw_headers]
    headers_lower = [h.lower() for h in headers]

    # Sample data rows (skip header)
    preview_rows = []
    for row in rows[1:preview_count + 1]:
        preview_rows.append([_cell_to_json(c) for c in row])

    # Auto-detect mapping
    mapping: Dict[str, int] = {}
    used_indices: set = set()

    # Special case: header exactly "M" maps to machine_program
    for i, h in enumerate(headers):
        if h.strip() == "M" and M_HEADER_FIELD not in mapping:
            mapping[M_HEADER_FIELD] = i
            used_indices.add(i)
            break

    # Match each field by header patterns
    for field_name, patterns in HEADER_PATTERNS.items():
        if field_name in mapping:
            continue  # already matched (e.g., machine_program via "M")
        for pattern in patterns:
            for i, h in enumerate(headers_lower):
                if i in used_indices:
                    continue
                if pattern == h or (len(pattern) > 2 and pattern in h):
                    # For "program", skip if this looks like a duplicate (column N often duplicates A)
                    if field_name == "program" and i > 0 and "program" in mapping:
                        continue
                    mapping[field_name] = i
                    used_indices.add(i)
                    break
            if field_name in mapping:
                break

    # Fill missing fields from defaults
    all_fields = list(DEFAULT_COLUMN_MAP.keys())
    matched = len(mapping)
    for f in all_fields:
        if f not in mapping:
            mapping[f] = DEFAULT_COLUMN_MAP[f]

    confidence = "high" if matched >= 5 else ("low" if matched < 3 else "medium")

    return DetectResult(
        headers=headers,
        preview_rows=preview_rows,
        detected_mapping=mapping,
        confidence=confidence,
    )


def _cell_to_json(val):
    """Convert a cell value to a JSON-safe type."""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        if isinstance(val, float) and val == int(val):
            return int(val)
        return val
    return str(val)


def parse_excel(path: str, column_map: Optional[Dict[str, int]] = None) -> ParseResult:
    """Read Excel and extract entries using column mapping.

    If column_map is not provided, uses DEFAULT_COLUMN_MAP.
    """
    cmap = column_map or DEFAULT_COLUMN_MAP
    idx_program = cmap["program"]
    idx_name1 = cmap["name_line1"]
    idx_name2 = cmap["name_line2"]
    idx_qty = cmap["quantity"]
    idx_com = cmap["com_no"]
    idx_m = cmap["machine_program"]
    max_idx = max(idx_program, idx_name1, idx_name2, idx_qty, idx_com, idx_m)

    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    result = ParseResult()

    for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if len(row) <= max_idx:
            continue

        program = row[idx_program]
        name1 = row[idx_name1]
        name2 = row[idx_name2]
        qty = row[idx_qty]
        com_no = row[idx_com]
        m_val = row[idx_m]

        if program is None or m_val is None:
            if program is not None or name1 is not None:
                result.warnings.append(f"Row {row_num}: missing program or M value, skipped")
            continue

        try:
            program = int(program)
        except (ValueError, TypeError):
            result.warnings.append(f"Row {row_num}: program '{program}' is not a number, skipped")
            continue

        if qty is None or qty < 1:
            if qty is not None and qty < 1:
                result.warnings.append(f"Row {row_num}: qty={qty} treated as 1")
            qty = 1
        else:
            qty = int(qty)

        result.entries.append(NameEntry(
            program=program,
            name_line1=str(name1 or ""),
            name_line2=str(name2 or ""),
            quantity=qty,
            com_no=str(int(com_no)) if com_no is not None and isinstance(com_no, (int, float)) else str(com_no or ""),
            machine_program=str(m_val),
        ))

    wb.close()
    return result


def group_entries(entries: List[NameEntry]) -> List[ComboGroup]:
    """Group by (machine_program, com_no), sort entries within each group by program number."""
    groups_dict = defaultdict(list)
    for entry in entries:
        key = (entry.machine_program, entry.com_no)
        groups_dict[key].append(entry)

    groups = []
    for (m, com), ents in sorted(groups_dict.items()):
        ents.sort(key=lambda e: e.program)
        groups.append(ComboGroup(machine_program=m, com_no=com, entries=ents))
    return groups


def expand_and_split(group: ComboGroup, max_slots: int = 20) -> List[ComboFile]:
    """
    Expand quantities and split into combo files of max_slots each.

    Sorts by (quantity DESC, program ASC) to cluster same-qty entries.
    Each entry is repeated by its quantity in the expanded slot list.
    Splits at max_slots boundaries.
    """
    sorted_entries = sorted(group.entries, key=lambda e: (-e.quantity, e.program))

    expanded = []
    for entry in sorted_entries:
        for _ in range(entry.quantity):
            expanded.append(entry)

    if not expanded:
        return []

    chunks = []
    for i in range(0, len(expanded), max_slots):
        chunks.append(expanded[i:i + max_slots])

    total_parts = len(chunks)
    combo_files = []
    for i, chunk in enumerate(chunks, 1):
        combo_files.append(ComboFile(
            machine_program=group.machine_program,
            com_no=group.com_no,
            part_number=i,
            total_parts=total_parts,
            slots=chunk,
        ))
    return combo_files


def generate_all_combos(entries: List[NameEntry], max_slots: int = 20) -> List[ComboFile]:
    """Full pipeline: group -> expand -> split. Returns all combo files."""
    groups = group_entries(entries)
    all_combos = []
    for group in groups:
        all_combos.extend(expand_and_split(group, max_slots))
    return all_combos
