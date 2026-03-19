"""
Excel order parser for embroidery combo workflow.
Reads an order Excel, groups names by (machine_program, com_no),
expands quantities, and splits into combo files of max 20 slots.
"""

from collections import defaultdict
from dataclasses import dataclass, field
from typing import List, Optional, Tuple

from openpyxl import load_workbook


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


def parse_excel(path: str) -> ParseResult:
    """Read Excel, extract columns A/F/G/H/O/P, return ParseResult."""
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    result = ParseResult()

    for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if len(row) < 16:
            continue

        program = row[0]   # A
        name1 = row[5]     # F
        name2 = row[6]     # G
        qty = row[7]       # H
        com_no = row[14]   # O
        m_val = row[15]    # P

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
