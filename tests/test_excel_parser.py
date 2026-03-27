"""Tests for Excel order parser."""

import os
import pytest
from openpyxl import Workbook

from app.core.excel_parser import (
    ComboFile, ComboGroup, NameEntry, ParseResult,
    expand_and_split, expand_and_split_with_heads,
    generate_all_combos, group_entries, parse_excel,
)


@pytest.fixture
def make_excel(tmp_path):
    """Factory to create test Excel files with specified rows."""
    def _make(rows, filename="test_order.xlsx"):
        wb = Workbook()
        ws = wb.active
        # Header row matching real Excel structure (16 columns A-P)
        ws.append([
            "Program", "order number", "commission", "artikelnr customer",
            "ASA-Artikelnr", "row 1", "row 2", "quantity",
            "Fabric color", "frame color", "name/embroidery color",
            "Velcro color", "size", "Program", "Com No", "M",
        ])
        for row in rows:
            ws.append(row)
        path = str(tmp_path / filename)
        wb.save(path)
        return path
    return _make


def _row(program, name1, name2, qty, com_no, m_val):
    """Helper to build a 16-column row from the fields we care about."""
    return [
        program, None, None, None, None,  # A-E
        name1, name2, qty,                # F-H
        None, None, None, None, None,     # I-M
        program, com_no, m_val,           # N-P
    ]


# --- parse_excel tests ---

class TestParseExcel:
    def test_basic(self, make_excel):
        path = make_excel([
            _row(1, "JOHN DOE", "Engineer", 2, 1, "MA53344"),
            _row(2, "JANE DOE", "", 1, 1, "MA53344"),
        ])
        result = parse_excel(path)
        assert len(result.entries) == 2
        assert result.entries[0].program == 1
        assert result.entries[0].name_line1 == "JOHN DOE"
        assert result.entries[0].name_line2 == "Engineer"
        assert result.entries[0].quantity == 2
        assert result.entries[0].com_no == "1"
        assert result.entries[0].machine_program == "MA53344"
        assert result.entries[1].name_line2 == ""

    def test_missing_program_skipped(self, make_excel):
        path = make_excel([
            _row(None, "NO PROG", "", 1, 1, "MA53344"),
            _row(1, "HAS PROG", "", 1, 1, "MA53344"),
        ])
        result = parse_excel(path)
        assert len(result.entries) == 1
        assert len(result.warnings) == 1

    def test_missing_m_value_skipped(self, make_excel):
        path = make_excel([
            _row(1, "NO M", "", 1, 1, None),
            _row(2, "HAS M", "", 1, 1, "MA53344"),
        ])
        result = parse_excel(path)
        assert len(result.entries) == 1

    def test_qty_zero_defaults_to_one(self, make_excel):
        path = make_excel([_row(1, "NAME", "", 0, 1, "MA53344")])
        result = parse_excel(path)
        assert result.entries[0].quantity == 1
        assert len(result.warnings) == 1

    def test_qty_none_defaults_to_one(self, make_excel):
        path = make_excel([_row(1, "NAME", "", None, 1, "MA53344")])
        result = parse_excel(path)
        assert result.entries[0].quantity == 1

    def test_empty_excel(self, make_excel):
        path = make_excel([])
        result = parse_excel(path)
        assert len(result.entries) == 0

    def test_non_numeric_program_skipped(self, make_excel):
        path = make_excel([_row("abc", "NAME", "", 1, 1, "MA53344")])
        result = parse_excel(path)
        assert len(result.entries) == 0
        assert len(result.warnings) == 1


# --- group_entries tests ---

class TestGroupEntries:
    def test_single_group(self):
        entries = [
            NameEntry(3, "C", "", 1, "1", "MA53344"),
            NameEntry(1, "A", "", 1, "1", "MA53344"),
            NameEntry(2, "B", "", 1, "1", "MA53344"),
        ]
        groups = group_entries(entries)
        assert len(groups) == 1
        assert groups[0].machine_program == "MA53344"
        assert groups[0].com_no == "1"
        # Sorted by program number
        assert [e.program for e in groups[0].entries] == [1, 2, 3]

    def test_multiple_groups(self):
        entries = [
            NameEntry(1, "A", "", 1, "1", "MA53344"),
            NameEntry(2, "B", "", 1, "2", "MA53344"),
            NameEntry(3, "C", "", 1, "1", "MA55451"),
        ]
        groups = group_entries(entries)
        assert len(groups) == 3
        keys = [g.group_key for g in groups]
        assert ("MA53344", "1") in keys
        assert ("MA53344", "2") in keys
        assert ("MA55451", "1") in keys

    def test_total_slots(self):
        entries = [
            NameEntry(1, "A", "", 2, "1", "MA53344"),
            NameEntry(2, "B", "", 3, "1", "MA53344"),
        ]
        groups = group_entries(entries)
        assert groups[0].total_slots == 5


# --- expand_and_split tests ---

class TestExpandAndSplit:
    def test_small_group_single_combo(self):
        group = ComboGroup("MA53344", "1", [
            NameEntry(i, f"Name{i}", "", 1, "1", "MA53344")
            for i in range(1, 6)
        ])
        combos = expand_and_split(group)
        assert len(combos) == 1
        assert combos[0].part_number == 1
        assert combos[0].total_parts == 1
        assert len(combos[0].slots) == 5

    def test_exactly_20_slots(self):
        group = ComboGroup("MA53344", "1", [
            NameEntry(i, f"Name{i}", "", 1, "1", "MA53344")
            for i in range(1, 21)
        ])
        combos = expand_and_split(group)
        assert len(combos) == 1
        assert len(combos[0].slots) == 20

    def test_21_slots_splits_into_two(self):
        group = ComboGroup("MA53344", "1", [
            NameEntry(i, f"Name{i}", "", 1, "1", "MA53344")
            for i in range(1, 22)
        ])
        combos = expand_and_split(group)
        assert len(combos) == 2
        assert len(combos[0].slots) == 20
        assert len(combos[1].slots) == 1
        assert combos[0].part_number == 1
        assert combos[0].total_parts == 2
        assert combos[1].part_number == 2

    def test_quantity_expansion(self):
        group = ComboGroup("MA53344", "1", [
            NameEntry(1, "A", "", 3, "1", "MA53344"),
            NameEntry(2, "B", "", 2, "1", "MA53344"),
        ])
        combos = expand_and_split(group)
        assert len(combos) == 1
        # qty=3 sorts first (higher qty), then qty=2
        assert len(combos[0].slots) == 5
        # First 3 slots should be entry A (qty=3), next 2 entry B (qty=2)
        assert combos[0].slots[0].program == 1
        assert combos[0].slots[2].program == 1
        assert combos[0].slots[3].program == 2

    def test_quantity_causes_split(self):
        group = ComboGroup("MA53344", "1", [
            NameEntry(i, f"Name{i}", "", 2, "1", "MA53344")
            for i in range(1, 12)  # 11 entries * qty 2 = 22 slots
        ])
        combos = expand_and_split(group)
        assert len(combos) == 2
        assert len(combos[0].slots) == 20
        assert len(combos[1].slots) == 2

    def test_left_right_columns(self):
        group = ComboGroup("MA53344", "1", [
            NameEntry(i, f"Name{i}", "", 1, "1", "MA53344")
            for i in range(1, 16)  # 15 entries
        ])
        combos = expand_and_split(group)
        assert len(combos[0].left_column) == 10
        assert len(combos[0].right_column) == 5

    def test_empty_group(self):
        group = ComboGroup("MA53344", "1", [])
        combos = expand_and_split(group)
        assert len(combos) == 0

    def test_filename_format(self):
        group = ComboGroup("MA53344", "1", [
            NameEntry(i, f"Name{i}", "", 1, "1", "MA53344")
            for i in range(1, 25)  # 24 entries -> 2 files
        ])
        combos = expand_and_split(group)
        assert combos[0].filename == "MA53344_Com1_1of2.dst"
        assert combos[1].filename == "MA53344_Com1_2of2.dst"

    def test_same_quantity_stays_together(self):
        """Names with same quantity should cluster together."""
        group = ComboGroup("MA53344", "1", [
            NameEntry(1, "A", "", 1, "1", "MA53344"),
            NameEntry(2, "B", "", 2, "1", "MA53344"),
            NameEntry(3, "C", "", 1, "1", "MA53344"),
            NameEntry(4, "D", "", 2, "1", "MA53344"),
        ])
        combos = expand_and_split(group)
        # qty=2 entries first (B, D), then qty=1 entries (A, C)
        slots = combos[0].slots
        assert slots[0].program == 2  # B, qty=2
        assert slots[1].program == 2  # B, qty=2 (repeated)
        assert slots[2].program == 4  # D, qty=2
        assert slots[3].program == 4  # D, qty=2 (repeated)
        assert slots[4].program == 1  # A, qty=1
        assert slots[5].program == 3  # C, qty=1


# --- generate_all_combos tests ---

class TestGenerateAllCombos:
    def test_full_pipeline(self):
        entries = [
            NameEntry(1, "A", "", 1, "1", "MA53344"),
            NameEntry(2, "B", "", 1, "1", "MA53344"),
            NameEntry(3, "C", "", 1, "2", "MA53344"),
            NameEntry(4, "D", "", 1, "1", "MA55451"),
        ]
        combos = generate_all_combos(entries)
        assert len(combos) == 3  # (MA53344,1), (MA53344,2), (MA55451,1)

    def test_large_group_splits(self):
        entries = [
            NameEntry(i, f"Name{i}", "", 1, "1", "MA53344")
            for i in range(1, 51)  # 50 entries
        ]
        combos = generate_all_combos(entries)
        assert len(combos) == 3  # 20 + 20 + 10
        assert combos[0].total_parts == 3


# --- 2-HEAD optimization tests ---

class TestExpandAndSplitWithHeads:
    """Tests for expand_and_split_with_heads() — 2-HEAD even quantity optimization."""

    def test_even_qty2_single_slot(self):
        """qty=2 → 1 slot, 2-HEAD."""
        group = ComboGroup("MA53345", "1", [
            NameEntry(1, "JACK", "", 2, "1", "MA53345"),
        ])
        combos = expand_and_split_with_heads(group)
        assert len(combos) == 1
        assert combos[0].head_mode == "2-HEAD"
        assert len(combos[0].slots) == 1
        assert combos[0].slots[0].name_line1 == "JACK"

    def test_even_qty4_two_slots(self):
        """qty=4 → 2 slots, 2-HEAD."""
        group = ComboGroup("MA53345", "1", [
            NameEntry(1, "STEVE", "", 4, "1", "MA53345"),
        ])
        combos = expand_and_split_with_heads(group)
        assert len(combos) == 1
        assert combos[0].head_mode == "2-HEAD"
        assert len(combos[0].slots) == 2

    def test_even_qty20_ten_slots(self):
        """qty=20 → 10 slots, 2-HEAD."""
        group = ComboGroup("MA53345", "1", [
            NameEntry(1, "WILL", "", 20, "1", "MA53345"),
        ])
        combos = expand_and_split_with_heads(group)
        assert len(combos) == 1
        assert combos[0].head_mode == "2-HEAD"
        assert len(combos[0].slots) == 10

    def test_even_qty80_two_files(self):
        """qty=80 → 40 slots → 2 files of 20 each, both 2-HEAD."""
        group = ComboGroup("MA53345", "1", [
            NameEntry(1, "BIG ORDER", "", 80, "1", "MA53345"),
        ])
        combos = expand_and_split_with_heads(group)
        even_combos = [c for c in combos if c.head_mode == "2-HEAD"]
        assert len(even_combos) == 2
        assert len(even_combos[0].slots) == 20
        assert len(even_combos[1].slots) == 20
        assert even_combos[0].total_parts == 2
        assert even_combos[1].total_parts == 2

    def test_odd_qty1(self):
        """qty=1 → 1 slot, 1-HEAD."""
        group = ComboGroup("MA53345", "1", [
            NameEntry(1, "SOLO", "", 1, "1", "MA53345"),
        ])
        combos = expand_and_split_with_heads(group)
        assert len(combos) == 1
        assert combos[0].head_mode == "1-HEAD"
        assert len(combos[0].slots) == 1

    def test_odd_qty3(self):
        """qty=3 → 3 slots, 1-HEAD."""
        group = ComboGroup("MA53345", "1", [
            NameEntry(1, "THREE", "", 3, "1", "MA53345"),
        ])
        combos = expand_and_split_with_heads(group)
        assert len(combos) == 1
        assert combos[0].head_mode == "1-HEAD"
        assert len(combos[0].slots) == 3

    def test_mixed_even_odd_separate_files(self):
        """qty=2 + qty=3 in same group → separate EVEN and ODD files."""
        group = ComboGroup("MA53345", "1", [
            NameEntry(1, "JACK", "", 2, "1", "MA53345"),
            NameEntry(2, "STEVE", "", 2, "1", "MA53345"),
            NameEntry(3, "WILL", "", 2, "1", "MA53345"),
            NameEntry(4, "ODD_GUY", "", 3, "1", "MA53345"),
        ])
        combos = expand_and_split_with_heads(group)
        even_combos = [c for c in combos if c.head_mode == "2-HEAD"]
        odd_combos = [c for c in combos if c.head_mode == "1-HEAD"]
        assert len(even_combos) >= 1
        assert len(odd_combos) >= 1
        # Even: 3 names * qty2 / 2 = 3 slots
        even_slots = sum(len(c.slots) for c in even_combos)
        assert even_slots == 3
        # Odd: 1 name * qty3 = 3 slots
        odd_slots = sum(len(c.slots) for c in odd_combos)
        assert odd_slots == 3

    def test_filename_even_format(self):
        """EVEN file includes Qty label."""
        group = ComboGroup("MA53345", "1", [
            NameEntry(1, "JACK", "", 2, "1", "MA53345"),
        ])
        combos = expand_and_split_with_heads(group)
        assert combos[0].filename == "MA53345_Com1_EVEN_Qty2_1of1.dst"

    def test_filename_odd_format(self):
        """ODD file has _ODD tag, no Qty label."""
        group = ComboGroup("MA53345", "1", [
            NameEntry(1, "SOLO", "", 1, "1", "MA53345"),
        ])
        combos = expand_and_split_with_heads(group)
        assert combos[0].filename == "MA53345_Com1_ODD_1of1.dst"

    def test_filename_legacy_format(self):
        """Legacy mode (no head optimization) has no tag."""
        group = ComboGroup("MA53345", "1", [
            NameEntry(1, "LEGACY", "", 1, "1", "MA53345"),
        ])
        combos = expand_and_split(group)  # legacy function
        assert combos[0].filename == "MA53345_Com1_1of1.dst"
        assert combos[0].head_mode == ""

    def test_generate_all_combos_flag_off(self):
        """optimize_heads=False → legacy behavior, no head_mode tags."""
        entries = [
            NameEntry(1, "A", "", 2, "1", "MA53345"),
            NameEntry(2, "B", "", 3, "1", "MA53345"),
        ]
        combos = generate_all_combos(entries, optimize_heads=False)
        assert all(c.head_mode == "" for c in combos)
        # Legacy: 2+3=5 total slots
        assert sum(len(c.slots) for c in combos) == 5

    def test_generate_all_combos_flag_on(self):
        """optimize_heads=True → uses 2-HEAD split."""
        entries = [
            NameEntry(1, "A", "", 2, "1", "MA53345"),
            NameEntry(2, "B", "", 3, "1", "MA53345"),
        ]
        combos = generate_all_combos(entries, optimize_heads=True)
        even_combos = [c for c in combos if c.head_mode == "2-HEAD"]
        odd_combos = [c for c in combos if c.head_mode == "1-HEAD"]
        assert len(even_combos) >= 1  # A (qty=2) → even
        assert len(odd_combos) >= 1   # B (qty=3) → odd
        # Even: 2/2 = 1 slot. Odd: 3 slots. Total: 4 (saved 1)
        total_slots = sum(len(c.slots) for c in combos)
        assert total_slots == 4
