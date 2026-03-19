"""Tests for pipeline orchestrator."""

import os
import pytest

from app.core.excel_parser import ComboFile, NameEntry
from app.core.pipeline import check_combo_ready, export_combo, export_all, resolve_dst_files


def _entry(prog, qty=1):
    return NameEntry(prog, f"Name{prog}", "", qty, "1", "MA53344")


def _combo(programs, part=1, total=1):
    slots = [_entry(p) for p in programs]
    return ComboFile("MA53344", "1", part, total, slots)


class TestResolveDstFiles:
    def test_all_found(self, make_dst):
        combo = _combo([1, 2, 3])
        # Create DST files named by program number
        for i in [1, 2, 3]:
            make_dst(f"{i}.dst")
        paths, missing = resolve_dst_files(combo, str(make_dst.__self__ if hasattr(make_dst, '__self__') else ""))
        # Use tmp_path from make_dst fixture
        # We need to get the tmp_path... let's use the directory of the first file
        first_dst = make_dst("1.dst")
        dst_folder = os.path.dirname(first_dst)
        paths, missing = resolve_dst_files(combo, dst_folder)
        assert len(missing) == 0

    def test_missing_files(self, make_dst):
        combo = _combo([1, 2, 99])
        make_dst("1.dst")
        make_dst("2.dst")
        dst_folder = os.path.dirname(make_dst("1.dst"))
        paths, missing = resolve_dst_files(combo, dst_folder)
        assert missing == [99]


class TestCheckComboReady:
    def test_ready(self, make_dst):
        combo = _combo([1, 2])
        f1 = make_dst("1.dst")
        make_dst("2.dst")
        dst_folder = os.path.dirname(f1)
        ready, missing = check_combo_ready(combo, dst_folder)
        assert ready is True
        assert missing == []

    def test_not_ready(self, make_dst):
        combo = _combo([1, 2, 3])
        f1 = make_dst("1.dst")
        dst_folder = os.path.dirname(f1)
        ready, missing = check_combo_ready(combo, dst_folder)
        assert ready is False
        assert 2 in missing
        assert 3 in missing


class TestExportCombo:
    def test_export_success(self, make_dst, tmp_path):
        combo = _combo([1, 2, 3])
        f1 = make_dst("1.dst")
        make_dst("2.dst")
        make_dst("3.dst")
        dst_folder = os.path.dirname(f1)
        output_folder = str(tmp_path / "output")
        os.makedirs(output_folder)

        result = export_combo(combo, dst_folder, output_folder)
        assert result.success is True
        assert os.path.isfile(result.output_path)
        assert result.validation.get("valid") is True

    def test_export_missing_files(self, make_dst, tmp_path):
        combo = _combo([1, 99])
        f1 = make_dst("1.dst")
        dst_folder = os.path.dirname(f1)
        output_folder = str(tmp_path / "output")
        os.makedirs(output_folder)

        result = export_combo(combo, dst_folder, output_folder)
        assert result.success is False
        assert "99" in result.error

    def test_export_two_column(self, make_dst, tmp_path):
        """Export a combo that uses both columns (>10 slots)."""
        programs = list(range(1, 16))  # 15 programs
        combo = _combo(programs)
        for p in programs:
            make_dst(f"{p}.dst")
        dst_folder = os.path.dirname(make_dst("1.dst"))
        output_folder = str(tmp_path / "output")
        os.makedirs(output_folder)

        result = export_combo(combo, dst_folder, output_folder)
        assert result.success is True
        # Output should be wider than single column
        assert result.validation["width_mm"] > 25


class TestExportAll:
    def test_multiple_combos(self, make_dst, tmp_path):
        combo1 = _combo([1, 2], part=1, total=2)
        combo2 = _combo([3, 4], part=2, total=2)
        for i in range(1, 5):
            make_dst(f"{i}.dst")
        dst_folder = os.path.dirname(make_dst("1.dst"))
        output_folder = str(tmp_path / "output")
        os.makedirs(output_folder)

        progress_calls = []
        results = export_all(
            [combo1, combo2], dst_folder, output_folder,
            progress_callback=lambda c, t: progress_calls.append((c, t)),
        )
        assert len(results) == 2
        assert all(r.success for r in results)
        assert len(progress_calls) == 2
