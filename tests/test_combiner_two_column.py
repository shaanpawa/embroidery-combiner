"""Tests for two-column combine function."""

import os
import pytest
import pyembroidery

from app.core.combiner import CombineError, combine_designs_two_column


class TestTwoColumnCombine:
    def test_left_only(self, make_dst):
        f1 = make_dst("1.dst", width=200, height=100)
        f2 = make_dst("2.dst", width=200, height=100)
        result = combine_designs_two_column([f1, f2], [])
        ext = result.extents()
        # Should be roughly same width as single column
        assert ext[2] - ext[0] == pytest.approx(200, abs=10)

    def test_right_only(self, make_dst):
        f1 = make_dst("1.dst", width=200, height=100)
        result = combine_designs_two_column([], [f1])
        assert len(result.stitches) > 0

    def test_both_columns(self, make_dst):
        left = [make_dst(f"l{i}.dst", width=200, height=100) for i in range(3)]
        right = [make_dst(f"r{i}.dst", width=200, height=100) for i in range(2)]
        result = combine_designs_two_column(left, right, gap_mm=3.0, column_gap_mm=10.0)
        ext = result.extents()
        # Width should be roughly 2x design width + column gap
        width = ext[2] - ext[0]
        assert width > 300  # At least wider than single column

    def test_column_gap_affects_width(self, make_dst):
        left = [make_dst("l1.dst", width=200, height=100)]
        right = [make_dst("r1.dst", width=200, height=100)]

        narrow = combine_designs_two_column(left, right, column_gap_mm=5.0)
        wide = combine_designs_two_column(left, right, column_gap_mm=20.0)

        narrow_width = narrow.extents()[2] - narrow.extents()[0]
        wide_width = wide.extents()[2] - wide.extents()[0]
        assert wide_width > narrow_width

    def test_columns_aligned_at_top(self, make_dst):
        left = [make_dst(f"l{i}.dst", width=200, height=100) for i in range(5)]
        right = [make_dst(f"r{i}.dst", width=200, height=100) for i in range(2)]
        result = combine_designs_two_column(left, right)
        ext = result.extents()
        # Top of the pattern (min_y) should be at 0
        assert ext[1] == pytest.approx(0, abs=10)

    def test_roundtrip(self, make_dst, tmp_folder):
        left = [make_dst(f"l{i}.dst", width=200, height=100) for i in range(3)]
        right = [make_dst(f"r{i}.dst", width=200, height=100) for i in range(3)]
        result = combine_designs_two_column(left, right, gap_mm=3.0, column_gap_mm=10.0)

        output_path = os.path.join(tmp_folder, "combo.dst")
        pyembroidery.write(result, output_path)

        readback = pyembroidery.read(output_path)
        assert readback is not None
        assert len(readback.stitches) > 0
        r_ext = readback.extents()
        assert r_ext is not None
        # Should have both columns' width
        assert (r_ext[2] - r_ext[0]) > 300

    def test_empty_raises(self):
        with pytest.raises(CombineError):
            combine_designs_two_column([], [])

    def test_single_file_each(self, make_dst):
        left = [make_dst("l.dst", width=200, height=100)]
        right = [make_dst("r.dst", width=200, height=100)]
        result = combine_designs_two_column(left, right)
        ext = result.extents()
        assert ext is not None
        assert (ext[2] - ext[0]) > 300

    def test_progress_callback(self, make_dst):
        left = [make_dst(f"l{i}.dst") for i in range(3)]
        right = [make_dst(f"r{i}.dst") for i in range(3)]
        calls = []
        combine_designs_two_column(left, right, progress_callback=lambda c, t: calls.append((c, t)))
        assert len(calls) > 0

    def test_full_20_slot_combo(self, make_dst):
        """10 left + 10 right = full 20-slot combo file."""
        left = [make_dst(f"l{i}.dst", width=200, height=100) for i in range(10)]
        right = [make_dst(f"r{i}.dst", width=200, height=100) for i in range(10)]
        result = combine_designs_two_column(left, right, gap_mm=3.0, column_gap_mm=10.0)

        output_path = os.path.join(str(make_dst.__wrapped__) if hasattr(make_dst, '__wrapped__') else "/tmp", "full_combo.dst")
        # Just verify it produces a valid pattern
        ext = result.extents()
        assert ext is not None
        height = ext[3] - ext[1]
        width = ext[2] - ext[0]
        # 10 designs * 100 height + 9 gaps * 30 = 1270 units for each column
        assert height > 1000
        assert width > 300
