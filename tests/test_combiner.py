"""Tests for combiner module."""

import os
import pytest
import pyembroidery

from app.core.combiner import (
    CombineError, combine_designs, save_combined, validate_combined_output,
)


class TestCombineDesigns:
    def test_two_files(self, make_dst):
        f1 = make_dst("216.dst", width=200, height=100)
        f2 = make_dst("217.dst", width=250, height=130)
        combined = combine_designs([f1, f2], gap_mm=3.0)
        ext = combined.extents()
        assert ext is not None
        # Combined height should be h1 + gap + h2
        expected_height = 100 + 30 + 130  # 1/10mm units, gap=3mm=30
        actual_height = ext[3] - ext[1]
        assert abs(actual_height - expected_height) < 5  # Allow small rounding

    def test_single_file(self, make_dst):
        f1 = make_dst("216.dst", width=200, height=100)
        combined = combine_designs([f1])
        assert combined is not None
        assert len(combined.stitches) > 0

    def test_empty_list(self):
        with pytest.raises(CombineError, match="No files"):
            combine_designs([])

    def test_bad_file(self, make_corrupt):
        bad = make_corrupt("bad.dst")
        with pytest.raises(CombineError):
            combine_designs([bad])

    def test_zero_gap(self, make_dst):
        f1 = make_dst("216.dst", width=200, height=100)
        f2 = make_dst("217.dst", width=200, height=100)
        combined = combine_designs([f1, f2], gap_mm=0)
        ext = combined.extents()
        expected_height = 200  # 100 + 0 + 100
        actual_height = ext[3] - ext[1]
        assert abs(actual_height - expected_height) < 5

    def test_progress_callback(self, make_dst):
        f1 = make_dst("216.dst")
        f2 = make_dst("217.dst")
        f3 = make_dst("218.dst")
        calls = []
        combine_designs([f1, f2, f3], progress_callback=lambda c, t: calls.append((c, t)))
        assert calls == [(1, 2), (2, 2)]


class TestSaveCombined:
    def test_save(self, make_dst, tmp_path):
        f1 = make_dst("216.dst")
        combined = combine_designs([f1])
        out = os.path.join(str(tmp_path), "output.dst")
        save_combined(combined, out)
        assert os.path.exists(out)
        assert os.path.getsize(out) > 0

    def test_overwrite_protection(self, make_dst, tmp_path):
        f1 = make_dst("216.dst")
        combined = combine_designs([f1])
        out = os.path.join(str(tmp_path), "output.dst")
        save_combined(combined, out)
        with pytest.raises(FileExistsError):
            save_combined(combined, out, overwrite=False)

    def test_overwrite_allowed(self, make_dst, tmp_path):
        f1 = make_dst("216.dst")
        combined = combine_designs([f1])
        out = os.path.join(str(tmp_path), "output.dst")
        save_combined(combined, out)
        save_combined(combined, out, overwrite=True)  # Should not raise


class TestWriteReadRoundtrip:
    """Verify that combined output survives write→read cycle (the critical bug fix)."""

    def test_two_files_roundtrip(self, make_dst, tmp_path):
        f1 = make_dst("216.dst", width=200, height=100)
        f2 = make_dst("217.dst", width=250, height=130)
        combined = combine_designs([f1, f2], gap_mm=3.0)

        mem_count = len(combined.stitches)
        mem_ext = combined.extents()

        out = os.path.join(str(tmp_path), "roundtrip.dst")
        save_combined(combined, out)

        readback = pyembroidery.read(out)
        assert readback is not None
        rb_ext = readback.extents()

        # Stitch count should be close (pyembroidery may add/remove a few during encode)
        assert len(readback.stitches) >= mem_count - 2
        # Dimensions must match
        assert abs((rb_ext[2] - rb_ext[0]) - (mem_ext[2] - mem_ext[0])) < 5
        assert abs((rb_ext[3] - rb_ext[1]) - (mem_ext[3] - mem_ext[1])) < 5

    def test_three_files_roundtrip(self, make_dst, tmp_path):
        f1 = make_dst("a.dst", width=200, height=100)
        f2 = make_dst("b.dst", width=250, height=130)
        f3 = make_dst("c.dst", width=300, height=160)
        combined = combine_designs([f1, f2, f3], gap_mm=3.0)

        mem_ext = combined.extents()
        out = os.path.join(str(tmp_path), "three.dst")
        save_combined(combined, out)

        readback = pyembroidery.read(out)
        rb_ext = readback.extents()

        # Height should span all 3 designs + 2 gaps
        expected_h = 100 + 30 + 130 + 30 + 160  # 450
        actual_h = rb_ext[3] - rb_ext[1]
        assert abs(actual_h - expected_h) < 5

        # File must be larger than a single design
        single_size = os.path.getsize(make_dst("single.dst", width=200, height=100))
        assert os.path.getsize(out) > single_size

    def test_roundtrip_stitch_count_not_just_first_file(self, make_dst, tmp_path):
        """The old bug: only the first design was saved. This test catches it."""
        f1 = make_dst("x.dst", width=100, height=50)
        f2 = make_dst("y.dst", width=100, height=50)
        single = combine_designs([f1])
        single_count = len(single.stitches)

        combined = combine_designs([f1, f2], gap_mm=3.0)
        out = os.path.join(str(tmp_path), "multi.dst")
        save_combined(combined, out)

        readback = pyembroidery.read(out)
        # Combined must have MORE stitches than a single file
        assert len(readback.stitches) > single_count

    def test_ten_files_roundtrip(self, make_dst, tmp_path):
        """Mirrors the real scenario: 10 files (216-225) combined with 3mm gap."""
        # Create 10 files with varying dimensions like real designs
        sizes = [
            (200, 100), (250, 130), (180, 90), (220, 110), (190, 95),
            (170, 85), (210, 105), (230, 115), (240, 120), (200, 100),
        ]
        files = []
        for i, (w, h) in enumerate(sizes):
            files.append(make_dst(f"{216 + i}.dst", width=w, height=h))

        combined = combine_designs(files, gap_mm=3.0)
        mem_ext = combined.extents()
        mem_count = len(combined.stitches)

        out = os.path.join(str(tmp_path), "216-225.dst")
        save_combined(combined, out)

        readback = pyembroidery.read(out)
        rb_ext = readback.extents()

        # All 10 designs must survive the write→read cycle
        assert len(readback.stitches) >= mem_count - 2

        # Height = sum of all heights + 9 gaps (3mm = 30 units each)
        expected_h = sum(h for _, h in sizes) + 9 * 30  # 1220 units
        actual_h = rb_ext[3] - rb_ext[1]
        assert abs(actual_h - expected_h) < 10

        # Width = max width across all designs
        expected_w = max(w for w, _ in sizes)  # 250
        actual_w = rb_ext[2] - rb_ext[0]
        assert abs(actual_w - expected_w) < 5

        # File must be larger than a single design
        single_size = os.path.getsize(files[0])
        assert os.path.getsize(out) > single_size


class TestValidateOutput:
    def test_valid_output(self, make_dst, tmp_path):
        f1 = make_dst("216.dst", width=200, height=100)
        f2 = make_dst("217.dst", width=250, height=130)
        combined = combine_designs([f1, f2])
        out = os.path.join(str(tmp_path), "output.dst")
        save_combined(combined, out)
        info = validate_combined_output(out)
        assert info["valid"]
        assert info["stitch_count"] > 0
        assert info["width_mm"] > 0
        assert info["height_mm"] > 0
