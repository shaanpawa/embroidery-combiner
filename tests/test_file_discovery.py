"""Tests for file discovery module."""

import os
import pytest

from app.core.file_discovery import (
    discover_folder, extract_number, generate_output_name,
    check_sequence_gaps, check_duplicates, DiscoveredFile,
)


class TestExtractNumber:
    def test_simple(self):
        assert extract_number("216.dst") == 216

    def test_with_prefix(self):
        assert extract_number("design_42.dst") == 42

    def test_no_number(self):
        assert extract_number("flower.dst") is None

    def test_multiple_numbers(self):
        # Takes first number found
        assert extract_number("batch2_design216.ngs") == 2


class TestDiscoverFolder:
    def test_empty_folder(self, tmp_folder):
        result = discover_folder(tmp_folder)
        assert result.total_files == 0
        assert not result.files

    def test_dst_only(self, tmp_folder, make_dst):
        make_dst("216.dst", width=200, height=100)
        make_dst("217.dst", width=250, height=130)
        result = discover_folder(tmp_folder)
        assert result.dst_count == 2
        assert result.ngs_count == 0
        assert not result.needs_conversion

    def test_sorts_by_number(self, tmp_folder, make_dst):
        make_dst("218.dst")
        make_dst("216.dst")
        make_dst("217.dst")
        result = discover_folder(tmp_folder)
        numbers = [f.number for f in result.files]
        assert numbers == [216, 217, 218]

    def test_non_embroidery_skipped(self, tmp_folder, make_dst):
        make_dst("216.dst")
        # Create a non-embroidery file
        with open(os.path.join(tmp_folder, "readme.txt"), 'w') as f:
            f.write("hello")
        result = discover_folder(tmp_folder)
        assert result.total_files == 1
        assert len(result.skipped_files) == 1
        assert "readme.txt" in result.skipped_files

    def test_single_file_warning(self, tmp_folder, make_dst):
        make_dst("216.dst")
        result = discover_folder(tmp_folder)
        assert any("Only 1 file" in w for w in result.warnings)

    def test_invalid_folder(self):
        result = discover_folder("/nonexistent/path")
        assert result.total_files == 0
        assert any("Not a valid" in w for w in result.warnings)


class TestSequenceGaps:
    def test_no_gap(self):
        files = [
            DiscoveredFile("", "216.dst", ".dst", 216, 0),
            DiscoveredFile("", "217.dst", ".dst", 217, 0),
            DiscoveredFile("", "218.dst", ".dst", 218, 0),
        ]
        assert check_sequence_gaps(files) == []

    def test_gap_detected(self):
        files = [
            DiscoveredFile("", "216.dst", ".dst", 216, 0),
            DiscoveredFile("", "218.dst", ".dst", 218, 0),
        ]
        warnings = check_sequence_gaps(files)
        assert len(warnings) == 1
        assert "217" in warnings[0]


class TestDuplicates:
    def test_no_duplicates(self):
        files = [
            DiscoveredFile("", "216.dst", ".dst", 216, 0),
            DiscoveredFile("", "217.dst", ".dst", 217, 0),
        ]
        assert check_duplicates(files) == []

    def test_duplicate_detected(self):
        files = [
            DiscoveredFile("", "216.dst", ".dst", 216, 0),
            DiscoveredFile("", "design_216.dst", ".dst", 216, 0),
        ]
        warnings = check_duplicates(files)
        assert len(warnings) == 1
        assert "Duplicate" in warnings[0]


class TestOutputName:
    def test_range(self):
        files = [
            DiscoveredFile("", "216.dst", ".dst", 216, 0),
            DiscoveredFile("", "218.dst", ".dst", 218, 0),
        ]
        assert generate_output_name(files) == "216-218.dst"

    def test_single(self):
        files = [DiscoveredFile("", "216.dst", ".dst", 216, 0)]
        assert generate_output_name(files) == "216.dst"

    def test_no_numbers(self):
        files = [DiscoveredFile("", "flower.dst", ".dst", None, 0)]
        assert generate_output_name(files) == "combined.dst"
