"""Tests for file validation module."""

import os
import pytest

from app.core.validator import validate_file, validate_batch, is_valid_for_combining


class TestValidateDST:
    def test_valid_file(self, make_dst):
        path = make_dst("216.dst", width=200, height=100)
        result = validate_file(path)
        assert result.valid
        assert result.stitch_count > 0
        assert result.width_mm > 0
        assert result.height_mm > 0
        assert result.status == "OK"

    def test_empty_file(self, make_empty):
        path = make_empty("empty.dst")
        result = validate_file(path)
        assert not result.valid
        assert "empty" in result.summary.lower()

    def test_corrupt_file(self, make_corrupt):
        path = make_corrupt("corrupt.dst")
        result = validate_file(path)
        assert not result.valid
        assert not is_valid_for_combining(result)

    def test_missing_file(self):
        result = validate_file("/nonexistent/file.dst")
        assert not result.valid
        assert "not found" in result.summary.lower()

    def test_is_valid_for_combining(self, make_dst):
        path = make_dst("good.dst")
        result = validate_file(path)
        assert is_valid_for_combining(result)


class TestValidateBatch:
    def test_batch(self, make_dst, make_corrupt):
        good = make_dst("216.dst")
        bad = make_corrupt("corrupt.dst")
        results = validate_batch([good, bad])
        assert len(results) == 2
        assert results[0].valid
        assert not results[1].valid

    def test_progress_callback(self, make_dst):
        path = make_dst("216.dst")
        calls = []
        validate_batch([path], progress_callback=lambda c, t, r: calls.append((c, t)))
        assert calls == [(1, 1)]
