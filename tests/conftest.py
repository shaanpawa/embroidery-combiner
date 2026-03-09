"""Shared test fixtures for generating test embroidery files."""

import os
import pytest
import pyembroidery


@pytest.fixture
def tmp_folder(tmp_path):
    """Return a temporary directory path."""
    return str(tmp_path)


@pytest.fixture
def make_dst(tmp_path):
    """Factory: create a DST file with a simple rectangular stitch pattern."""
    def _make(name="test.dst", width=200, height=100):
        pattern = pyembroidery.EmbPattern()
        pattern.add_stitch_absolute(pyembroidery.STITCH, 0, 0)
        pattern.add_stitch_absolute(pyembroidery.STITCH, width, 0)
        pattern.add_stitch_absolute(pyembroidery.STITCH, width, height)
        pattern.add_stitch_absolute(pyembroidery.STITCH, 0, height)
        pattern.add_stitch_absolute(pyembroidery.STITCH, 0, 0)
        pattern.end()
        path = os.path.join(str(tmp_path), name)
        pyembroidery.write(pattern, path)
        return path
    return _make


@pytest.fixture
def make_corrupt(tmp_path):
    """Factory: create a corrupt DST file."""
    def _make(name="corrupt.dst"):
        path = os.path.join(str(tmp_path), name)
        with open(path, 'wb') as f:
            f.write(b"NOT A DST FILE AT ALL - CORRUPT DATA")
        return path
    return _make


@pytest.fixture
def make_empty(tmp_path):
    """Factory: create an empty file."""
    def _make(name="empty.dst"):
        path = os.path.join(str(tmp_path), name)
        with open(path, 'wb') as f:
            pass  # 0 bytes
        return path
    return _make
