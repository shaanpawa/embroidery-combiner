"""Tests for config module."""

import json
import os
import pytest

from app.config import Config, DEFAULT_GAP_MM


class TestConfig:
    def test_defaults(self):
        config = Config()
        assert config.gap_mm == DEFAULT_GAP_MM
        assert config.theme == "dark"
        assert config.last_folder == ""

    def test_save_and_load(self, tmp_path):
        config = Config()
        config._path = os.path.join(str(tmp_path), "settings.json")
        config.gap_mm = 5.0
        config.theme = "light"
        config.last_folder = "/some/folder"
        config.save()

        config2 = Config()
        config2._path = config._path
        config2.load()
        assert config2.gap_mm == 5.0
        assert config2.theme == "light"
        assert config2.last_folder == "/some/folder"

    def test_missing_config(self, tmp_path):
        config = Config()
        config._path = os.path.join(str(tmp_path), "nonexistent.json")
        config.load()  # Should not raise
        assert config.gap_mm == DEFAULT_GAP_MM

    def test_corrupt_config(self, tmp_path):
        path = os.path.join(str(tmp_path), "settings.json")
        with open(path, 'w') as f:
            f.write("NOT VALID JSON {{{{")
        config = Config()
        config._path = path
        config.load()  # Should not raise
        assert config.gap_mm == DEFAULT_GAP_MM
