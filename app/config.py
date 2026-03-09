"""
Settings persistence and application constants.
Stores user preferences as JSON next to the executable.
"""

import json
import os
import sys


# Application constants
APP_NAME = "Embroidery Combiner"
APP_VERSION = "1.0.1"

SUPPORTED_EXTENSIONS = {'.ngs', '.dst'}

GAP_PRESETS = {
    "Tight": 1.0,
    "Normal": 3.0,
    "Wide": 5.0,
    "Extra Wide": 10.0,
}

DEFAULT_GAP_MM = 3.0
MAX_FILE_SIZE_MB = 50
CONVERSION_TIMEOUT_SECONDS = 30

CONFIG_FILENAME = "combiner_settings.json"


def _get_app_dir() -> str:
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


class Config:
    """Persistent user settings."""

    def __init__(self):
        self.last_folder: str = ""
        self.gap_mm: float = DEFAULT_GAP_MM
        self.theme: str = "dark"
        self.window_geometry: str = "720x680"
        self.editor_path: str = ""  # Manual Wings/My Editor exe path
        self._path = os.path.join(_get_app_dir(), CONFIG_FILENAME)

    def load(self) -> None:
        if not os.path.exists(self._path):
            return
        try:
            with open(self._path, 'r') as f:
                data = json.load(f)
            self.last_folder = data.get("last_folder", self.last_folder)
            self.gap_mm = float(data.get("gap_mm", self.gap_mm))
            self.theme = data.get("theme", self.theme)
            self.window_geometry = data.get("window_geometry", self.window_geometry)
            self.editor_path = data.get("editor_path", self.editor_path)
        except (json.JSONDecodeError, ValueError, OSError):
            pass  # Use defaults if config is corrupt

    def save(self) -> None:
        data = {
            "last_folder": self.last_folder,
            "gap_mm": self.gap_mm,
            "theme": self.theme,
            "window_geometry": self.window_geometry,
            "editor_path": self.editor_path,
        }
        try:
            with open(self._path, 'w') as f:
                json.dump(data, f, indent=2)
        except OSError:
            pass  # Silently skip if can't write
