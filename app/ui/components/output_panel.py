"""
Output file configuration: auto-generated name, path, overwrite warning.
"""

import os

import customtkinter as ctk
from typing import List, Optional

from app.core.file_discovery import DiscoveredFile, generate_output_name
from app.ui.theme import COLORS, FONTS, PAD_SM, PAD_XS, RADIUS_SM


class OutputPanel(ctk.CTkFrame):
    """Output filename and directory controls."""

    def __init__(self, parent, **kwargs):
        theme = COLORS.get(ctk.get_appearance_mode().lower(), COLORS["dark"])
        super().__init__(parent, fg_color="transparent", **kwargs)

        self._output_dir = ""

        # Label
        ctk.CTkLabel(
            self,
            text="Save as",
            font=FONTS["subheading"],
            text_color=theme["text_primary"],
        ).pack(side="left", padx=(0, PAD_SM))

        # Filename entry
        self._name_var = ctk.StringVar(value="combined.dst")
        self._name_entry = ctk.CTkEntry(
            self,
            textvariable=self._name_var,
            width=160,
            height=32,
            font=FONTS["body"],
        )
        self._name_entry.pack(side="left", padx=(0, PAD_SM))

        # Overwrite warning
        self._overwrite_label = ctk.CTkLabel(
            self,
            text="",
            font=FONTS["tiny"],
            text_color=theme["warning"],
        )
        self._overwrite_label.pack(side="left", padx=(0, PAD_SM))

        # Bind to check overwrite on name change
        self._name_var.trace_add("write", self._check_overwrite)

    def set_auto_name(self, files: List[DiscoveredFile]):
        """Auto-generate output name from included files."""
        name = generate_output_name(files)
        self._name_var.set(name)

    def set_output_dir(self, directory: str):
        self._output_dir = directory
        self._check_overwrite()

    def get_output_path(self) -> str:
        name = self._name_var.get().strip()
        if not name:
            name = "combined.dst"
        if not name.lower().endswith('.dst'):
            name += '.dst'
        return os.path.join(self._output_dir, name)

    def check_overwrite(self) -> bool:
        """Returns True if output file already exists."""
        path = self.get_output_path()
        return os.path.exists(path)

    def _check_overwrite(self, *_):
        if self._output_dir and self.check_overwrite():
            self._overwrite_label.configure(text="File exists — will overwrite")
        else:
            self._overwrite_label.configure(text="")
