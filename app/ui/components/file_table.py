"""
Scrollable file list with checkboxes and per-file status badges.
"""

import customtkinter as ctk
from typing import Callable, List, Optional

from app.core.file_discovery import DiscoveredFile
from app.ui.theme import COLORS, FONTS, PAD_SM, PAD_XS, PAD_MD, RADIUS_SM


class StatusBadge(ctk.CTkLabel):
    """Colored status indicator."""

    def __init__(self, parent, **kwargs):
        super().__init__(
            parent,
            text="",
            font=FONTS["tiny"],
            corner_radius=RADIUS_SM,
            height=22,
            **kwargs,
        )

    def set_status(self, status: str, detail: str = "", level: str = "ok"):
        theme = COLORS.get(ctk.get_appearance_mode().lower(), COLORS["dark"])

        color_map = {
            "ok": theme["success"],
            "warning": theme["warning"],
            "error": theme["error"],
            "converting": theme["accent"],
            "waiting": theme["text_muted"],
            "done": theme["success"],
        }

        fg = color_map.get(level, theme["text_secondary"])
        text = detail if detail else status

        # Truncate long text
        if len(text) > 40:
            text = text[:37] + "..."

        self.configure(text=f" {text} ", text_color=fg)


class FileRow(ctk.CTkFrame):
    """Single row in the file table."""

    def __init__(
        self,
        parent,
        file: DiscoveredFile,
        index: int,
        on_toggle: Optional[Callable] = None,
        **kwargs,
    ):
        theme = COLORS.get(ctk.get_appearance_mode().lower(), COLORS["dark"])
        bg = theme["bg_secondary"] if index % 2 == 0 else theme["bg_primary"]
        super().__init__(parent, fg_color=bg, height=36, **kwargs)
        self.pack_propagate(False)

        self.file = file
        self._on_toggle = on_toggle

        # Checkbox
        self.check_var = ctk.BooleanVar(value=file.included)
        self.checkbox = ctk.CTkCheckBox(
            self,
            text="",
            variable=self.check_var,
            width=24,
            height=24,
            command=self._toggled,
            fg_color=theme["accent"],
            hover_color=theme["accent_hover"],
        )
        self.checkbox.pack(side="left", padx=(PAD_SM, PAD_XS))

        # Number
        num_text = str(file.number) if file.number is not None else "—"
        ctk.CTkLabel(
            self,
            text=num_text,
            font=FONTS["small"],
            text_color=theme["text_muted"],
            width=35,
            anchor="e",
        ).pack(side="left", padx=(0, PAD_SM))

        # Filename
        ctk.CTkLabel(
            self,
            text=file.filename,
            font=FONTS["body"],
            text_color=theme["text_primary"],
            anchor="w",
        ).pack(side="left", padx=(0, PAD_SM), fill="x", expand=True)

        # Size
        ctk.CTkLabel(
            self,
            text=file.size_display,
            font=FONTS["small"],
            text_color=theme["text_muted"],
            width=60,
            anchor="e",
        ).pack(side="left", padx=(0, PAD_SM))

        # Status badge
        self.badge = StatusBadge(self)
        self.badge.pack(side="left", padx=(0, PAD_SM))
        self.badge.set_status("Ready", level="waiting")

    def _toggled(self):
        self.file.included = self.check_var.get()
        if self._on_toggle:
            self._on_toggle()


class FileTable(ctk.CTkScrollableFrame):
    """Scrollable file list with per-file checkboxes and status."""

    def __init__(self, parent, on_toggle: Optional[Callable] = None, **kwargs):
        theme = COLORS.get(ctk.get_appearance_mode().lower(), COLORS["dark"])
        super().__init__(
            parent,
            fg_color=theme["bg_primary"],
            corner_radius=RADIUS_SM,
            **kwargs,
        )
        self._rows: List[FileRow] = []
        self._on_toggle = on_toggle

    def populate(self, files: List[DiscoveredFile]):
        """Fill the table with discovered files."""
        self.clear()
        for i, f in enumerate(files):
            row = FileRow(self, f, i, on_toggle=self._on_toggle)
            row.pack(fill="x", pady=1)
            self._rows.append(row)

    def update_status(self, index: int, status: str, detail: str = "", level: str = "ok"):
        """Update a file's status badge."""
        if 0 <= index < len(self._rows):
            self._rows[index].badge.set_status(status, detail, level)

    def get_included_files(self) -> List[DiscoveredFile]:
        """Return files that are checked/included."""
        return [r.file for r in self._rows if r.check_var.get()]

    def clear(self):
        for row in self._rows:
            row.destroy()
        self._rows.clear()

    def set_all_included(self, included: bool):
        for row in self._rows:
            row.check_var.set(included)
            row.file.included = included
