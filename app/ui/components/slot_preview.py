"""
Two-column slot preview for a combo file.
Shows left column (slots 1-10) and right column (slots 11-20) with name details.
"""

import customtkinter as ctk

from app.core.excel_parser import ComboFile
from app.ui.theme import COLORS, FONTS, PAD_SM, PAD_MD, PAD_XS, RADIUS_SM


class SlotPreview(ctk.CTkFrame):
    """Displays the two-column layout preview of a combo file."""

    def __init__(self, master, **kwargs):
        theme = COLORS.get(ctk.get_appearance_mode().lower(), COLORS["dark"])
        super().__init__(master, fg_color=theme["bg_surface"], corner_radius=RADIUS_SM, **kwargs)
        self._theme = theme
        self._build_empty()

    def _build_empty(self):
        for w in self.winfo_children():
            w.destroy()
        ctk.CTkLabel(
            self,
            text="Select a combo file to preview",
            font=FONTS["small"],
            text_color=self._theme["text_muted"],
        ).pack(pady=PAD_MD)

    def show_combo(self, combo: ComboFile, missing_programs: list = None):
        """Display the slot layout for a combo file."""
        for w in self.winfo_children():
            w.destroy()

        missing_set = set(missing_programs or [])
        theme = self._theme

        # Title
        ctk.CTkLabel(
            self,
            text=combo.filename,
            font=FONTS["subheading"],
            text_color=theme["text_primary"],
        ).pack(pady=(PAD_SM, PAD_XS))

        # Two column container
        columns = ctk.CTkFrame(self, fg_color="transparent")
        columns.pack(fill="both", expand=True, padx=PAD_SM, pady=(0, PAD_SM))

        # Left column
        left_frame = ctk.CTkFrame(columns, fg_color="transparent")
        left_frame.pack(side="left", fill="both", expand=True, padx=(0, PAD_SM))

        ctk.CTkLabel(
            left_frame,
            text="Left Column",
            font=FONTS["tiny"],
            text_color=theme["text_muted"],
        ).pack(anchor="w")

        for i, entry in enumerate(combo.left_column):
            self._slot_row(left_frame, i + 1, entry, entry.program in missing_set)

        # Pad left column if fewer than 10
        for i in range(len(combo.left_column), 10):
            self._empty_slot(left_frame, i + 1)

        # Right column
        if combo.right_column:
            right_frame = ctk.CTkFrame(columns, fg_color="transparent")
            right_frame.pack(side="left", fill="both", expand=True)

            ctk.CTkLabel(
                right_frame,
                text="Right Column",
                font=FONTS["tiny"],
                text_color=theme["text_muted"],
            ).pack(anchor="w")

            for i, entry in enumerate(combo.right_column):
                self._slot_row(right_frame, i + 11, entry, entry.program in missing_set)

    def _slot_row(self, parent, slot_num, entry, is_missing=False):
        theme = self._theme
        row = ctk.CTkFrame(parent, fg_color="transparent", height=22)
        row.pack(fill="x", pady=1)
        row.pack_propagate(False)

        color = theme["error"] if is_missing else theme["text_primary"]
        prefix = "!" if is_missing else ""

        ctk.CTkLabel(
            row,
            text=f"{slot_num:2d}.",
            font=FONTS["tiny"],
            text_color=theme["text_muted"],
            width=24,
        ).pack(side="left")

        ctk.CTkLabel(
            row,
            text=f"{prefix}{entry.program}",
            font=FONTS["tiny"],
            text_color=color,
            width=36,
        ).pack(side="left")

        name_text = entry.name_line1
        if entry.name_line2:
            name_text += f" | {entry.name_line2}"

        ctk.CTkLabel(
            row,
            text=name_text,
            font=FONTS["tiny"],
            text_color=color,
            anchor="w",
        ).pack(side="left", fill="x", expand=True)

    def _empty_slot(self, parent, slot_num):
        theme = self._theme
        row = ctk.CTkFrame(parent, fg_color="transparent", height=22)
        row.pack(fill="x", pady=1)
        row.pack_propagate(False)

        ctk.CTkLabel(
            row,
            text=f"{slot_num:2d}.",
            font=FONTS["tiny"],
            text_color=theme["text_muted"],
            width=24,
        ).pack(side="left")

        ctk.CTkLabel(
            row,
            text="—",
            font=FONTS["tiny"],
            text_color=theme["text_muted"],
        ).pack(side="left")
