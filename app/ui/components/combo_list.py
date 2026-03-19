"""
Scrollable list of combo groups with checkboxes for selection.
Each group can be expanded to show individual combo files.
"""

import customtkinter as ctk

from app.core.excel_parser import ComboFile, ComboGroup
from app.ui.theme import COLORS, FONTS, PAD_SM, PAD_MD, PAD_XS, RADIUS_SM


class ComboList(ctk.CTkScrollableFrame):
    """Displays combo groups with checkboxes for selecting which combos to export."""

    def __init__(self, master, on_select_change=None, **kwargs):
        theme = COLORS.get(ctk.get_appearance_mode().lower(), COLORS["dark"])
        super().__init__(
            master,
            fg_color=theme["bg_secondary"],
            corner_radius=RADIUS_SM,
            **kwargs,
        )
        self._on_select_change = on_select_change
        self._combo_vars = {}  # ComboFile -> BooleanVar
        self._group_widgets = []
        self._combos = []

    def clear(self):
        for w in self.winfo_children():
            w.destroy()
        self._combo_vars.clear()
        self._group_widgets.clear()
        self._combos.clear()

    def populate(self, groups: list, combo_files_by_group: dict):
        """
        Populate the list with combo groups.
        groups: list of ComboGroup
        combo_files_by_group: dict mapping group_key -> list of ComboFile
        """
        self.clear()
        theme = COLORS.get(ctk.get_appearance_mode().lower(), COLORS["dark"])

        for group in groups:
            combos = combo_files_by_group.get(group.group_key, [])
            if not combos:
                continue

            self._combos.extend(combos)

            # Group header
            group_frame = ctk.CTkFrame(self, fg_color=theme["bg_surface"], corner_radius=RADIUS_SM)
            group_frame.pack(fill="x", pady=(0, PAD_XS))

            header = ctk.CTkFrame(group_frame, fg_color="transparent")
            header.pack(fill="x", padx=PAD_SM, pady=(PAD_SM, PAD_XS))

            n_files = len(combos)
            total_slots = sum(len(c.slots) for c in combos)
            ctk.CTkLabel(
                header,
                text=f"{group.machine_program} / Com {group.com_no}",
                font=FONTS["subheading"],
                text_color=theme["text_primary"],
            ).pack(side="left")

            ctk.CTkLabel(
                header,
                text=f"  {n_files} file{'s' if n_files != 1 else ''}, {total_slots} slots",
                font=FONTS["small"],
                text_color=theme["text_secondary"],
            ).pack(side="left")

            # Individual combo files within this group
            for combo in combos:
                var = ctk.BooleanVar(value=True)
                self._combo_vars[id(combo)] = (combo, var)

                row = ctk.CTkFrame(group_frame, fg_color="transparent")
                row.pack(fill="x", padx=(PAD_MD, PAD_SM), pady=(0, PAD_XS))

                cb = ctk.CTkCheckBox(
                    row,
                    text="",
                    variable=var,
                    width=20,
                    height=20,
                    checkbox_width=18,
                    checkbox_height=18,
                    fg_color=theme["accent"],
                    hover_color=theme["accent_hover"],
                    border_color=theme["border"],
                    command=self._on_toggle,
                )
                cb.pack(side="left", padx=(0, PAD_SM))

                # File name
                ctk.CTkLabel(
                    row,
                    text=combo.filename,
                    font=FONTS["small"],
                    text_color=theme["text_primary"],
                ).pack(side="left")

                # Slot count
                n_left = len(combo.left_column)
                n_right = len(combo.right_column)
                slot_text = f"{n_left}L"
                if n_right:
                    slot_text += f" + {n_right}R"
                ctk.CTkLabel(
                    row,
                    text=slot_text,
                    font=FONTS["tiny"],
                    text_color=theme["text_muted"],
                ).pack(side="right")

            # Bottom padding
            ctk.CTkFrame(group_frame, fg_color="transparent", height=PAD_XS).pack()

            self._group_widgets.append(group_frame)

    def _on_toggle(self):
        if self._on_select_change:
            self._on_select_change()

    def get_selected_combos(self) -> list:
        """Return list of selected ComboFile objects."""
        selected = []
        for combo, var in self._combo_vars.values():
            if var.get():
                selected.append(combo)
        return selected

    def select_all(self):
        for _, var in self._combo_vars.values():
            var.set(True)
        self._on_toggle()

    def deselect_all(self):
        for _, var in self._combo_vars.values():
            var.set(False)
        self._on_toggle()

    @property
    def total_combos(self) -> int:
        return len(self._combo_vars)

    @property
    def selected_count(self) -> int:
        return sum(1 for _, var in self._combo_vars.values() if var.get())
