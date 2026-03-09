"""
Gap configuration with presets and custom value input.
"""

import customtkinter as ctk
from typing import Callable, Optional

from app.config import GAP_PRESETS, DEFAULT_GAP_MM
from app.ui.theme import COLORS, FONTS, PAD_SM, PAD_XS, RADIUS_SM


class GapControls(ctk.CTkFrame):
    """Gap size selector with preset buttons and custom input."""

    def __init__(
        self,
        parent,
        initial_gap: float = DEFAULT_GAP_MM,
        on_change: Optional[Callable] = None,
        **kwargs,
    ):
        theme = COLORS.get(ctk.get_appearance_mode().lower(), COLORS["dark"])
        super().__init__(parent, fg_color="transparent", **kwargs)
        self._on_change = on_change
        self._active_preset = None

        # Label
        ctk.CTkLabel(
            self,
            text="Gap",
            font=FONTS["subheading"],
            text_color=theme["text_primary"],
        ).pack(side="left", padx=(0, PAD_SM))

        # Custom value entry
        self._gap_var = ctk.StringVar(value=str(initial_gap))
        self._gap_var.trace_add("write", self._on_value_change)

        self._entry = ctk.CTkEntry(
            self,
            textvariable=self._gap_var,
            width=60,
            height=32,
            font=FONTS["body"],
            justify="center",
        )
        self._entry.pack(side="left", padx=(0, PAD_XS))

        ctk.CTkLabel(
            self,
            text="mm",
            font=FONTS["small"],
            text_color=theme["text_muted"],
        ).pack(side="left", padx=(0, PAD_SM))

        # Separator
        ctk.CTkLabel(
            self,
            text="|",
            font=FONTS["body"],
            text_color=theme["border"],
        ).pack(side="left", padx=PAD_SM)

        # Preset buttons
        self._preset_buttons = {}
        for name, value in GAP_PRESETS.items():
            btn = ctk.CTkButton(
                self,
                text=name,
                width=70,
                height=30,
                font=FONTS["small"],
                corner_radius=RADIUS_SM,
                fg_color="transparent",
                text_color=theme["text_secondary"],
                hover_color=theme["bg_hover"],
                command=lambda v=value, n=name: self._select_preset(n, v),
            )
            btn.pack(side="left", padx=2)
            self._preset_buttons[name] = btn

        # Highlight initial preset if it matches
        self._highlight_matching_preset(initial_gap)

    def _select_preset(self, name: str, value: float):
        self._gap_var.set(str(value))
        self._highlight_preset(name)

    def _highlight_preset(self, name: str):
        theme = COLORS.get(ctk.get_appearance_mode().lower(), COLORS["dark"])
        for n, btn in self._preset_buttons.items():
            if n == name:
                btn.configure(
                    fg_color=theme["accent"],
                    text_color=theme["text_primary"],
                )
            else:
                btn.configure(
                    fg_color="transparent",
                    text_color=theme["text_secondary"],
                )
        self._active_preset = name

    def _highlight_matching_preset(self, value: float):
        for name, preset_value in GAP_PRESETS.items():
            if abs(value - preset_value) < 0.01:
                self._highlight_preset(name)
                return
        # No match — clear all
        theme = COLORS.get(ctk.get_appearance_mode().lower(), COLORS["dark"])
        for btn in self._preset_buttons.values():
            btn.configure(fg_color="transparent", text_color=theme["text_secondary"])
        self._active_preset = None

    def _on_value_change(self, *_):
        try:
            val = float(self._gap_var.get())
            if val < 0:
                val = 0
                self._gap_var.set("0")
            elif val > 50:
                val = 50
                self._gap_var.set("50")
            self._highlight_matching_preset(val)
            if self._on_change:
                self._on_change(val)
        except ValueError:
            pass  # Ignore non-numeric input while typing

    def get_gap_mm(self) -> float:
        try:
            val = float(self._gap_var.get())
            return max(0.0, min(50.0, val))
        except ValueError:
            return DEFAULT_GAP_MM

    def set_gap_mm(self, value: float):
        self._gap_var.set(str(value))
