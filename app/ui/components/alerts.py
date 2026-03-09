"""
Warning/info banners displayed below the folder selector.
Shows sequence gaps, duplicates, conversion info, etc.
"""

import customtkinter as ctk

from app.ui.theme import COLORS, FONTS, PAD_SM, PAD_XS, RADIUS_SM


class AlertsBanner(ctk.CTkFrame):
    """Horizontal banner showing warning/info pills."""

    def __init__(self, parent, **kwargs):
        super().__init__(parent, fg_color="transparent", **kwargs)
        self._alerts = []
        self._widgets = []

    def set_alerts(self, alerts: list):
        """
        Set alerts to display.

        Args:
            alerts: List of (level, message) tuples.
                    level is "info", "warning", or "error".
        """
        self.clear()
        self._alerts = alerts

        if not alerts:
            self.pack_forget()
            return

        theme = COLORS.get(ctk.get_appearance_mode().lower(), COLORS["dark"])

        for level, message in alerts:
            if level == "error":
                fg = theme["error"]
                bg = theme["bg_surface"]
            elif level == "warning":
                fg = theme["warning"]
                bg = theme["bg_surface"]
            else:
                fg = theme["text_secondary"]
                bg = theme["bg_surface"]

            pill = ctk.CTkLabel(
                self,
                text=f"  {message}  ",
                font=FONTS["small"],
                text_color=fg,
                fg_color=bg,
                corner_radius=RADIUS_SM,
                height=28,
            )
            pill.pack(side="left", padx=(0, PAD_XS), pady=PAD_XS)
            self._widgets.append(pill)

    def clear(self):
        for w in self._widgets:
            w.destroy()
        self._widgets.clear()
        self._alerts.clear()
