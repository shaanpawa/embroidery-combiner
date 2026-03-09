"""
Progress display with phase label and progress bar.
"""

import customtkinter as ctk

from app.ui.theme import COLORS, FONTS, PAD_SM, PAD_XS


class ProgressPanel(ctk.CTkFrame):
    """Progress bar with phase label. Hidden when not processing."""

    def __init__(self, parent, **kwargs):
        theme = COLORS.get(ctk.get_appearance_mode().lower(), COLORS["dark"])
        super().__init__(parent, fg_color="transparent", **kwargs)

        self._phase_label = ctk.CTkLabel(
            self,
            text="",
            font=FONTS["small"],
            text_color=theme["text_secondary"],
            anchor="w",
        )
        self._phase_label.pack(fill="x", pady=(0, PAD_XS))

        self._progress = ctk.CTkProgressBar(
            self,
            height=6,
            corner_radius=3,
            progress_color=theme["accent"],
        )
        self._progress.pack(fill="x")
        self._progress.set(0)

        # Start hidden
        self.pack_forget()

    def show(self):
        self.pack(fill="x", pady=(PAD_SM, 0))

    def hide(self):
        self.pack_forget()
        self._progress.set(0)
        self._phase_label.configure(text="")

    def set_phase(self, text: str):
        self._phase_label.configure(text=text)

    def set_progress(self, current: int, total: int):
        if total > 0:
            self._progress.set(current / total)
        else:
            self._progress.set(0)

    def set_indeterminate(self):
        self._progress.set(0.5)

    def set_complete(self):
        self._progress.set(1.0)
