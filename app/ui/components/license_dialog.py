"""
License activation dialog using customtkinter.
"""

import customtkinter as ctk

from app.licensing import get_machine_id, activate_license
from app.ui.theme import COLORS, FONTS, PAD_MD, PAD_LG, PAD_SM, RADIUS_MD


class LicenseDialog(ctk.CTkToplevel):
    """Modal dialog for entering license key on first run."""

    def __init__(self, parent):
        super().__init__(parent)
        self.title("License Activation")
        self.geometry("420x300")
        self.resizable(False, False)
        self.grab_set()
        self.result = False

        theme = COLORS.get(ctk.get_appearance_mode().lower(), COLORS["dark"])
        self.configure(fg_color=theme["bg_primary"])

        # Center on parent
        self.transient(parent)

        frame = ctk.CTkFrame(self, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=PAD_LG, pady=PAD_LG)

        # Title
        ctk.CTkLabel(
            frame,
            text="Embroidery Combiner",
            font=FONTS["heading"],
            text_color=theme["text_primary"],
        ).pack(pady=(0, PAD_SM))

        ctk.CTkLabel(
            frame,
            text="Enter your license key to activate",
            font=FONTS["body"],
            text_color=theme["text_secondary"],
        ).pack(pady=(0, PAD_LG))

        # Machine ID
        machine_id = get_machine_id()
        id_frame = ctk.CTkFrame(frame, fg_color=theme["bg_surface"], corner_radius=RADIUS_MD)
        id_frame.pack(fill="x", pady=(0, PAD_SM))

        ctk.CTkLabel(
            id_frame,
            text="Machine ID",
            font=FONTS["tiny"],
            text_color=theme["text_muted"],
        ).pack(anchor="w", padx=PAD_MD, pady=(PAD_SM, 0))

        id_label = ctk.CTkLabel(
            id_frame,
            text=machine_id,
            font=FONTS["mono"],
            text_color=theme["text_primary"],
        )
        id_label.pack(anchor="w", padx=PAD_MD, pady=(0, PAD_SM))

        ctk.CTkLabel(
            frame,
            text="Send this ID to your administrator for a key",
            font=FONTS["tiny"],
            text_color=theme["text_muted"],
        ).pack(pady=(0, PAD_MD))

        # License key input
        self._key_var = ctk.StringVar()
        self._key_entry = ctk.CTkEntry(
            frame,
            textvariable=self._key_var,
            placeholder_text="XXXX-XXXX-XXXX-XXXX",
            height=40,
            font=FONTS["mono"],
        )
        self._key_entry.pack(fill="x", pady=(0, PAD_MD))
        self._key_entry.focus_set()
        self._key_entry.bind("<Return>", lambda e: self._activate())

        # Error label (hidden initially)
        self._error_label = ctk.CTkLabel(
            frame,
            text="",
            font=FONTS["small"],
            text_color=theme["error"],
        )
        self._error_label.pack(pady=(0, PAD_SM))

        # Buttons
        btn_frame = ctk.CTkFrame(frame, fg_color="transparent")
        btn_frame.pack(fill="x")

        ctk.CTkButton(
            btn_frame,
            text="Exit",
            width=80,
            height=36,
            font=FONTS["body"],
            fg_color="transparent",
            text_color=theme["text_secondary"],
            hover_color=theme["bg_hover"],
            command=self._exit,
        ).pack(side="left")

        ctk.CTkButton(
            btn_frame,
            text="Activate",
            width=120,
            height=36,
            font=FONTS["body"],
            fg_color=theme["accent"],
            hover_color=theme["accent_hover"],
            command=self._activate,
        ).pack(side="right")

    def _activate(self):
        key = self._key_var.get().strip()
        if not key:
            self._error_label.configure(text="Please enter a license key")
            return

        if activate_license(key):
            self.result = True
            self.destroy()
        else:
            self._error_label.configure(text="Invalid license key for this machine")

    def _exit(self):
        self.result = False
        self.destroy()
