"""
Embroidery Design Combiner — Entry point.
Combines multiple embroidery design files into one, stacked vertically.
"""

import sys

import customtkinter as ctk

from app.config import Config
from app.licensing import check_license
from app.ui.components.license_dialog import LicenseDialog
from app.ui.app import CombinerApp


def main():
    config = Config()
    config.load()

    # Check license
    if not check_license():
        root = ctk.CTk()
        root.withdraw()
        ctk.set_appearance_mode(config.theme)

        dialog = LicenseDialog(root)
        root.wait_window(dialog)

        if not dialog.result:
            root.destroy()
            sys.exit(0)

        root.destroy()

    # Launch main app
    app = CombinerApp(config)
    app.mainloop()


if __name__ == "__main__":
    main()
