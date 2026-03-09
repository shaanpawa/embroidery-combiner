"""Build script for creating standalone .exe with Nuitka."""

import subprocess
import sys


def build():
    cmd = [
        sys.executable, "-m", "nuitka",
        "--standalone",
        "--onefile",
        "--windows-console-mode=disable",
        "--enable-plugin=tk-inter",
        "--include-package=customtkinter",
        "--include-data-dir=customtkinter=customtkinter",
        "--output-filename=EmbroideryC.exe",
        "--windows-icon-from-ico=icon.ico",
        "--company-name=FM Embroidery",
        "--product-name=Embroidery Combiner",
        "--file-version=1.0.0",
        "--product-version=1.0.0",
        "main.py",
    ]
    print("Building with Nuitka...")
    print(" ".join(cmd))
    subprocess.run(cmd, check=True)
    print("Build complete!")


if __name__ == "__main__":
    build()
