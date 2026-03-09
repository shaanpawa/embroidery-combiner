"""
NGS to DST converter via GUI automation (pywinauto).
Supports Wings XP (paid) and My Editor (free) — auto-detects whichever is installed.

This module only works on Windows.
On other platforms, check_conversion_capability() returns False
and the app instructs the user to convert files manually.
"""

import os
import sys
import tempfile
import time
from dataclasses import dataclass
from typing import Callable, List, Optional, Tuple


class ConversionError(Exception):
    """Raised when file conversion fails."""
    pass


@dataclass
class ConversionResult:
    ngs_path: str
    dst_path: Optional[str]
    success: bool
    error: Optional[str] = None


def is_windows() -> bool:
    return sys.platform == 'win32'


@dataclass
class EditorInfo:
    """Info about a detected Wings editor."""
    name: str           # "Wings XP" or "My Editor"
    exe_path: str       # Full path to executable
    process_name: str   # e.g., "WingsXP.exe" or "MyEditor.exe"
    window_title_re: str  # Regex to match the main window title


# Search order: Wings XP first (more capable), then My Editor (free)
_EDITOR_SEARCH = [
    {
        "name": "Wings XP",
        "process_name": "WingsXP.exe",
        "window_title_re": ".*Wings.*",
        "paths": [
            r"C:\Program Files\Wings Systems\Wings XP\WingsXP.exe",
            r"C:\Program Files (x86)\Wings Systems\Wings XP\WingsXP.exe",
            r"C:\Wings\WingsXP.exe",
            r"C:\Program Files\Wings\WingsXP.exe",
            r"C:\Program Files (x86)\Wings\WingsXP.exe",
        ],
        "which": ["WingsXP", "WingsXP.exe"],
    },
    {
        "name": "My Editor",
        "process_name": "MyEditor.exe",
        "window_title_re": ".*My Editor.*",
        "paths": [
            r"C:\Program Files\Wings Systems\My Editor\MyEditor.exe",
            r"C:\Program Files (x86)\Wings Systems\My Editor\MyEditor.exe",
            r"C:\Wings\My Editor\MyEditor.exe",
        ],
        "which": ["MyEditor", "MyEditor.exe"],
    },
]


def find_editor() -> Optional[EditorInfo]:
    """Locate Wings XP or My Editor on the system. Returns the first one found."""
    if not is_windows():
        return None

    import shutil as sh

    for entry in _EDITOR_SEARCH:
        # Check common install paths
        for path in entry["paths"]:
            if os.path.exists(path):
                return EditorInfo(
                    name=entry["name"],
                    exe_path=path,
                    process_name=entry["process_name"],
                    window_title_re=entry["window_title_re"],
                )
        # Check PATH
        for cmd in entry["which"]:
            found = sh.which(cmd)
            if found:
                return EditorInfo(
                    name=entry["name"],
                    exe_path=found,
                    process_name=entry["process_name"],
                    window_title_re=entry["window_title_re"],
                )

    return None


# Keep old name as alias for backward compatibility
def find_my_editor() -> Optional[str]:
    """Locate any Wings editor. Returns the exe path or None."""
    info = find_editor()
    return info.exe_path if info else None


def check_conversion_capability() -> Tuple[bool, str]:
    """
    Check if this system can convert NGS files.

    Returns:
        (capable, message) — capable is True if conversion is possible.
    """
    if not is_windows():
        return False, (
            "NGS conversion requires Windows with Wings XP or My Editor.\n\n"
            "On this system, you can only combine .dst files.\n"
            "To use .ngs files:\n"
            "1. Open each .ngs file in Wings on a Windows machine\n"
            "2. Save As .dst format\n"
            "3. Place the .dst files in a folder\n"
            "4. Use this tool to combine them"
        )

    editor = find_editor()
    if editor is None:
        return False, (
            "Wings XP or My Editor not found.\n\n"
            "Install Wings My Editor (free) from:\n"
            "https://www.wingssystems.com/index.php/products/myeditor/\n\n"
            "After installing, restart this application."
        )

    return True, f"{editor.name} found: {editor.exe_path}"


def _ensure_editor_closed(editor: Optional[EditorInfo] = None) -> None:
    """Kill any running editor instances before starting conversion."""
    if not is_windows():
        return
    import subprocess

    names_to_kill = []
    if editor:
        names_to_kill.append(editor.process_name)
    else:
        names_to_kill.extend(e["process_name"] for e in _EDITOR_SEARCH)

    for name in names_to_kill:
        subprocess.run(
            ["taskkill", "/f", "/im", name],
            capture_output=True,
        )
    time.sleep(1)


# Backward compat alias
_ensure_my_editor_closed = _ensure_editor_closed


def convert_ngs_to_dst(
    ngs_path: str,
    dst_path: str,
    editor: Optional[EditorInfo] = None,
    timeout: int = 30,
) -> None:
    """
    Convert a single .ngs file to .dst using Wings XP or My Editor GUI automation.

    Raises:
        ConversionError: If conversion fails.
    """
    if not is_windows():
        raise ConversionError("NGS conversion requires Windows")

    if editor is None:
        editor = find_editor()
    if editor is None:
        raise ConversionError("Wings XP or My Editor not found")

    try:
        from pywinauto import Application
    except ImportError:
        raise ConversionError("pywinauto is required — install with: pip install pywinauto")

    ngs_path = os.path.abspath(ngs_path)
    dst_path = os.path.abspath(dst_path)

    app = None
    try:
        app = Application(backend='uia').start(editor.exe_path, timeout=10)
        main_window = app.window(title_re=editor.window_title_re)
        main_window.wait('ready', timeout=timeout)

        # File > Open
        main_window.menu_select("File->Open")
        time.sleep(1)

        open_dialog = main_window.child_window(title="Open", control_type="Window")
        open_dialog.wait('ready', timeout=10)

        filename_edit = open_dialog.child_window(title="File name:", control_type="Edit")
        filename_edit.set_text(ngs_path)
        time.sleep(0.5)

        open_button = open_dialog.child_window(title="Open", control_type="Button")
        open_button.click()
        time.sleep(2)

        # File > Save As
        main_window.menu_select("File->Save As")
        time.sleep(1)

        save_dialog = main_window.child_window(title="Save As", control_type="Window")
        save_dialog.wait('ready', timeout=10)

        # Set file type to DST
        type_combo = save_dialog.child_window(title="Save as type:", control_type="ComboBox")
        type_combo.select("Tajima (*.dst)")
        time.sleep(0.5)

        # Set output filename
        fname_edit = save_dialog.child_window(title="File name:", control_type="Edit")
        fname_edit.set_text(dst_path)
        time.sleep(0.5)

        # Click Save
        save_button = save_dialog.child_window(title="Save", control_type="Button")
        save_button.click()
        time.sleep(1)

        # Handle possible overwrite confirmation dialog
        try:
            confirm = main_window.child_window(title_re=".*Confirm.*|.*Replace.*", control_type="Window")
            confirm.wait('exists', timeout=2)
            yes_btn = confirm.child_window(title_re="Yes|OK", control_type="Button")
            yes_btn.click()
            time.sleep(0.5)
        except Exception:
            pass  # No overwrite dialog appeared

        # Close the file
        try:
            main_window.menu_select("File->Close")
            time.sleep(0.5)
        except Exception:
            pass

    except Exception as e:
        if "timeout" in str(e).lower() or "timed out" in str(e).lower():
            raise ConversionError(
                f"{editor.name} timed out converting {os.path.basename(ngs_path)}"
            )
        raise ConversionError(
            f"Conversion failed for {os.path.basename(ngs_path)}: {e}"
        )
    finally:
        _close_editor(app)

    # Verify output
    if not os.path.exists(dst_path):
        raise ConversionError(
            f"Conversion produced no output for {os.path.basename(ngs_path)}"
        )
    if os.path.getsize(dst_path) == 0:
        raise ConversionError(
            f"Conversion produced empty file for {os.path.basename(ngs_path)}"
        )


def _close_editor(app) -> None:
    """Try to close the editor gracefully, then force-kill if needed."""
    if app is None:
        return
    try:
        app.kill()
    except Exception:
        pass


def batch_convert(
    ngs_files: List[str],
    output_dir: Optional[str] = None,
    progress_callback: Optional[Callable] = None,
    timeout_per_file: int = 30,
) -> List[ConversionResult]:
    """
    Convert multiple .ngs files to .dst with per-file error recovery.

    Args:
        ngs_files: Paths to .ngs files.
        output_dir: Directory for output .dst files. Uses temp dir if None.
        progress_callback: Called as progress_callback(current, total, result).
        timeout_per_file: Seconds before giving up on one file.

    Returns:
        List of ConversionResult — check each for success/failure.
    """
    if output_dir is None:
        output_dir = tempfile.mkdtemp(prefix="embroidery_")
    os.makedirs(output_dir, exist_ok=True)

    editor = find_editor()
    _ensure_editor_closed(editor)

    results = []

    for i, ngs_path in enumerate(ngs_files):
        basename = os.path.splitext(os.path.basename(ngs_path))[0]
        dst_path = os.path.join(output_dir, f"{basename}.dst")

        try:
            convert_ngs_to_dst(ngs_path, dst_path, editor, timeout_per_file)

            # Post-conversion validation
            from app.core.validator import validate_file
            val = validate_file(dst_path)
            if not val.valid:
                results.append(ConversionResult(
                    ngs_path, dst_path, False,
                    f"Converted but invalid: {val.summary}",
                ))
            else:
                results.append(ConversionResult(ngs_path, dst_path, True))

        except ConversionError as e:
            results.append(ConversionResult(ngs_path, None, False, str(e)))
        except Exception as e:
            results.append(ConversionResult(ngs_path, None, False, f"Unexpected: {e}"))

        if progress_callback:
            progress_callback(i + 1, len(ngs_files), results[-1])

    return results


def cleanup_temp_files(results: List[ConversionResult], output_dir: str) -> None:
    """Remove temporary DST files and directory."""
    for r in results:
        if r.dst_path:
            try:
                os.remove(r.dst_path)
            except OSError:
                pass
    try:
        os.rmdir(output_dir)
    except OSError:
        pass
