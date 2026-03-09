"""
Main application window.
Orchestrates discovery, validation, conversion, and combining.
"""

import os
import tempfile
import threading

import customtkinter as ctk

from app.config import Config, APP_NAME
from app.core.file_discovery import (
    DiscoveryResult, discover_folder, generate_output_name,
)
from app.core.validator import validate_batch, is_valid_for_combining
from app.core.combiner import (
    CombineError, combine_designs, save_combined, validate_combined_output,
)
from app.core.converter import (
    check_conversion_capability, batch_convert, cleanup_temp_files,
)
from app.ui.theme import COLORS, FONTS, PAD_SM, PAD_MD, PAD_LG, PAD_XL, RADIUS_MD
from app.ui.components.file_table import FileTable
from app.ui.components.gap_controls import GapControls
from app.ui.components.output_panel import OutputPanel
from app.ui.components.progress_panel import ProgressPanel
from app.ui.components.alerts import AlertsBanner


class CombinerApp(ctk.CTk):
    """Main application window."""

    def __init__(self, config: Config):
        super().__init__()
        self.config = config
        self.title(APP_NAME)
        self.geometry(config.window_geometry or "720x680")
        self.minsize(640, 560)

        ctk.set_appearance_mode(config.theme)

        self._discovery: DiscoveryResult | None = None
        self._validation_results = {}
        self._is_processing = False

        self._build_layout()

        # Restore last folder if it still exists
        if config.last_folder and os.path.isdir(config.last_folder):
            self._folder_var.set(config.last_folder)
            self.after(100, lambda: self._on_folder_selected(config.last_folder))

        # Save geometry on close
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_layout(self):
        theme = COLORS.get(ctk.get_appearance_mode().lower(), COLORS["dark"])
        self.configure(fg_color=theme["bg_primary"])

        main = ctk.CTkFrame(self, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=PAD_XL, pady=PAD_LG)

        # ── Header ──
        header = ctk.CTkFrame(main, fg_color="transparent")
        header.pack(fill="x", pady=(0, PAD_LG))

        ctk.CTkLabel(
            header,
            text=APP_NAME,
            font=FONTS["heading"],
            text_color=theme["text_primary"],
        ).pack(side="left")

        # Theme toggle
        self._theme_btn = ctk.CTkButton(
            header,
            text="Light" if self.config.theme == "dark" else "Dark",
            width=60,
            height=30,
            font=FONTS["small"],
            fg_color="transparent",
            text_color=theme["text_muted"],
            hover_color=theme["bg_hover"],
            command=self._toggle_theme,
        )
        self._theme_btn.pack(side="right")

        # ── Folder selector ──
        folder_frame = ctk.CTkFrame(main, fg_color="transparent")
        folder_frame.pack(fill="x", pady=(0, PAD_SM))

        ctk.CTkLabel(
            folder_frame,
            text="Folder",
            font=FONTS["subheading"],
            text_color=theme["text_primary"],
        ).pack(side="left", padx=(0, PAD_SM))

        self._folder_var = ctk.StringVar()
        ctk.CTkEntry(
            folder_frame,
            textvariable=self._folder_var,
            state="readonly",
            height=36,
            font=FONTS["body"],
        ).pack(side="left", fill="x", expand=True, padx=(0, PAD_SM))

        ctk.CTkButton(
            folder_frame,
            text="Browse",
            width=80,
            height=36,
            font=FONTS["body"],
            fg_color=theme["bg_surface"],
            text_color=theme["text_primary"],
            hover_color=theme["bg_hover"],
            command=self._browse_folder,
        ).pack(side="left")

        # ── Alerts banner ──
        self._alerts = AlertsBanner(main)
        self._alerts.pack(fill="x", pady=(PAD_SM, 0))

        # ── File table ──
        self._file_table = FileTable(
            main,
            on_toggle=self._on_file_toggled,
            height=280,
        )
        self._file_table.pack(fill="both", expand=True, pady=(PAD_SM, PAD_SM))

        # ── Controls row ──
        controls = ctk.CTkFrame(main, fg_color="transparent")
        controls.pack(fill="x", pady=(0, PAD_SM))

        self._gap_controls = GapControls(
            controls,
            initial_gap=self.config.gap_mm,
            on_change=self._on_gap_changed,
        )
        self._gap_controls.pack(side="left")

        # ── Output row ──
        self._output_panel = OutputPanel(main)
        self._output_panel.pack(fill="x", pady=(0, PAD_SM))

        # ── Progress ──
        self._progress = ProgressPanel(main)
        # Starts hidden, shown during processing

        # ── Combine button ──
        self._combine_btn = ctk.CTkButton(
            main,
            text="COMBINE",
            height=48,
            font=(FONTS["subheading"][0], 15, "bold"),
            fg_color=theme["accent"],
            hover_color=theme["accent_hover"],
            corner_radius=RADIUS_MD,
            command=self._start_pipeline,
        )
        self._combine_btn.pack(fill="x", pady=(PAD_SM, 0))
        self._combine_btn.configure(state="disabled")

        # ── Status bar ──
        self._status_var = ctk.StringVar(value="Select a folder to begin")
        ctk.CTkLabel(
            main,
            textvariable=self._status_var,
            font=FONTS["tiny"],
            text_color=theme["text_muted"],
            anchor="w",
        ).pack(fill="x", pady=(PAD_SM, 0))

    # ── Folder selection ──

    def _browse_folder(self):
        folder = ctk.filedialog.askdirectory(
            title="Select folder with embroidery files",
            initialdir=self.config.last_folder or None,
        )
        if folder:
            self._folder_var.set(folder)
            self._on_folder_selected(folder)

    def _on_folder_selected(self, folder: str):
        self.config.last_folder = folder
        self.config.save()

        # Run discovery
        self._discovery = discover_folder(folder)
        result = self._discovery

        if not result.files:
            msg = "No embroidery files found"
            if result.skipped_files:
                msg += f" ({len(result.skipped_files)} non-embroidery files ignored)"
            self._status_var.set(msg)
            self._file_table.clear()
            self._alerts.clear()
            self._combine_btn.configure(state="disabled")
            return

        # Populate file table
        self._file_table.populate(result.files)

        # Show alerts
        alerts = []
        for w in result.warnings:
            if "Missing" in w or "Duplicate" in w:
                alerts.append(("warning", w))
            elif "convert" in w.lower():
                alerts.append(("info", w))
            else:
                alerts.append(("info", w))
        self._alerts.set_alerts(alerts)

        # Update output panel
        self._output_panel.set_output_dir(folder)
        self._output_panel.set_auto_name(result.files)

        # Update status
        self._status_var.set(
            f"{result.total_files} files found "
            f"({result.dst_count} DST, {result.ngs_count} NGS)"
        )

        # Run validation in background
        threading.Thread(target=self._run_validation, daemon=True).start()

    # ── Validation ──

    def _run_validation(self):
        if not self._discovery:
            return

        paths = [f.path for f in self._discovery.files]

        def on_progress(current, total, result):
            idx = current - 1
            if result.errors:
                level = "error"
                detail = result.summary
                # Auto-exclude errored files
                if idx < len(self._discovery.files):
                    self._discovery.files[idx].included = False
            elif result.warnings:
                level = "warning"
                detail = result.summary
            else:
                level = "ok"
                detail = result.summary

            self.after(0, lambda: self._file_table.update_status(
                idx, result.status, detail, level
            ))

        results = validate_batch(paths, progress_callback=on_progress)
        self._validation_results = {r.path: r for r in results}

        # Update combine button based on valid files
        valid_count = sum(1 for r in results if is_valid_for_combining(r))
        self.after(0, lambda: self._post_validation(valid_count))

    def _post_validation(self, valid_count: int):
        if valid_count >= 1:
            self._combine_btn.configure(state="normal")
            self._status_var.set(
                f"{valid_count} valid file(s) ready to combine"
            )
        else:
            self._combine_btn.configure(state="disabled")
            self._status_var.set("No valid files to combine")

        # Update output name based on included files
        self._on_file_toggled()

    # ── File toggle ──

    def _on_file_toggled(self):
        if not self._discovery:
            return
        included = self._file_table.get_included_files()
        self._output_panel.set_auto_name(included)

        # Disable combine if no included files
        if not included:
            self._combine_btn.configure(state="disabled")
        else:
            self._combine_btn.configure(state="normal")

    # ── Gap change ──

    def _on_gap_changed(self, value: float):
        self.config.gap_mm = value
        self.config.save()

    # ── Pipeline ──

    def _start_pipeline(self):
        if self._is_processing:
            return

        included = self._file_table.get_included_files()
        if not included:
            return

        # Check for overwrite
        if self._output_panel.check_overwrite():
            dialog = ctk.CTkInputDialog(
                text="Output file already exists. Type 'overwrite' to confirm:",
                title="Confirm Overwrite",
            )
            response = dialog.get_input()
            if response != "overwrite":
                return

        self._is_processing = True
        self._combine_btn.configure(state="disabled")
        self._progress.show()

        threading.Thread(
            target=self._run_pipeline,
            args=(included,),
            daemon=True,
        ).start()

    def _run_pipeline(self, files):
        try:
            gap = self._gap_controls.get_gap_mm()
            output_path = self._output_panel.get_output_path()
            overwrite = os.path.exists(output_path)

            # Separate NGS and DST files
            ngs_files = [f for f in files if f.extension == '.ngs']
            dst_files = [f for f in files if f.extension == '.dst']
            dst_paths = [f.path for f in dst_files]

            temp_dir = None
            conversion_results = None

            # ── Phase 1: Convert NGS → DST ──
            if ngs_files:
                capable, msg = check_conversion_capability()
                if not capable:
                    self.after(0, lambda: self._pipeline_error(msg))
                    return

                self.after(0, lambda: self._progress.set_phase(
                    f"Converting {len(ngs_files)} NGS files..."
                ))

                temp_dir = tempfile.mkdtemp(prefix="embroidery_")
                ngs_paths = [f.path for f in ngs_files]

                def on_convert(current, total, result):
                    # Find the index in the original file list
                    for i, f in enumerate(self._discovery.files):
                        if f.path == result.ngs_path:
                            if result.success:
                                self.after(0, lambda idx=i: self._file_table.update_status(
                                    idx, "Converted", "OK", "done"
                                ))
                            else:
                                self.after(0, lambda idx=i, e=result.error: self._file_table.update_status(
                                    idx, "Failed", e or "Unknown error", "error"
                                ))
                            break
                    self.after(0, lambda: self._progress.set_progress(current, total))

                conversion_results = batch_convert(
                    ngs_paths, temp_dir,
                    progress_callback=on_convert,
                )

                # Collect successful conversions
                for cr in conversion_results:
                    if cr.success and cr.dst_path:
                        dst_paths.append(cr.dst_path)

                failed = [cr for cr in conversion_results if not cr.success]
                if failed:
                    names = [os.path.basename(cr.ngs_path) for cr in failed]
                    self.after(0, lambda: self._status_var.set(
                        f"Warning: {len(failed)} file(s) failed conversion: {', '.join(names)}"
                    ))

            if not dst_paths:
                self.after(0, lambda: self._pipeline_error(
                    "No valid files to combine after conversion"
                ))
                return

            # Sort DST paths by extracted number
            from app.core.file_discovery import extract_number
            dst_paths.sort(key=lambda p: extract_number(os.path.basename(p)) or 0)

            # ── Phase 2: Combine ──
            self.after(0, lambda: self._progress.set_phase(
                f"Combining {len(dst_paths)} designs..."
            ))

            def on_combine(current, total):
                self.after(0, lambda: self._progress.set_progress(current, total))

            combined = combine_designs(dst_paths, gap_mm=gap, progress_callback=on_combine)

            # ── Phase 3: Save ──
            self.after(0, lambda: self._progress.set_phase("Saving..."))
            save_combined(combined, output_path, overwrite=overwrite)

            # ── Phase 4: Verify ──
            info = validate_combined_output(output_path)

            # Cleanup temp files
            if temp_dir and conversion_results:
                cleanup_temp_files(conversion_results, temp_dir)

            # Success
            self.after(0, lambda: self._pipeline_success(output_path, info))

        except (CombineError, FileExistsError, Exception) as e:
            self.after(0, lambda: self._pipeline_error(str(e)))

    def _pipeline_success(self, output_path: str, info: dict):
        self._is_processing = False
        self._progress.hide()

        if info.get("valid"):
            self._status_var.set(
                f"Done — {info['stitch_count']} stitches, "
                f"{info['width_mm']:.1f} x {info['height_mm']:.1f} mm → "
                f"{os.path.basename(output_path)}"
            )
        else:
            self._status_var.set(f"Saved: {os.path.basename(output_path)}")

        # Mark all included files as done
        if self._discovery:
            for i, f in enumerate(self._discovery.files):
                if f.included:
                    self._file_table.update_status(i, "Done", "Combined", "done")

        self._combine_btn.configure(state="normal")

        # Show success dialog
        theme = COLORS.get(ctk.get_appearance_mode().lower(), COLORS["dark"])
        dialog = ctk.CTkToplevel(self)
        dialog.title("Success")
        dialog.geometry("380x180")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()
        dialog.configure(fg_color=theme["bg_primary"])

        frame = ctk.CTkFrame(dialog, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=PAD_LG, pady=PAD_LG)

        ctk.CTkLabel(
            frame,
            text="Combined successfully!",
            font=FONTS["subheading"],
            text_color=theme["success"],
        ).pack(pady=(0, PAD_SM))

        ctk.CTkLabel(
            frame,
            text=os.path.basename(output_path),
            font=FONTS["mono"],
            text_color=theme["text_primary"],
        ).pack(pady=(0, PAD_SM))

        if info.get("valid"):
            ctk.CTkLabel(
                frame,
                text=f"{info['stitch_count']} stitches  |  {info['width_mm']:.1f} x {info['height_mm']:.1f} mm",
                font=FONTS["small"],
                text_color=theme["text_secondary"],
            ).pack(pady=(0, PAD_MD))

        ctk.CTkButton(
            frame,
            text="OK",
            width=100,
            height=36,
            fg_color=theme["accent"],
            hover_color=theme["accent_hover"],
            command=dialog.destroy,
        ).pack()

    def _pipeline_error(self, message: str):
        self._is_processing = False
        self._progress.hide()
        self._combine_btn.configure(state="normal")
        self._status_var.set(f"Error: {message}")

        theme = COLORS.get(ctk.get_appearance_mode().lower(), COLORS["dark"])
        dialog = ctk.CTkToplevel(self)
        dialog.title("Error")
        dialog.geometry("400x200")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()
        dialog.configure(fg_color=theme["bg_primary"])

        frame = ctk.CTkFrame(dialog, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=PAD_LG, pady=PAD_LG)

        ctk.CTkLabel(
            frame,
            text="Something went wrong",
            font=FONTS["subheading"],
            text_color=theme["error"],
        ).pack(pady=(0, PAD_SM))

        ctk.CTkLabel(
            frame,
            text=message,
            font=FONTS["body"],
            text_color=theme["text_secondary"],
            wraplength=340,
        ).pack(pady=(0, PAD_MD), fill="x")

        ctk.CTkButton(
            frame,
            text="OK",
            width=100,
            height=36,
            fg_color=theme["bg_surface"],
            hover_color=theme["bg_hover"],
            text_color=theme["text_primary"],
            command=dialog.destroy,
        ).pack()

    # ── Theme toggle ──

    def _toggle_theme(self):
        new_theme = "light" if self.config.theme == "dark" else "dark"
        self.config.theme = new_theme
        self.config.save()
        ctk.set_appearance_mode(new_theme)
        self._theme_btn.configure(
            text="Light" if new_theme == "dark" else "Dark"
        )

    # ── Window close ──

    def _on_close(self):
        self.config.window_geometry = self.geometry()
        self.config.save()
        self.destroy()
