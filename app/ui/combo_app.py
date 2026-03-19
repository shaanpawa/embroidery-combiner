"""
Excel-driven combo workflow app.
Single-screen layout: Excel upload, DST folder, combo preview, export.
"""

import os
import platform
import subprocess
import threading
from datetime import date

import customtkinter as ctk

from app.config import (
    Config, APP_NAME, APP_VERSION, DEFAULT_COLUMN_GAP_MM, DEFAULT_GAP_MM,
)
from app.core.excel_parser import (
    ComboFile, ComboGroup, generate_all_combos, group_entries, parse_excel,
)
from app.core.pipeline import check_combo_ready, export_all
from app.ui.theme import (
    COLORS, FONTS, PAD_SM, PAD_MD, PAD_LG, PAD_XL, PAD_XS,
    RADIUS_MD, RADIUS_SM,
)
from app.ui.components.combo_list import ComboList


class ComboApp(ctk.CTk):
    """Main application window for Excel-driven combo workflow."""

    def __init__(self, config: Config):
        super().__init__()
        self.config = config
        self.title(f"{APP_NAME} v{APP_VERSION}")
        self.geometry(config.window_geometry or "900x750")
        self.minsize(800, 600)

        ctk.set_appearance_mode(config.theme)

        self._entries = []
        self._groups = []
        self._combos = []
        self._combo_files_by_group = {}
        self._dst_folder = ""
        self._is_processing = False

        self._build_layout()

        # Restore previous paths
        if config.last_excel_path and os.path.isfile(config.last_excel_path):
            self._excel_var.set(config.last_excel_path)
            self.after(100, lambda: self._load_excel(config.last_excel_path))

        if config.last_dst_folder and os.path.isdir(config.last_dst_folder):
            self._dst_var.set(config.last_dst_folder)
            self._dst_folder = config.last_dst_folder

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
            header, text=APP_NAME,
            font=FONTS["heading"], text_color=theme["text_primary"],
        ).pack(side="left")

        ctk.CTkLabel(
            header, text=f"v{APP_VERSION}",
            font=FONTS["tiny"], text_color=theme["text_muted"],
        ).pack(side="left", padx=(PAD_SM, 0), pady=(4, 0))

        # ── Session name ──
        session_frame = ctk.CTkFrame(main, fg_color="transparent")
        session_frame.pack(fill="x", pady=(0, PAD_SM))

        ctk.CTkLabel(
            session_frame, text="Session",
            font=FONTS["subheading"], text_color=theme["text_primary"],
        ).pack(side="left", padx=(0, PAD_SM))

        self._session_var = ctk.StringVar(value=date.today().strftime("%d %b %Y"))
        ctk.CTkEntry(
            session_frame, textvariable=self._session_var,
            height=36, font=FONTS["body"], width=200,
        ).pack(side="left")

        # ── Excel upload ──
        excel_frame = ctk.CTkFrame(main, fg_color="transparent")
        excel_frame.pack(fill="x", pady=(0, PAD_SM))

        ctk.CTkLabel(
            excel_frame, text="Order Excel",
            font=FONTS["subheading"], text_color=theme["text_primary"],
        ).pack(side="left", padx=(0, PAD_SM))

        self._excel_var = ctk.StringVar()
        ctk.CTkEntry(
            excel_frame, textvariable=self._excel_var,
            state="readonly", height=36, font=FONTS["body"],
        ).pack(side="left", fill="x", expand=True, padx=(0, PAD_SM))

        ctk.CTkButton(
            excel_frame, text="Browse", width=80, height=36,
            font=FONTS["body"], fg_color=theme["bg_surface"],
            text_color=theme["text_primary"], hover_color=theme["bg_hover"],
            command=self._browse_excel,
        ).pack(side="left")

        # ── Excel parse summary ──
        self._excel_summary_var = ctk.StringVar()
        self._excel_summary = ctk.CTkLabel(
            main, textvariable=self._excel_summary_var,
            font=FONTS["small"], text_color=theme["text_secondary"],
            anchor="w",
        )
        self._excel_summary.pack(fill="x", pady=(0, PAD_SM))

        # ── DST folder ──
        dst_frame = ctk.CTkFrame(main, fg_color="transparent")
        dst_frame.pack(fill="x", pady=(0, PAD_SM))

        ctk.CTkLabel(
            dst_frame, text="DST Folder",
            font=FONTS["subheading"], text_color=theme["text_primary"],
        ).pack(side="left", padx=(0, PAD_SM))

        self._dst_var = ctk.StringVar()
        ctk.CTkEntry(
            dst_frame, textvariable=self._dst_var,
            state="readonly", height=36, font=FONTS["body"],
        ).pack(side="left", fill="x", expand=True, padx=(0, PAD_SM))

        ctk.CTkButton(
            dst_frame, text="Browse", width=80, height=36,
            font=FONTS["body"], fg_color=theme["bg_surface"],
            text_color=theme["text_primary"], hover_color=theme["bg_hover"],
            command=self._browse_dst,
        ).pack(side="left")

        # ── DST match summary ──
        self._dst_summary_var = ctk.StringVar()
        ctk.CTkLabel(
            main, textvariable=self._dst_summary_var,
            font=FONTS["small"], text_color=theme["text_secondary"],
            anchor="w",
        ).pack(fill="x", pady=(0, PAD_SM))

        # ── Combo list section ──
        combo_header = ctk.CTkFrame(main, fg_color="transparent")
        combo_header.pack(fill="x", pady=(PAD_SM, PAD_XS))

        ctk.CTkLabel(
            combo_header, text="Combo Files",
            font=FONTS["subheading"], text_color=theme["text_primary"],
        ).pack(side="left")

        self._select_label = ctk.CTkLabel(
            combo_header, text="",
            font=FONTS["small"], text_color=theme["text_secondary"],
        )
        self._select_label.pack(side="left", padx=(PAD_SM, 0))

        ctk.CTkButton(
            combo_header, text="Deselect All", width=90, height=28,
            font=FONTS["tiny"], fg_color="transparent",
            text_color=theme["text_muted"], hover_color=theme["bg_hover"],
            command=self._deselect_all,
        ).pack(side="right")

        ctk.CTkButton(
            combo_header, text="Select All", width=80, height=28,
            font=FONTS["tiny"], fg_color="transparent",
            text_color=theme["text_muted"], hover_color=theme["bg_hover"],
            command=self._select_all,
        ).pack(side="right")

        self._combo_list = ComboList(
            main, on_select_change=self._on_selection_change, height=280,
        )
        self._combo_list.pack(fill="both", expand=True, pady=(0, PAD_SM))

        # ── Export section ──
        export_frame = ctk.CTkFrame(main, fg_color="transparent")
        export_frame.pack(fill="x", pady=(0, PAD_SM))

        ctk.CTkLabel(
            export_frame, text="Output",
            font=FONTS["subheading"], text_color=theme["text_primary"],
        ).pack(side="left", padx=(0, PAD_SM))

        self._output_var = ctk.StringVar()
        ctk.CTkEntry(
            export_frame, textvariable=self._output_var,
            state="readonly", height=36, font=FONTS["body"],
        ).pack(side="left", fill="x", expand=True, padx=(0, PAD_SM))

        ctk.CTkButton(
            export_frame, text="Browse", width=80, height=36,
            font=FONTS["body"], fg_color=theme["bg_surface"],
            text_color=theme["text_primary"], hover_color=theme["bg_hover"],
            command=self._browse_output,
        ).pack(side="left")

        # ── Export button ──
        self._export_btn = ctk.CTkButton(
            main, text="Export Selected", height=48,
            font=(FONTS["subheading"][0], 15, "bold"),
            fg_color=theme["accent"], hover_color=theme["accent_hover"],
            corner_radius=RADIUS_MD, command=self._start_export,
        )
        self._export_btn.pack(fill="x", pady=(PAD_SM, 0))
        self._export_btn.configure(state="disabled")

        # ── Progress bar ──
        self._progress_var = ctk.DoubleVar(value=0)
        self._progress_bar = ctk.CTkProgressBar(
            main, variable=self._progress_var,
            height=6, corner_radius=3,
        )
        # Hidden initially

        # ── Status bar ──
        self._status_var = ctk.StringVar(value="Upload an Excel order file to get started")
        ctk.CTkLabel(
            main, textvariable=self._status_var,
            font=FONTS["tiny"], text_color=theme["text_muted"],
            anchor="w",
        ).pack(fill="x", pady=(PAD_SM, 0))

    # ── Browse handlers ──

    def _browse_excel(self):
        path = ctk.filedialog.askopenfilename(
            title="Select order Excel file",
            initialdir=os.path.dirname(self.config.last_excel_path) or None,
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if path:
            self._excel_var.set(path)
            self._load_excel(path)

    def _browse_dst(self):
        folder = ctk.filedialog.askdirectory(
            title="Select folder with DST files",
            initialdir=self.config.last_dst_folder or None,
        )
        if folder:
            self._dst_var.set(folder)
            self._dst_folder = folder
            self.config.last_dst_folder = folder
            self.config.save()
            self._check_dst_matches()

    def _browse_output(self):
        folder = ctk.filedialog.askdirectory(
            title="Select output folder",
            initialdir=self.config.last_output_folder or None,
        )
        if folder:
            self._output_var.set(folder)
            self.config.last_output_folder = folder
            self.config.save()
            self._update_export_state()

    # ── Excel loading ──

    def _load_excel(self, path: str):
        self.config.last_excel_path = path
        self.config.save()

        try:
            result = parse_excel(path)
        except Exception as e:
            self._excel_summary_var.set(f"Error reading Excel: {e}")
            return

        self._entries = result.entries

        if not self._entries:
            self._excel_summary_var.set("No valid entries found in Excel")
            return

        # Generate combos
        self._groups = group_entries(self._entries)
        self._combos = generate_all_combos(self._entries)

        # Build lookup for combo list
        self._combo_files_by_group = {}
        for combo in self._combos:
            key = (combo.machine_program, combo.com_no)
            self._combo_files_by_group.setdefault(key, []).append(combo)

        # Populate UI
        self._combo_list.populate(self._groups, self._combo_files_by_group)

        # Summary
        total_slots = sum(e.quantity for e in self._entries)
        warnings_text = ""
        if result.warnings:
            warnings_text = f"  ({len(result.warnings)} warning{'s' if len(result.warnings) != 1 else ''})"
        self._excel_summary_var.set(
            f"{len(self._entries)} names, {len(self._groups)} groups, "
            f"{len(self._combos)} combo files, {total_slots} total slots{warnings_text}"
        )

        self._on_selection_change()
        self._check_dst_matches()

    # ── DST file matching ──

    def _check_dst_matches(self):
        if not self._entries or not self._dst_folder:
            self._dst_summary_var.set("")
            return

        all_programs = set()
        for entry in self._entries:
            all_programs.add(entry.program)

        found = 0
        missing = []
        for prog in sorted(all_programs):
            path = os.path.join(self._dst_folder, f"{prog}.dst")
            if os.path.isfile(path):
                found += 1
            else:
                missing.append(prog)

        total = len(all_programs)
        if missing:
            show_missing = missing[:10]
            more = f" +{len(missing)-10} more" if len(missing) > 10 else ""
            self._dst_summary_var.set(
                f"{found}/{total} DST files found. "
                f"Missing: {', '.join(str(p) for p in show_missing)}{more}"
            )
        else:
            self._dst_summary_var.set(f"All {total} DST files found")

        self._update_export_state()

    # ── Selection ──

    def _select_all(self):
        self._combo_list.select_all()

    def _deselect_all(self):
        self._combo_list.deselect_all()

    def _on_selection_change(self):
        selected = self._combo_list.selected_count
        total = self._combo_list.total_combos
        self._select_label.configure(text=f"{selected}/{total} selected")
        self._update_export_state()

    def _update_export_state(self):
        selected = self._combo_list.get_selected_combos()
        has_output = bool(self._output_var.get())
        has_dst = bool(self._dst_folder)

        if selected and has_output and has_dst and not self._is_processing:
            self._export_btn.configure(state="normal")
            self._status_var.set(f"Ready to export {len(selected)} combo file(s)")
        else:
            self._export_btn.configure(state="disabled")
            if not selected:
                self._status_var.set("Select combo files to export")
            elif not has_dst:
                self._status_var.set("Select a DST folder")
            elif not has_output:
                self._status_var.set("Select an output folder")

    # ── Export ──

    def _start_export(self):
        if self._is_processing:
            return

        selected = self._combo_list.get_selected_combos()
        if not selected:
            return

        self._is_processing = True
        self._export_btn.configure(state="disabled")
        self._progress_bar.pack(fill="x", pady=(PAD_XS, 0))
        self._progress_var.set(0)

        threading.Thread(
            target=self._run_export,
            args=(selected,),
            daemon=True,
        ).start()

    def _run_export(self, combos):
        try:
            output_folder = self._output_var.get()
            gap = self.config.gap_mm
            col_gap = self.config.column_gap_mm

            def on_progress(current, total):
                self.after(0, lambda: self._progress_var.set(current / total))
                self.after(0, lambda: self._status_var.set(
                    f"Exporting {current}/{total}..."
                ))

            results = export_all(
                combos, self._dst_folder, output_folder,
                gap_mm=gap, column_gap_mm=col_gap,
                overwrite=True,
                progress_callback=on_progress,
            )

            self.after(0, lambda: self._export_done(results))

        except Exception as e:
            self.after(0, lambda: self._export_error(str(e)))

    def _export_done(self, results):
        self._is_processing = False
        self._progress_bar.pack_forget()
        self._export_btn.configure(state="normal")

        success = [r for r in results if r.success]
        failed = [r for r in results if not r.success]

        if failed:
            fail_names = [r.combo.filename for r in failed[:5]]
            more = f" +{len(failed)-5} more" if len(failed) > 5 else ""
            self._status_var.set(
                f"Exported {len(success)}, failed {len(failed)}: "
                f"{', '.join(fail_names)}{more}"
            )
        else:
            self._status_var.set(
                f"Exported {len(success)} combo file(s) successfully"
            )

        # Show success dialog
        if success:
            self._show_success_dialog(success, failed)

    def _export_error(self, message):
        self._is_processing = False
        self._progress_bar.pack_forget()
        self._export_btn.configure(state="normal")
        self._status_var.set(f"Error: {message}")

    def _show_success_dialog(self, success, failed):
        theme = COLORS.get(ctk.get_appearance_mode().lower(), COLORS["dark"])
        dialog = ctk.CTkToplevel(self)
        dialog.title("Export Complete")
        dialog.geometry("400x200")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()
        dialog.configure(fg_color=theme["bg_primary"])

        frame = ctk.CTkFrame(dialog, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=PAD_LG, pady=PAD_LG)

        if failed:
            ctk.CTkLabel(
                frame,
                text=f"Exported {len(success)} of {len(success)+len(failed)} files",
                font=FONTS["subheading"], text_color=theme["warning"],
            ).pack(pady=(0, PAD_SM))
        else:
            ctk.CTkLabel(
                frame,
                text=f"Exported {len(success)} file(s) successfully!",
                font=FONTS["subheading"], text_color=theme["success"],
            ).pack(pady=(0, PAD_SM))

        output_folder = self._output_var.get()
        ctk.CTkLabel(
            frame,
            text=output_folder,
            font=FONTS["tiny"], text_color=theme["text_secondary"],
        ).pack(pady=(0, PAD_MD))

        btn_frame = ctk.CTkFrame(frame, fg_color="transparent")
        btn_frame.pack()

        ctk.CTkButton(
            btn_frame, text="Open Folder", width=110, height=36,
            fg_color=theme["bg_surface"], hover_color=theme["bg_hover"],
            text_color=theme["text_primary"],
            command=lambda: (self._open_folder(output_folder), dialog.destroy()),
        ).pack(side="left", padx=(0, PAD_SM))

        ctk.CTkButton(
            btn_frame, text="OK", width=100, height=36,
            fg_color=theme["accent"], hover_color=theme["accent_hover"],
            command=dialog.destroy,
        ).pack(side="left")

    @staticmethod
    def _open_folder(folder_path: str):
        try:
            if platform.system() == "Windows":
                os.startfile(folder_path)
            elif platform.system() == "Darwin":
                subprocess.Popen(["open", folder_path])
            else:
                subprocess.Popen(["xdg-open", folder_path])
        except Exception:
            pass

    def _on_close(self):
        self.config.window_geometry = self.geometry()
        self.config.save()
        self.destroy()
