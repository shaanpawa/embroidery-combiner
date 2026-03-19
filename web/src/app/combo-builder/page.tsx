"use client";

import Link from "next/link";
import Image from "next/image";
import { useState, useCallback, useRef, useEffect } from "react";
import { useTheme } from "../theme-provider";

const API = process.env.NEXT_PUBLIC_API_URL || "http://localhost:8000";

interface Slot { program: number; name_line1: string; name_line2: string; quantity: number; }
interface ComboFile { filename: string; part_number: number; total_parts: number; slot_count: number; left_count: number; right_count: number; slots: Slot[]; }
interface Group { machine_program: string; com_no: string; entry_count: number; total_slots: number; combos: ComboFile[]; }
interface EntryPreview { program: number; name_line1: string; name_line2: string; quantity: number; com_no: string; machine_program: string; }
interface ParseResponse { session_id: string; entries_count: number; total_slots: number; groups: Group[]; combo_count: number; warnings: string[]; entries_preview?: EntryPreview[]; }
interface DstResponse { session_id: string; uploaded_count: number; needed_count: number; missing_programs: number[]; all_matched: boolean; }
interface SessionSummary { session_id: string; name: string; created_at: string; updated_at: string; has_excel: boolean; entries_count: number; combo_count: number; dst_count: number; exported: boolean; }
interface FullSession {
  session_id: string; name: string; entries_count: number; total_slots: number;
  groups: Group[]; combo_count: number; entries_preview?: EntryPreview[];
  dst_count: number; dst_all_matched: boolean; missing_programs: number[];
  exported: boolean; exported_at: string | null; excel_filename: string | null;
  gap_mm: number; column_gap_mm: number; warnings: string[];
}

function useCountUp(target: number, duration = 600) {
  const [value, setValue] = useState(0);
  useEffect(() => {
    if (target === 0) { setValue(0); return; }
    const start = performance.now();
    const step = (now: number) => {
      const progress = Math.min((now - start) / duration, 1);
      setValue(Math.round((1 - Math.pow(1 - progress, 3)) * target));
      if (progress < 1) requestAnimationFrame(step);
    };
    requestAnimationFrame(step);
  }, [target, duration]);
  return value;
}

function StatCard({ value, label, description, delay = 0 }: { value: number; label: string; description?: string; delay?: number }) {
  const display = useCountUp(value);
  return (
    <div className="stat-card group relative" style={{ animation: `countUp 0.4s ease ${delay}s forwards`, opacity: 0 }}>
      <div className="stat-number">{display}</div>
      <div className="stat-label">{label}</div>
      {description && (
        <div className="absolute -top-9 left-1/2 -translate-x-1/2 whitespace-nowrap text-[9px] px-2.5 py-1 rounded-md opacity-0 group-hover:opacity-100 transition-opacity pointer-events-none" style={{ background: "var(--foreground)", color: "var(--background)", zIndex: 50 }}>
          {description}
        </div>
      )}
    </div>
  );
}

function Toast({ message, type, onClose }: { message: string; type: "error" | "warning" | "success"; onClose: () => void }) {
  useEffect(() => { const t = setTimeout(onClose, 6000); return () => clearTimeout(t); }, [onClose]);
  const colors = { error: "var(--danger)", warning: "var(--warning)", success: "var(--success)" };
  return (
    <div className="fixed top-4 right-4 z-50 glass-panel px-4 py-3 flex items-center gap-3 max-w-sm" style={{ borderLeft: `3px solid ${colors[type]}`, animation: "fadeSlideIn 0.3s ease" }}>
      <span className="text-xs flex-1">{message}</span>
      <button onClick={onClose} className="text-xs" style={{ color: "var(--muted)" }}>✕</button>
    </div>
  );
}

function formatDate(iso: string): string {
  try {
    const d = new Date(iso);
    return d.toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" });
  } catch { return iso; }
}

export default function ComboBuilder() {
  const { theme, toggle } = useTheme();
  const [sessionStarted, setSessionStarted] = useState(false);
  const [sessionId, setSessionId] = useState("");
  const [sessionName, setSessionName] = useState(
    new Date().toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" })
  );
  const [savedSessions, setSavedSessions] = useState<SessionSummary[]>([]);
  const [showSessionDropdown, setShowSessionDropdown] = useState(false);
  const [deleteConfirm, setDeleteConfirm] = useState<string | null>(null);
  const [loadingSession, setLoadingSession] = useState(false);

  const [excelFile, setExcelFile] = useState("");
  const [parseData, setParseData] = useState<ParseResponse | null>(null);
  const [excelLoading, setExcelLoading] = useState(false);
  const [dstUploaded, setDstUploaded] = useState(false);
  const [dstData, setDstData] = useState<DstResponse | null>(null);
  const [dstLoading, setDstLoading] = useState(false);
  const [dstFileName, setDstFileName] = useState("");
  const [selectedCombos, setSelectedCombos] = useState<Set<string>>(new Set());
  const [expandedGroups, setExpandedGroups] = useState<Set<string>>(new Set());
  const [previewCombo, setPreviewCombo] = useState<ComboFile | null>(null);
  const [exporting, setExporting] = useState(false);
  const [exportProgress, setExportProgress] = useState(0);
  const [downloadUrl, setDownloadUrl] = useState("");
  const [exported, setExported] = useState(false);
  const [excelDragOver, setExcelDragOver] = useState(false);
  const [dstDragOver, setDstDragOver] = useState(false);
  const [toast, setToast] = useState<{ message: string; type: "error" | "warning" | "success" } | null>(null);
  const [showExcelPreview, setShowExcelPreview] = useState(false);
  const [showSettings, setShowSettings] = useState(false);
  const [gapMm, setGapMm] = useState(3);
  const [columnGapMm, setColumnGapMm] = useState(5);
  const excelInputRef = useRef<HTMLInputElement>(null);
  const zipInputRef = useRef<HTMLInputElement>(null);

  const showToast = useCallback((message: string, type: "error" | "warning" | "success" = "error") => setToast({ message, type }), []);

  const fetchSessions = useCallback(async () => {
    try {
      const res = await fetch(`${API}/api/sessions`);
      if (res.ok) setSavedSessions(await res.json());
    } catch { /* ignore */ }
  }, []);

  // Fetch sessions on mount
  useEffect(() => { fetchSessions(); }, [fetchSessions]);

  // Save session name to server when it changes
  const saveSessionName = useCallback(async (sid: string, name: string) => {
    const form = new FormData(); form.append("session_id", sid); form.append("name", name);
    try { await fetch(`${API}/api/session/name`, { method: "POST", body: form }); } catch { /* ignore */ }
  }, []);

  const resetSession = useCallback(() => {
    setDstUploaded(false); setDstData(null); setDstFileName("");
    setSelectedCombos(new Set()); setExpandedGroups(new Set());
    setPreviewCombo(null); setDownloadUrl(""); setExportProgress(0);
    setShowExcelPreview(false); setExported(false);
  }, []);

  const applyParseData = useCallback((data: ParseResponse) => {
    setParseData(data);
    setSessionId(data.session_id);
    const all = new Set<string>();
    data.groups.forEach((g) => g.combos.forEach((c) => all.add(c.filename)));
    setSelectedCombos(all);
    const allGroups = new Set<string>();
    data.groups.forEach((g) => allGroups.add(`${g.machine_program}_${g.com_no}`));
    setExpandedGroups(allGroups);
    if (data.groups[0]?.combos[0]) setPreviewCombo(data.groups[0].combos[0]);
    if (data.warnings?.length) showToast(`${data.warnings.length} warning(s) during parsing`, "warning");
    saveSessionName(data.session_id, sessionName);
    fetchSessions();
  }, [showToast, sessionName, saveSessionName, fetchSessions]);

  const applyFullSession = useCallback((data: FullSession) => {
    setSessionId(data.session_id);
    setSessionName(data.name);
    setExported(data.exported);

    // Apply gap settings
    setGapMm(data.gap_mm);
    setColumnGapMm(data.column_gap_mm);

    // Apply Excel data
    if (data.excel_filename) {
      setExcelFile(data.excel_filename);
    } else {
      setExcelFile("");
    }

    // Apply parse data as ParseResponse
    if (data.entries_count > 0) {
      const pd: ParseResponse = {
        session_id: data.session_id,
        entries_count: data.entries_count,
        total_slots: data.total_slots,
        groups: data.groups,
        combo_count: data.combo_count,
        warnings: data.warnings || [],
        entries_preview: data.entries_preview,
      };
      setParseData(pd);
      const all = new Set<string>();
      data.groups.forEach((g) => g.combos.forEach((c) => all.add(c.filename)));
      setSelectedCombos(all);
      const allGroups = new Set<string>();
      data.groups.forEach((g) => allGroups.add(`${g.machine_program}_${g.com_no}`));
      setExpandedGroups(allGroups);
      if (data.groups[0]?.combos[0]) setPreviewCombo(data.groups[0].combos[0]);
      else setPreviewCombo(null);
    } else {
      setParseData(null);
      setSelectedCombos(new Set());
      setExpandedGroups(new Set());
      setPreviewCombo(null);
    }

    // Apply DST data
    if (data.dst_count > 0) {
      setDstUploaded(true);
      setDstFileName(`${data.dst_count} DST files`);
      setDstData({
        session_id: data.session_id,
        uploaded_count: data.dst_count,
        needed_count: data.entries_count,
        missing_programs: data.missing_programs || [],
        all_matched: data.dst_all_matched,
      });
    } else {
      setDstUploaded(false);
      setDstData(null);
      setDstFileName("");
    }

    // Reset transient UI state
    setDownloadUrl("");
    setExportProgress(0);
    setShowExcelPreview(false);
    setShowSettings(false);
  }, []);

  const loadFullSession = useCallback(async (sid: string) => {
    setLoadingSession(true);
    try {
      const res = await fetch(`${API}/api/session/${sid}/full`);
      if (!res.ok) throw new Error(`Failed to load session: ${res.status}`);
      const data: FullSession = await res.json();
      applyFullSession(data);
      setSessionStarted(true);
    } catch (e) {
      showToast(`${e instanceof Error ? e.message : "Failed to load session"}`);
    }
    setLoadingSession(false);
  }, [applyFullSession, showToast]);

  const deleteSession = useCallback(async (sid: string) => {
    try {
      const res = await fetch(`${API}/api/session/${sid}`, { method: "DELETE" });
      if (!res.ok) throw new Error("Delete failed");
      showToast("Session deleted", "success");
      fetchSessions();
      // If we deleted the active session, go back to picker
      if (sid === sessionId) {
        resetSession();
        setExcelFile("");
        setParseData(null);
        setSessionStarted(false);
        setSessionId("");
      }
    } catch (e) {
      showToast(`${e instanceof Error ? e.message : "Failed to delete session"}`);
    }
    setDeleteConfirm(null);
  }, [sessionId, resetSession, showToast, fetchSessions]);

  const removeExcel = useCallback(async () => {
    if (!sessionId) return;
    try {
      const res = await fetch(`${API}/api/session/${sessionId}/excel`, { method: "DELETE" });
      if (!res.ok) throw new Error("Failed to remove Excel");
      setExcelFile("");
      setParseData(null);
      setSelectedCombos(new Set());
      setExpandedGroups(new Set());
      setPreviewCombo(null);
      setDownloadUrl("");
      setExportProgress(0);
      setShowExcelPreview(false);
      setExported(false);
      showToast("Excel data removed", "success");
      fetchSessions();
    } catch (e) {
      showToast(`${e instanceof Error ? e.message : "Failed to remove Excel"}`);
    }
  }, [sessionId, showToast, fetchSessions]);

  const uploadExcel = useCallback(async (file: File) => {
    if (!file.name.match(/\.(xlsx|xls)$/i)) { showToast("Please upload an Excel file (.xlsx or .xls)"); return; }
    resetSession();
    setSessionId(""); // Force new server session
    setExcelLoading(true); setExcelFile(file.name);
    const form = new FormData(); form.append("file", file);
    // Don't send old sessionId — always create fresh session for new Excel
    try {
      const res = await fetch(`${API}/api/parse-excel`, { method: "POST", body: form });
      if (!res.ok) throw new Error(`Server error: ${res.status}`);
      const data = await res.json();
      if (data.entries_count === 0) {
        showToast("No valid entries found. Expected columns: A (Program), F (Name), H (Qty), O (Com No), P (Machine Program)", "warning");
        setExcelFile("");
      } else {
        applyParseData(data);
      }
    } catch (e) { showToast(`Failed to parse Excel: ${e instanceof Error ? e.message : "Connection error"}`); setExcelFile(""); }
    setExcelLoading(false);
  }, [sessionId, applyParseData, resetSession, showToast]);

  const uploadDst = useCallback(async (files: FileList | File[]) => {
    if (!sessionId) { showToast("Upload an Excel order first"); return; }
    const fileArr = Array.from(files);
    const ngsFiles = fileArr.filter((f) => f.name.toLowerCase().endsWith(".ngs"));
    if (ngsFiles.length > 0) showToast(`${ngsFiles.length} NGS file(s) skipped — convert to DST first using Wings XP on Windows`, "warning");
    const isZip = fileArr.length === 1 && fileArr[0].name.toLowerCase().endsWith(".zip");
    const dstFiles = fileArr.filter((f) => f.name.toLowerCase().endsWith(".dst"));
    if (!isZip && dstFiles.length === 0) { showToast("No DST files found. Upload .dst files or a .zip containing them."); return; }
    setDstLoading(true);
    const form = new FormData(); form.append("session_id", sessionId);
    if (isZip) { form.append("zip_file", fileArr[0]); setDstFileName(fileArr[0].name); }
    else { dstFiles.forEach((f) => form.append("files", f)); setDstFileName(`${dstFiles.length} DST files`); }
    try {
      const res = await fetch(`${API}/api/upload-dst`, { method: "POST", body: form });
      if (!res.ok) throw new Error(`Server error: ${res.status}`);
      const data: DstResponse = await res.json();
      setDstData(data); setDstUploaded(true);
      if (data.uploaded_count === 0) { showToast("No DST files found in the uploaded content"); setDstUploaded(false); }
      else if (data.missing_programs.length > 0) {
        const missing = data.missing_programs.slice(0, 5).join(", ");
        showToast(`Missing ${data.missing_programs.length} DST files: ${missing}${data.missing_programs.length > 5 ? " ..." : ""}`, "warning");
      }
    } catch (e) { showToast(`Upload failed: ${e instanceof Error ? e.message : "Connection error"}`); }
    setDstLoading(false);
  }, [sessionId, showToast]);

  const handleExport = useCallback(async () => {
    if (!sessionId || selectedCombos.size === 0) return;
    if (!dstUploaded) { showToast("Upload DST files before exporting"); return; }
    setExporting(true); setExportProgress(5); setDownloadUrl("");
    const form = new FormData();
    form.append("session_id", sessionId);
    form.append("selected_filenames", Array.from(selectedCombos).join(","));
    form.append("gap_mm", String(gapMm));
    form.append("column_gap_mm", String(columnGapMm));
    // Animate progress while waiting for server (export can take 1-3 min on free tier)
    let progress = 5;
    const progressInterval = setInterval(() => {
      progress = Math.min(progress + 0.5, 85);
      setExportProgress(Math.round(progress));
    }, 1000);
    try {
      const res = await fetch(`${API}/api/export`, { method: "POST", body: form });
      clearInterval(progressInterval);
      if (!res.ok) throw new Error(await res.text().catch(() => "Export failed"));
      setExportProgress(90);
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      setDownloadUrl(url); setExportProgress(100);
      setExported(true);
      const a = document.createElement("a");
      a.href = url; a.download = `combos_${sessionName.replace(/\s+/g, "_")}.zip`; a.click();
      showToast(`Exported ${res.headers.get("X-Export-Success") || selectedCombos.size} output files`, "success");
      fetchSessions();
    } catch (e) { clearInterval(progressInterval); showToast(`Export failed: ${e instanceof Error ? e.message : "Connection error"}`); }
    setExporting(false);
  }, [sessionId, selectedCombos, dstUploaded, sessionName, gapMm, columnGapMm, showToast, fetchSessions]);

  const loadSampleData = useCallback(async () => {
    setExcelLoading(true); resetSession();
    try {
      const res = await fetch(`${API}/api/dev/load-sample`);
      if (!res.ok) throw new Error("API not running on port 8000");
      const data = await res.json();
      setExcelFile("nameorder_04032026-2 add column.xlsx");
      applyParseData(data);
      if (data.dst_count > 0) { setDstUploaded(true); setDstFileName(`${data.dst_count} DST files`); setDstData({ session_id: data.session_id, uploaded_count: data.dst_count, needed_count: data.entries_count, missing_programs: [], all_matched: true }); }
    } catch (e) { showToast(`${e instanceof Error ? e.message : "Failed to load sample data"}`); }
    setExcelLoading(false);
  }, [applyParseData, resetSession, showToast]);

  const handleExcelDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault(); setExcelDragOver(false);
    const file = e.dataTransfer.files[0];
    if (!file) return;
    if (!file.name.match(/\.(xlsx|xls)$/i)) { showToast("Drop an Excel file (.xlsx or .xls)"); return; }
    uploadExcel(file);
  }, [uploadExcel, showToast]);

  const handleDstDrop = useCallback(async (e: React.DragEvent) => {
    e.preventDefault(); setDstDragOver(false);
    const items = e.dataTransfer.items; const files: File[] = [];
    if (items?.length) {
      for (let i = 0; i < items.length; i++) {
        const entry = items[i].webkitGetAsEntry?.();
        if (entry?.isDirectory) {
          const reader = (entry as FileSystemDirectoryEntry).createReader();
          const dirFiles = await new Promise<File[]>((resolve) => {
            reader.readEntries(async (entries) => {
              const results: File[] = [];
              for (const e of entries) { if (e.isFile) results.push(await new Promise<File>((res) => (e as FileSystemFileEntry).file(res))); }
              resolve(results);
            });
          });
          files.push(...dirFiles); continue;
        }
      }
    }
    if (!files.length) files.push(...Array.from(e.dataTransfer.files));
    if (files.length) uploadDst(files);
  }, [uploadDst]);

  const toggleCombo = (f: string) => setSelectedCombos((p) => { const n = new Set(p); n.has(f) ? n.delete(f) : n.add(f); return n; });
  const selectAll = () => { if (!parseData) return; const a = new Set<string>(); parseData.groups.forEach((g) => g.combos.forEach((c) => a.add(c.filename))); setSelectedCombos(a); };
  const deselectAll = () => setSelectedCombos(new Set());
  const toggleGroup = (k: string) => setExpandedGroups((p) => { const n = new Set(p); n.has(k) ? n.delete(k) : n.add(k); return n; });
  const totalCombos = parseData?.groups.reduce((s, g) => s + g.combos.length, 0) ?? 0;

  const startSession = () => {
    if (!sessionName.trim()) return;
    setSessionStarted(true);
  };

  const startDemoSession = () => {
    setSessionName("Demo Session");
    setSessionStarted(true);
    setTimeout(() => loadSampleData(), 100);
  };

  /* ── Session Picker Screen ── */
  if (!sessionStarted) {
    return (
      <div className="min-h-screen flex flex-col" style={{ background: "var(--background)" }}>
        {toast && <Toast message={toast.message} type={toast.type} onClose={() => setToast(null)} />}

        {/* Nav */}
        <nav className="flex items-center gap-4 px-6 py-3.5 shrink-0" style={{ borderBottom: "1px solid var(--border)" }}>
          <Link href="/"><Image src="/micro-logo.svg" alt="Micro" width={56} height={16} className="micro-logo opacity-40 hover:opacity-70 transition-opacity" /></Link>
          <div className="flex-1" />
          <button onClick={toggle} className="theme-toggle">{theme === "light" ? "☾" : "☀"}</button>
          <div className="flex items-center gap-1.5">
            <Image src="/ossia-mark.svg?v2" alt="" width={20} height={20} className="micro-logo" />
            <span className="text-sm font-bold" style={{ color: "var(--foreground)", letterSpacing: "-0.04em" }}>ossia</span>
          </div>
        </nav>

        <div className="flex-1 flex flex-col items-center px-6 py-10 overflow-y-auto">
          <div className="max-w-lg w-full" style={{ animation: "slideUp 0.5s ease" }}>
            <h1 className="text-3xl font-semibold tracking-tight mb-2" style={{ letterSpacing: "-0.03em" }}>Combo Builder</h1>
            <p className="text-sm mb-8" style={{ color: "var(--muted)" }}>
              Create a new session to start combining embroidery name programs into production files.
            </p>

            {/* New Session */}
            <div className="glass-panel p-6 mb-4">
              <label className="text-[11px] font-medium block mb-2" style={{ color: "var(--muted)" }}>Session Name</label>
              <input
                type="text"
                value={sessionName}
                onChange={(e) => setSessionName(e.target.value)}
                onKeyDown={(e) => e.key === "Enter" && startSession()}
                placeholder="e.g. 18 Mar 2026"
                className="w-full text-sm px-4 py-2.5 rounded-xl bg-transparent mb-4"
                style={{ border: "1px solid var(--border)", color: "var(--foreground)" }}
                autoFocus
              />
              <button
                className="accent-btn w-full"
                onClick={startSession}
                disabled={!sessionName.trim()}
              >
                Start Session
              </button>
            </div>

            {/* Demo shortcut */}
            <button
              onClick={startDemoSession}
              className="w-full text-center text-[11px] py-2 rounded-xl transition-colors mb-8"
              style={{ color: "var(--accent)", background: "var(--accent-glow)" }}
            >
              or load demo data to explore
            </button>

            {/* Existing Sessions */}
            {savedSessions.length > 0 && (
              <div>
                <h2 className="text-xs font-semibold mb-3" style={{ color: "var(--muted)" }}>
                  Previous Sessions
                </h2>
                <div className="flex flex-col gap-2">
                  {savedSessions.map((s) => (
                    <div
                      key={s.session_id}
                      className="glass-panel px-4 py-3 flex items-center gap-3 cursor-pointer transition-all hover:scale-[1.01]"
                      style={{ animation: "fadeSlideIn 0.3s ease" }}
                      onClick={() => loadFullSession(s.session_id)}
                    >
                      <div className="flex-1 min-w-0">
                        <div className="flex items-center gap-2">
                          <span className="text-sm font-medium truncate">{s.name}</span>
                          {s.exported && (
                            <span className="text-[9px] font-medium px-1.5 py-0.5 rounded-md shrink-0" style={{ background: "var(--success)", color: "white" }}>
                              Exported
                            </span>
                          )}
                        </div>
                        <div className="flex items-center gap-3 mt-0.5">
                          <span className="text-[10px]" style={{ color: "var(--muted)" }}>{formatDate(s.created_at)}</span>
                          {s.entries_count > 0 && (
                            <span className="text-[10px]" style={{ color: "var(--muted)" }}>
                              {s.entries_count} names
                            </span>
                          )}
                          {s.combo_count > 0 && (
                            <span className="text-[10px]" style={{ color: "var(--muted)" }}>
                              {s.combo_count} combos
                            </span>
                          )}
                        </div>
                      </div>
                      {/* Delete button */}
                      <button
                        className="text-xs w-6 h-6 flex items-center justify-center rounded-lg transition-colors shrink-0"
                        style={{ color: "var(--muted)", background: deleteConfirm === s.session_id ? "var(--danger)" : "transparent" }}
                        onClick={(e) => {
                          e.stopPropagation();
                          if (deleteConfirm === s.session_id) {
                            deleteSession(s.session_id);
                          } else {
                            setDeleteConfirm(s.session_id);
                            setTimeout(() => setDeleteConfirm(null), 3000);
                          }
                        }}
                        title={deleteConfirm === s.session_id ? "Click again to confirm" : "Delete session"}
                      >
                        <span style={{ color: deleteConfirm === s.session_id ? "white" : "var(--muted)" }}>
                          {deleteConfirm === s.session_id ? "?" : "✕"}
                        </span>
                      </button>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {loadingSession && (
              <div className="flex items-center justify-center py-8">
                <p className="text-sm" style={{ color: "var(--accent)" }}>Loading session...</p>
              </div>
            )}
          </div>
        </div>
      </div>
    );
  }

  /* ── Main Workflow ── */
  return (
    <div className="min-h-screen flex flex-col" style={{ background: "var(--background)" }}>
      {toast && <Toast message={toast.message} type={toast.type} onClose={() => setToast(null)} />}

      {/* Nav */}
      <nav className="flex items-center gap-4 px-6 py-3.5 shrink-0" style={{ borderBottom: "1px solid var(--border)", animation: "fadeIn 0.3s ease" }}>
        <Link href="/"><Image src="/micro-logo.svg" alt="Micro" width={56} height={16} className="micro-logo opacity-40 hover:opacity-70 transition-opacity" /></Link>
        <div className="relative ml-3" style={{ borderLeft: "1px solid var(--border)", paddingLeft: "12px" }}>
          <button
            onClick={() => { fetchSessions(); setShowSessionDropdown(!showSessionDropdown); }}
            className="flex items-center gap-1.5 text-[11px] font-medium py-1 px-2 rounded-lg transition-colors"
            style={{ background: showSessionDropdown ? "var(--surface)" : "transparent" }}
          >
            {sessionName}
            {exported && <span className="text-[9px] font-medium px-1 py-0.5 rounded" style={{ background: "var(--success)", color: "white" }}>Exported</span>}
            <span className="text-[8px]" style={{ color: "var(--muted)" }}>▼</span>
          </button>
          {showSessionDropdown && (
            <>
              {/* Backdrop to close dropdown */}
              <div className="fixed inset-0 z-40" onClick={() => setShowSessionDropdown(false)} />
              <div className="absolute top-full left-0 mt-1.5 py-1.5 min-w-[260px] z-50 rounded-xl overflow-hidden" style={{ background: "var(--glass-strong)", backdropFilter: "blur(24px)", border: "1px solid var(--glass-border)", boxShadow: "var(--shadow-float)", animation: "fadeSlideIn 0.15s ease" }}>
                {savedSessions.length > 0 && savedSessions.map((s) => (
                  <button
                    key={s.session_id}
                    className="w-full text-left px-3.5 py-2 text-[11px] transition-colors flex items-center justify-between hover:bg-[var(--surface-hover)]"
                    style={{ background: s.session_id === sessionId ? "var(--accent-glow)" : "transparent", color: s.session_id === sessionId ? "var(--accent)" : "var(--foreground)" }}
                    onClick={async () => {
                      setShowSessionDropdown(false);
                      if (s.session_id === sessionId) return;
                      await loadFullSession(s.session_id);
                    }}
                  >
                    <div className="flex items-center gap-1.5">
                      <span className="font-medium">{s.name}</span>
                      {s.exported && <span className="text-[8px] px-1 py-0.5 rounded" style={{ background: "var(--success)", color: "white" }}>Exported</span>}
                    </div>
                    <span className="text-[10px]" style={{ color: "var(--muted)" }}>
                      {s.entries_count > 0 ? `${s.entries_count} names` : "empty"}
                    </span>
                  </button>
                ))}
                {savedSessions.length === 0 && (
                  <p className="text-[10px] px-3.5 py-2" style={{ color: "var(--muted)" }}>No sessions yet</p>
                )}
                <div style={{ borderTop: "1px solid var(--border)", marginTop: "4px", paddingTop: "4px" }}>
                  <button
                    className="w-full text-left px-3.5 py-2 text-[11px] transition-colors hover:bg-[var(--surface-hover)]"
                    style={{ color: "var(--muted)" }}
                    onClick={() => { setShowSessionDropdown(false); resetSession(); setExcelFile(""); setParseData(null); setSessionStarted(false); setSessionId(""); }}
                  >
                    ← Back to sessions
                  </button>
                  <button
                    className="w-full text-left px-3.5 py-2 text-[11px] font-medium transition-colors hover:bg-[var(--surface-hover)]"
                    style={{ color: "var(--accent)" }}
                    onClick={() => { setShowSessionDropdown(false); resetSession(); setExcelFile(""); setParseData(null); setSessionId(""); setSessionName(new Date().toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" })); }}
                  >
                    + New Session
                  </button>
                </div>
              </div>
            </>
          )}
        </div>
        <div className="flex-1" />
        <button onClick={toggle} className="theme-toggle">{theme === "light" ? "☾" : "☀"}</button>
        <div className="flex items-center gap-1.5">
          <Image src="/ossia-mark.svg?v2" alt="" width={20} height={20} className="micro-logo" />
          <span className="text-sm font-bold" style={{ color: "var(--foreground)", letterSpacing: "-0.04em" }}>ossia</span>
        </div>
      </nav>

      <main className="flex-1 max-w-6xl w-full mx-auto px-6 py-5 flex flex-col gap-4 min-h-0">
        <div style={{ animation: "slideUp 0.4s ease" }}>
          <h1 className="text-2xl font-semibold tracking-tight" style={{ letterSpacing: "-0.02em" }}>Combo Builder</h1>
          <p className="text-[11px] mt-0.5" style={{ color: "var(--muted)" }}>Upload an Excel order and DST programs to generate combo files</p>
        </div>

        {/* Upload Zones */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-3" style={{ animation: "slideUp 0.4s ease 0.05s forwards", opacity: 0 }}>
          <div className={`drop-zone ${excelDragOver ? "drag-over" : ""} ${excelFile ? "has-file" : ""}`} onDragOver={(e) => { e.preventDefault(); setExcelDragOver(true); }} onDragLeave={() => setExcelDragOver(false)} onDrop={handleExcelDrop} onClick={() => excelInputRef.current?.click()}>
            <input ref={excelInputRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={(e) => { if (e.target.files?.[0]) uploadExcel(e.target.files[0]); e.target.value = ""; }} />
            {excelLoading ? <p className="text-sm" style={{ color: "var(--accent)" }}>Parsing...</p>
            : excelFile ? (
              <div style={{ animation: "fadeSlideIn 0.3s ease" }}>
                <div className="flex items-center justify-center gap-2">
                  <span style={{ color: "var(--success)", fontSize: "16px" }}>✓</span>
                  <span className="text-sm font-medium" style={{ color: "var(--success)" }}>{excelFile}</span>
                  <button
                    className="w-5 h-5 flex items-center justify-center rounded-md text-[10px] transition-colors"
                    style={{ color: "var(--muted)", background: "var(--surface)" }}
                    onClick={(e) => { e.stopPropagation(); removeExcel(); }}
                    title="Remove Excel"
                  >
                    ✕
                  </button>
                </div>
              </div>
            )
            : <div><svg className="mx-auto mb-2.5 opacity-25" width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" /><polyline points="14 2 14 8 20 8" /><line x1="16" y1="13" x2="8" y2="13" /><line x1="16" y1="17" x2="8" y2="17" /></svg><p className="text-sm font-medium mb-0.5">Order Excel</p><p className="text-[11px]" style={{ color: "var(--muted)" }}>Drop .xlsx or click to browse</p></div>}
          </div>

          <div className={`drop-zone ${dstDragOver ? "drag-over" : ""} ${dstUploaded ? "has-file" : ""} ${!sessionId ? "opacity-25 pointer-events-none" : ""}`} onDragOver={(e) => { e.preventDefault(); setDstDragOver(true); }} onDragLeave={() => setDstDragOver(false)} onDrop={handleDstDrop} onClick={() => sessionId && zipInputRef.current?.click()}>
            <input ref={zipInputRef} type="file" accept=".zip" className="hidden" onChange={(e) => { if (e.target.files) uploadDst(e.target.files); e.target.value = ""; }} />
            {dstLoading ? <p className="text-sm" style={{ color: "var(--accent)" }}>Uploading...</p>
            : dstUploaded && dstData ? <div style={{ animation: "fadeSlideIn 0.3s ease" }}><div className="flex items-center justify-center gap-2"><span style={{ color: dstData.all_matched ? "var(--success)" : "var(--warning)", fontSize: "16px" }}>{dstData.all_matched ? "✓" : "⚠"}</span><span className="text-sm font-medium" style={{ color: dstData.all_matched ? "var(--success)" : "var(--warning)" }}>{dstFileName}</span></div><p className="text-[11px] mt-1" style={{ color: "var(--muted)" }}>{dstData.uploaded_count}/{dstData.needed_count} programs matched{dstData.missing_programs.length > 0 && <span style={{ color: "var(--danger)" }}> · {dstData.missing_programs.length} missing</span>}</p></div>
            : <div><svg className="mx-auto mb-2.5 opacity-25" width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" /><polyline points="17 8 12 3 7 8" /><line x1="12" y1="3" x2="12" y2="15" /></svg><p className="text-sm font-medium mb-0.5">DST Programs</p><p className="text-[11px]" style={{ color: "var(--muted)" }}>{sessionId ? "Drop folder or click to browse .zip" : "Upload Excel first"}</p></div>}
          </div>
        </div>

        {/* Excel Preview (collapsible) */}
        {parseData?.entries_preview && parseData.entries_preview.length > 0 && (
          <div style={{ animation: "fadeSlideIn 0.3s ease" }}>
            <button onClick={() => setShowExcelPreview(!showExcelPreview)} className="text-[11px] flex items-center gap-1.5 mb-1" style={{ color: "var(--accent)" }}>
              <span className="text-[9px]">{showExcelPreview ? "▼" : "▶"}</span>
              {showExcelPreview ? "Hide" : "View"} order data ({parseData.entries_count} rows)
            </button>
            {showExcelPreview && (
              <div className="glass-panel p-3 overflow-auto custom-scroll" style={{ animation: "fadeSlideIn 0.2s ease", maxHeight: "320px" }}>
                <table className="w-full text-[10px]" style={{ fontFamily: "var(--font-geist-mono)" }}>
                  <thead className="sticky top-0" style={{ background: "var(--glass-strong)" }}>
                    <tr style={{ borderBottom: "1px solid var(--border)" }}>
                      {["Program", "Name", "Title", "Qty", "Combo", "Machine", "→ Group"].map((h) => (
                        <th key={h} className="text-left py-1.5 px-2 font-semibold" style={{ color: "var(--muted)" }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {parseData.entries_preview.map((e, i) => (
                      <tr key={i} style={{ borderBottom: "1px solid var(--border)" }}>
                        <td className="py-1 px-2" style={{ color: "var(--accent)" }}>{e.program}</td>
                        <td className="py-1 px-2">{e.name_line1}</td>
                        <td className="py-1 px-2" style={{ color: "var(--muted)" }}>{e.name_line2 || "—"}</td>
                        <td className="py-1 px-2">{e.quantity}</td>
                        <td className="py-1 px-2">{e.com_no}</td>
                        <td className="py-1 px-2" style={{ color: "var(--muted)" }}>{e.machine_program}</td>
                        <td className="py-1 px-2" style={{ color: "var(--accent)", fontWeight: 500 }}>{e.machine_program}/{e.com_no}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                {parseData.entries_count > parseData.entries_preview.length && (
                  <p className="text-[10px] mt-2 px-2" style={{ color: "var(--muted)" }}>Showing {parseData.entries_preview.length} of {parseData.entries_count} rows</p>
                )}
              </div>
            )}
          </div>
        )}

        {/* Settings (collapsible) */}
        {parseData && (
          <div style={{ animation: "fadeSlideIn 0.3s ease 0.05s forwards", opacity: 0 }}>
            <button onClick={() => setShowSettings(!showSettings)} className="text-[11px] flex items-center gap-1.5 mb-1" style={{ color: "var(--muted)" }}>
              <span className="text-[9px]">{showSettings ? "▼" : "▶"}</span>
              Settings
            </button>
            {showSettings && (
              <div className="glass-panel p-4 flex gap-6" style={{ animation: "fadeSlideIn 0.2s ease" }}>
                <label className="flex items-center gap-2 text-[11px]">
                  <span style={{ color: "var(--muted)" }}>Vertical gap</span>
                  <input type="number" value={gapMm} onChange={(e) => setGapMm(Number(e.target.value))} min={0} max={20} step={0.5} className="w-14 text-center text-[11px] px-2 py-1 rounded-lg bg-transparent" style={{ border: "1px solid var(--border)", color: "var(--foreground)" }} />
                  <span style={{ color: "var(--muted)" }}>mm</span>
                </label>
                <label className="flex items-center gap-2 text-[11px]">
                  <span style={{ color: "var(--muted)" }}>Column gap</span>
                  <input type="number" value={columnGapMm} onChange={(e) => setColumnGapMm(Number(e.target.value))} min={0} max={30} step={0.5} className="w-14 text-center text-[11px] px-2 py-1 rounded-lg bg-transparent" style={{ border: "1px solid var(--border)", color: "var(--foreground)" }} />
                  <span style={{ color: "var(--muted)" }}>mm</span>
                </label>
              </div>
            )}
          </div>
        )}

        {/* Pipeline Stats */}
        {parseData && (
          <div className="flex items-center justify-center gap-3" style={{ animation: "slideUp 0.4s ease 0.1s forwards", opacity: 0, overflow: "visible", paddingTop: "12px" }}>
            <StatCard value={parseData.entries_count} label="Names" description="Individual name entries from Excel" delay={0.05} />
            <span className="stat-arrow">→</span>
            <StatCard value={parseData.groups.length} label="Groups" description="Grouped by machine program + combo number" delay={0.1} />
            <span className="stat-arrow">→</span>
            <StatCard value={parseData.combo_count} label="Output Files" description="DST files to generate (max 20 names each)" delay={0.15} />
            <span className="stat-arrow">→</span>
            <StatCard value={parseData.total_slots} label="Slots" description="Total name positions across all files" delay={0.2} />
          </div>
        )}

        {/* Two-Panel: Combo List + Slot Preview */}
        {parseData && parseData.groups.length > 0 && (
          <div className="flex-1 grid grid-cols-1 lg:grid-cols-[1fr_360px] gap-3 min-h-0" style={{ animation: "slideUp 0.4s ease 0.15s forwards", opacity: 0 }}>
            <div className="glass-panel overflow-hidden flex flex-col min-h-[300px]">
              <div className="flex items-center justify-between px-4 py-2.5 shrink-0" style={{ borderBottom: "1px solid var(--border)" }}>
                <span className="text-xs font-semibold">Output Files <span className="font-normal" style={{ color: "var(--muted)" }}>{selectedCombos.size}/{totalCombos}</span></span>
                <div className="flex gap-1"><button onClick={selectAll} className="glass-btn text-[10px]">All</button><button onClick={deselectAll} className="glass-btn text-[10px]">None</button></div>
              </div>
              <div className="flex-1 overflow-y-auto custom-scroll">
                {parseData.groups.map((group) => {
                  const gk = `${group.machine_program}_${group.com_no}`;
                  const exp = expandedGroups.has(gk);
                  const sel = group.combos.filter((c) => selectedCombos.has(c.filename)).length;
                  return (
                    <div key={gk} style={{ borderBottom: "1px solid var(--border)" }}>
                      <button className="w-full flex flex-col px-4 py-2.5 text-left transition-colors" style={{ background: exp ? "var(--surface)" : "transparent" }} onClick={() => toggleGroup(gk)}>
                        <div className="flex items-center gap-2.5 w-full">
                          <span className="text-[10px] w-3" style={{ color: "var(--muted)" }}>{exp ? "▼" : "▶"}</span>
                          <span className="text-xs font-medium">{group.machine_program}<span style={{ color: "var(--muted)", fontWeight: 400 }}> / Combo {group.com_no}</span></span>
                          <span className="text-[10px] ml-auto tabular-nums" style={{ color: "var(--muted)" }}>
                            {group.entry_count} names → {group.combos.length} {group.combos.length === 1 ? "file" : "files"} ({group.total_slots} slots)
                          </span>
                        </div>
                      </button>
                      {exp && group.combos.map((combo, ci) => (
                        <div key={combo.filename} className="flex items-center gap-2.5 px-4 py-2 cursor-pointer transition-all" style={{ background: previewCombo?.filename === combo.filename ? "var(--accent-glow)" : "var(--surface)", borderTop: "1px solid var(--border)", animation: `fadeSlideIn 0.15s ease ${ci * 0.02}s forwards`, opacity: 0 }} onClick={() => setPreviewCombo(combo)}>
                          <input type="checkbox" className="custom-checkbox" checked={selectedCombos.has(combo.filename)} onChange={() => toggleCombo(combo.filename)} onClick={(e) => e.stopPropagation()} />
                          <span className="text-[11px] font-mono">{combo.filename}</span>
                          <span className="text-[10px] ml-auto tabular-nums" style={{ color: "var(--muted)" }}>{combo.left_count} left{combo.right_count > 0 ? ` + ${combo.right_count} right` : ""}</span>
                        </div>
                      ))}
                    </div>
                  );
                })}
              </div>
            </div>

            <div className="glass-panel p-4 overflow-y-auto custom-scroll lg:sticky lg:top-0 lg:self-start min-h-[300px]" style={{ maxHeight: "calc(100vh - 340px)" }}>
              {previewCombo ? (
                <>
                  <h3 className="text-[11px] font-semibold font-mono mb-3 pb-2" style={{ borderBottom: "1px solid var(--border)", color: "var(--accent)" }}>{previewCombo.filename}</h3>
                  <div className="grid grid-cols-2 gap-4">
                    <div>
                      <p className="text-[9px] font-semibold uppercase tracking-wider mb-1.5" style={{ color: "var(--muted)" }}>Left Column</p>
                      {previewCombo.slots.slice(0, 10).map((s, i) => (
                        <div key={i} className="flex items-center gap-1 py-[2px]" style={{ borderBottom: "1px solid var(--border)" }}>
                          <span className="text-[9px] w-3.5 text-right tabular-nums" style={{ color: "var(--border-strong)" }}>{i + 1}</span>
                          <span className="text-[9px] font-mono w-6 tabular-nums" style={{ color: "var(--accent)" }}>{s.program}</span>
                          <span className="text-[9px] truncate">{s.name_line1}</span>
                        </div>
                      ))}
                      {Array.from({ length: Math.max(0, 10 - previewCombo.left_count) }).map((_, i) => (
                        <div key={`e${i}`} className="flex items-center gap-1 py-[2px]" style={{ borderBottom: "1px solid var(--border)" }}>
                          <span className="text-[9px] w-3.5 text-right tabular-nums" style={{ color: "var(--border-strong)" }}>{previewCombo.left_count + i + 1}</span>
                          <span className="text-[9px]" style={{ color: "var(--border-strong)" }}>—</span>
                        </div>
                      ))}
                    </div>
                    <div>
                      <p className="text-[9px] font-semibold uppercase tracking-wider mb-1.5" style={{ color: "var(--muted)" }}>Right Column</p>
                      {previewCombo.slots.slice(10).map((s, i) => (
                        <div key={i} className="flex items-center gap-1 py-[2px]" style={{ borderBottom: "1px solid var(--border)" }}>
                          <span className="text-[9px] w-3.5 text-right tabular-nums" style={{ color: "var(--border-strong)" }}>{i + 11}</span>
                          <span className="text-[9px] font-mono w-6 tabular-nums" style={{ color: "var(--accent)" }}>{s.program}</span>
                          <span className="text-[9px] truncate">{s.name_line1}</span>
                        </div>
                      ))}
                      {previewCombo.right_count === 0 && <p className="text-[9px] py-2" style={{ color: "var(--border-strong)" }}>No right column</p>}
                    </div>
                  </div>
                </>
              ) : (
                <div className="flex items-center justify-center h-full"><p className="text-[11px]" style={{ color: "var(--muted)" }}>Click a combo to preview</p></div>
              )}
            </div>
          </div>
        )}
      </main>

      {/* Export Bar */}
      {parseData && (
        <div className="shrink-0 px-6 py-3" style={{ background: "var(--glass-strong)", backdropFilter: "blur(24px)", borderTop: "1px solid var(--glass-border)", animation: "slideUp 0.3s ease 0.2s forwards", opacity: 0 }}>
          <div className="max-w-6xl mx-auto flex items-center gap-4">
            <div className="flex-1 min-w-0">
              {exporting && (
                <div>
                  <div className="flex items-center gap-2 mb-1">
                    <div className="animate-spin w-4 h-4 border-2 rounded-full" style={{ borderColor: "var(--accent)", borderTopColor: "transparent" }} />
                    <p className="text-xs font-medium" style={{ color: "var(--foreground)" }}>Exporting... {exportProgress}%</p>
                  </div>
                  <div className="progress-bar" style={{ height: "8px" }}>
                    <div className="progress-bar-fill" style={{ width: `${exportProgress}%`, transition: "width 1s ease" }} />
                  </div>
                  <p className="text-[10px] mt-1" style={{ color: "var(--muted)" }}>Combining {selectedCombos.size} files on server... this may take a minute</p>
                </div>
              )}
              {downloadUrl && !exporting && (
                <div className="flex items-center gap-2">
                  <p className="text-[11px]" style={{ color: "var(--success)" }}>✓ Downloaded combos_{sessionName.replace(/\s+/g, "_")}.zip</p>
                  <a href={downloadUrl} download={`combos_${sessionName.replace(/\s+/g, "_")}.zip`} className="text-[11px] underline" style={{ color: "var(--accent)" }}>Download again</a>
                </div>
              )}
              {!exporting && !downloadUrl && exported && (
                <p className="text-[11px]" style={{ color: "var(--success)" }}>✓ Previously exported</p>
              )}
              {!exporting && !downloadUrl && !exported && !dstUploaded && <p className="text-[11px]" style={{ color: "var(--muted)" }}>Upload DST files to enable export</p>}
            </div>
            <button className="accent-btn" disabled={selectedCombos.size === 0 || !dstUploaded || exporting} onClick={handleExport}>
              {exporting ? "Exporting..." : `Export ${selectedCombos.size} File${selectedCombos.size !== 1 ? "s" : ""}`}
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
