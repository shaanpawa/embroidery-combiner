"use client";

import Link from "next/link";
import Image from "next/image";
import { useState, useCallback, useRef, useEffect } from "react";
import { useSession, signOut } from "next-auth/react";
import { useTheme } from "../theme-provider";
import { useLanguage } from "../i18n";
import { authFetch, clearAuthToken, warmupBackend } from "@/lib/api";

const API = process.env.NEXT_PUBLIC_API_URL || "http://localhost:8000";

interface Slot { program: number; name_line1: string; name_line2: string; quantity: number; }
interface ComboFile { filename: string; part_number: number; total_parts: number; slot_count: number; left_count: number; right_count: number; slots: Slot[]; }
interface Group { machine_program: string; com_no: string; entry_count: number; total_slots: number; combos: ComboFile[]; }
interface EntryPreview { program: number; name_line1: string; name_line2: string; quantity: number; com_no: string; machine_program: string; }
interface ParseResponse { session_id: string; entries_count: number; total_slots: number; groups: Group[]; combo_count: number; warnings: string[]; entries_preview?: EntryPreview[]; }
interface DetectResponse { session_id: string; excel_filename: string; headers: string[]; preview_rows: (string | number | null)[][]; detected_mapping: Record<string, number>; confidence: string; }
const FIELD_KEYS = ["program", "name_line1", "name_line2", "quantity", "com_no", "machine_program"] as const;
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
  const { lang, toggle: toggleLang, t } = useLanguage();
  const { data: session } = useSession();
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
  const [exportElapsed, setExportElapsed] = useState(0);
  const [excelDragOver, setExcelDragOver] = useState(false);
  const [dstDragOver, setDstDragOver] = useState(false);
  const [toast, setToast] = useState<{ message: string; type: "error" | "warning" | "success" } | null>(null);
  const [showExcelPreview, setShowExcelPreview] = useState(false);
  const [showSettings, setShowSettings] = useState(false);
  const [showMobilePreview, setShowMobilePreview] = useState(false);
  const [detectData, setDetectData] = useState<DetectResponse | null>(null);
  const [columnMapping, setColumnMapping] = useState<Record<string, number>>({});
  const [mappingConfirmed, setMappingConfirmed] = useState(false);
  const [gapMm, setGapMm] = useState(3);
  const [columnGapMm, setColumnGapMm] = useState(5);
  const excelInputRef = useRef<HTMLInputElement>(null);
  const zipInputRef = useRef<HTMLInputElement>(null);

  const showToast = useCallback((message: string, type: "error" | "warning" | "success" = "error") => setToast({ message, type }), []);

  const fetchSessions = useCallback(async () => {
    try {
      const res = await authFetch(`${API}/api/sessions`);
      if (res.ok) setSavedSessions(await res.json());
    } catch { /* ignore */ }
  }, []);

  // Fetch sessions + warm up backend on mount (wakes Render free tier from sleep)
  useEffect(() => { fetchSessions(); warmupBackend(API); }, [fetchSessions]);

  // Save session name to server when it changes
  const saveSessionName = useCallback(async (sid: string, name: string) => {
    const form = new FormData(); form.append("session_id", sid); form.append("name", name);
    try { await authFetch(`${API}/api/session/name`, { method: "POST", body: form }); } catch { /* ignore */ }
  }, []);

  const resetSession = useCallback(() => {
    setDetectData(null); setColumnMapping({}); setMappingConfirmed(false);
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
    if (data.warnings?.length) showToast(`${data.warnings.length} ${t("ok.warnings")}`, "warning");
    saveSessionName(data.session_id, sessionName);
    fetchSessions();
  }, [showToast, sessionName, saveSessionName, fetchSessions, t]);

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
      const res = await authFetch(`${API}/api/session/${sid}/full`);
      if (!res.ok) throw new Error(`Failed to load session: ${res.status}`);
      const data: FullSession = await res.json();
      applyFullSession(data);
      setSessionStarted(true);
    } catch (e) {
      showToast(`${e instanceof Error ? e.message : t("err.load_session")}`);
    }
    setLoadingSession(false);
  }, [applyFullSession, showToast, t]);

  const deleteSession = useCallback(async (sid: string) => {
    try {
      const res = await authFetch(`${API}/api/session/${sid}`, { method: "DELETE" });
      if (!res.ok) throw new Error("Delete failed");
      showToast(t("ok.session_deleted"), "success");
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
      showToast(`${e instanceof Error ? e.message : t("err.delete_session")}`);
    }
    setDeleteConfirm(null);
  }, [sessionId, resetSession, showToast, fetchSessions, t]);

  const removeExcel = useCallback(async () => {
    if (!sessionId) return;
    try {
      const res = await authFetch(`${API}/api/session/${sessionId}/excel`, { method: "DELETE" });
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
      showToast(t("ok.excel_removed"), "success");
      fetchSessions();
    } catch (e) {
      showToast(`${e instanceof Error ? e.message : t("err.excel_read")}`);
    }
  }, [sessionId, showToast, fetchSessions, t]);

  const uploadExcel = useCallback(async (file: File) => {
    if (!file.name.match(/\.(xlsx|xls)$/i)) { showToast(t("err.excel_format")); return; }
    resetSession();
    setSessionId(""); // Force new server session
    setExcelLoading(true); setExcelFile(file.name);
    const form = new FormData(); form.append("file", file);
    try {
      const res = await authFetch(`${API}/api/detect-columns`, { method: "POST", body: form });
      if (!res.ok) throw new Error(`Server error: ${res.status}`);
      const data: DetectResponse = await res.json();
      setDetectData(data);
      setColumnMapping(data.detected_mapping);
      setSessionId(data.session_id);
      saveSessionName(data.session_id, sessionName);
      fetchSessions();
    } catch (e) { showToast(`${t("err.excel_read")}: ${e instanceof Error ? e.message : t("err.connection")}`); setExcelFile(""); }
    setExcelLoading(false);
  }, [sessionId, resetSession, showToast, sessionName, saveSessionName, fetchSessions, t]);

  const confirmMapping = useCallback(async () => {
    if (!sessionId || !detectData) return;
    setExcelLoading(true);
    const form = new FormData();
    form.append("session_id", sessionId);
    form.append("column_map", JSON.stringify(columnMapping));
    try {
      const res = await authFetch(`${API}/api/parse-excel`, { method: "POST", body: form });
      if (!res.ok) throw new Error(`Server error: ${res.status}`);
      const data = await res.json();
      if (data.entries_count === 0) {
        showToast(t("err.no_entries"), "warning");
      } else {
        setMappingConfirmed(true);
        applyParseData(data);
      }
    } catch (e) { showToast(`${t("err.parse_fail")}: ${e instanceof Error ? e.message : t("err.connection")}`); }
    setExcelLoading(false);
  }, [sessionId, detectData, columnMapping, applyParseData, showToast, t]);

  const uploadDst = useCallback(async (files: FileList | File[]) => {
    if (!sessionId) { showToast(t("err.upload_excel_first")); return; }
    const fileArr = Array.from(files);
    const ngsFiles = fileArr.filter((f) => f.name.toLowerCase().endsWith(".ngs"));
    if (ngsFiles.length > 0) showToast(`${ngsFiles.length} ${t("err.ngs_skipped")}`, "warning");
    const isZip = fileArr.length === 1 && fileArr[0].name.toLowerCase().endsWith(".zip");
    const dstFiles = fileArr.filter((f) => f.name.toLowerCase().endsWith(".dst"));
    if (!isZip && dstFiles.length === 0) { showToast(t("err.no_dst")); return; }
    setDstLoading(true);
    const form = new FormData(); form.append("session_id", sessionId);
    if (isZip) { form.append("zip_file", fileArr[0]); setDstFileName(fileArr[0].name); }
    else { dstFiles.forEach((f) => form.append("files", f)); setDstFileName(`${dstFiles.length} DST files`); }
    try {
      const res = await authFetch(`${API}/api/upload-dst`, { method: "POST", body: form });
      if (!res.ok) throw new Error(`Server error: ${res.status}`);
      const data: DstResponse = await res.json();
      setDstData(data); setDstUploaded(true);
      if (data.uploaded_count === 0) { showToast(t("err.no_dst_content")); setDstUploaded(false); }
      else if (data.missing_programs.length > 0) {
        const missing = data.missing_programs.slice(0, 5).join(", ");
        showToast(`${t("err.missing_dst").replace("{n}", String(data.missing_programs.length))}: ${missing}${data.missing_programs.length > 5 ? " ..." : ""}`, "warning");
      }
    } catch (e) { showToast(`${t("err.upload_fail")}: ${e instanceof Error ? e.message : t("err.connection")}`); }
    setDstLoading(false);
  }, [sessionId, showToast, t]);

  const handleExport = useCallback(async () => {
    if (!sessionId || selectedCombos.size === 0) return;
    if (!dstUploaded) { showToast(t("err.upload_dst_first")); return; }
    setExporting(true); setExportProgress(0); setDownloadUrl("");
    const form = new FormData();
    form.append("session_id", sessionId);
    form.append("selected_filenames", Array.from(selectedCombos).join(","));
    form.append("gap_mm", String(gapMm));
    form.append("column_gap_mm", String(columnGapMm));
    // Track elapsed time for user feedback
    const startTime = Date.now();
    const elapsedTimer = setInterval(() => {
      const elapsed = Math.round((Date.now() - startTime) / 1000);
      setExportProgress(elapsed); // Use as seconds counter, not percentage
    }, 1000);
    try {
      // Export can take 2-3 min on free tier — use 5 min timeout
      const res = await authFetch(`${API}/api/export`, { method: "POST", body: form }, 300_000);
      clearInterval(elapsedTimer);
      setExportElapsed(Math.round((Date.now() - startTime) / 1000));
      if (!res.ok) throw new Error(await res.text().catch(() => "Export failed"));
      setExportProgress(-1); // Signal: downloading
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      setDownloadUrl(url); setExportProgress(-2); // Signal: complete
      setExported(true);
      const a = document.createElement("a");
      a.href = url; a.download = `combos_${sessionName.replace(/\s+/g, "_")}.zip`; a.click();
      showToast(t("cb.export.success").replace("{n}", String(res.headers.get("X-Export-Success") || selectedCombos.size)), "success");
      fetchSessions();
    } catch (e) { clearInterval(elapsedTimer); showToast(`${t("err.export_fail")}: ${e instanceof Error ? e.message : t("err.connection")}`); }
    setExporting(false); setExportProgress(0);
  }, [sessionId, selectedCombos, dstUploaded, sessionName, gapMm, columnGapMm, showToast, fetchSessions, t]);

  const loadSampleData = useCallback(async () => {
    setExcelLoading(true); resetSession();
    try {
      const res = await authFetch(`${API}/api/dev/load-sample`);
      if (!res.ok) throw new Error("API not running on port 8000");
      const data = await res.json();
      setExcelFile("nameorder_04032026-2 add column.xlsx");
      applyParseData(data);
      if (data.dst_count > 0) { setDstUploaded(true); setDstFileName(`${data.dst_count} DST files`); setDstData({ session_id: data.session_id, uploaded_count: data.dst_count, needed_count: data.entries_count, missing_programs: [], all_matched: true }); }
    } catch (e) { showToast(`${e instanceof Error ? e.message : t("err.load_sample")}`); }
    setExcelLoading(false);
  }, [applyParseData, resetSession, showToast, t]);

  const handleExcelDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault(); setExcelDragOver(false);
    const file = e.dataTransfer.files[0];
    if (!file) return;
    if (!file.name.match(/\.(xlsx|xls)$/i)) { showToast(t("err.excel_drop")); return; }
    uploadExcel(file);
  }, [uploadExcel, showToast, t]);

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
          <button onClick={toggleLang} className="text-[10px] font-semibold px-3 py-1.5 rounded-lg" style={{ background: "var(--surface)", border: "1px solid var(--border)", color: "var(--muted)" }}>{lang === "en" ? "TH" : "EN"}</button>
          <button onClick={toggle} className="theme-toggle">{theme === "light" ? "☾" : "☀"}</button>
          <div className="hidden sm:flex items-center gap-1.5">
            <Image src="/ossia-mark.svg?v2" alt="" width={20} height={20} className="micro-logo" />
            <span className="text-sm font-bold" style={{ color: "var(--foreground)", letterSpacing: "-0.04em" }}>ossia</span>
          </div>
        </nav>

        <div className="flex-1 flex flex-col items-center px-4 sm:px-6 py-6 sm:py-10 overflow-y-auto">
          <div className="max-w-lg w-full" style={{ animation: "slideUp 0.5s ease" }}>
            <h1 className="text-2xl sm:text-3xl font-semibold tracking-tight mb-2" style={{ letterSpacing: "-0.03em" }}>{t("cb.title")}</h1>
            <p className="text-sm mb-8" style={{ color: "var(--muted)" }}>
              {t("cb.session.subtitle")}
            </p>

            {/* New Session */}
            <div className="glass-panel p-6 mb-4">
              <label className="text-[11px] font-medium block mb-2" style={{ color: "var(--muted)" }}>{t("cb.session.name")}</label>
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
                {t("cb.session.start")}
              </button>
            </div>

            {/* Demo shortcut */}
            <button
              onClick={startDemoSession}
              className="w-full text-center text-[11px] py-2 rounded-xl transition-colors mb-8"
              style={{ color: "var(--accent)", background: "var(--accent-glow)" }}
            >
              {t("cb.session.demo")}
            </button>

            {/* Existing Sessions */}
            {savedSessions.length > 0 && (
              <div>
                <h2 className="text-xs font-semibold mb-3" style={{ color: "var(--muted)" }}>
                  {t("cb.session.previous")}
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
                              {t("cb.exported")}
                            </span>
                          )}
                        </div>
                        <div className="flex items-center gap-3 mt-0.5">
                          <span className="text-[10px]" style={{ color: "var(--muted)" }}>{formatDate(s.created_at)}</span>
                          {s.entries_count > 0 && (
                            <span className="text-[10px]" style={{ color: "var(--muted)" }}>
                              {s.entries_count} {t("names")}
                            </span>
                          )}
                          {s.combo_count > 0 && (
                            <span className="text-[10px]" style={{ color: "var(--muted)" }}>
                              {s.combo_count} {t("combos")}
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
                        title={deleteConfirm === s.session_id ? t("cb.session.delete_confirm") : t("cb.session.delete")}
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
                <p className="text-sm" style={{ color: "var(--accent)" }}>{t("cb.session.loading")}</p>
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
            {exported && <span className="text-[9px] font-medium px-1 py-0.5 rounded" style={{ background: "var(--success)", color: "white" }}>{t("cb.exported")}</span>}
            <span className="text-[8px]" style={{ color: "var(--muted)" }}>▼</span>
          </button>
          {showSessionDropdown && (
            <>
              {/* Backdrop to close dropdown */}
              <div className="fixed inset-0 z-40" onClick={() => setShowSessionDropdown(false)} />
              <div className="absolute top-full left-0 mt-1.5 py-1.5 min-w-[260px] max-w-[calc(100vw-2rem)] z-50 rounded-xl overflow-hidden" style={{ background: "var(--glass-strong)", backdropFilter: "blur(24px)", border: "1px solid var(--glass-border)", boxShadow: "var(--shadow-float)", animation: "fadeSlideIn 0.15s ease" }}>
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
                      {s.exported && <span className="text-[8px] px-1 py-0.5 rounded" style={{ background: "var(--success)", color: "white" }}>{t("cb.exported")}</span>}
                    </div>
                    <span className="text-[10px]" style={{ color: "var(--muted)" }}>
                      {s.entries_count > 0 ? `${s.entries_count} ${t("names")}` : t("empty")}
                    </span>
                  </button>
                ))}
                {savedSessions.length === 0 && (
                  <p className="text-[10px] px-3.5 py-2" style={{ color: "var(--muted)" }}>{t("cb.session.no_sessions")}</p>
                )}
                <div style={{ borderTop: "1px solid var(--border)", marginTop: "4px", paddingTop: "4px" }}>
                  <button
                    className="w-full text-left px-3.5 py-2 text-[11px] transition-colors hover:bg-[var(--surface-hover)]"
                    style={{ color: "var(--muted)" }}
                    onClick={() => { setShowSessionDropdown(false); resetSession(); setExcelFile(""); setParseData(null); setSessionStarted(false); setSessionId(""); }}
                  >
                    {t("cb.session.back")}
                  </button>
                  <button
                    className="w-full text-left px-3.5 py-2 text-[11px] font-medium transition-colors hover:bg-[var(--surface-hover)]"
                    style={{ color: "var(--accent)" }}
                    onClick={() => { setShowSessionDropdown(false); resetSession(); setExcelFile(""); setParseData(null); setSessionId(""); setSessionName(new Date().toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" })); }}
                  >
                    {t("cb.session.new")}
                  </button>
                </div>
              </div>
            </>
          )}
        </div>
        <div className="flex-1" />
        {session?.user && (
          <button
            onClick={() => { clearAuthToken(); signOut({ callbackUrl: "/login" }); }}
            className="text-[10px] font-semibold px-3 py-1.5 rounded-lg transition-colors"
            style={{ background: "var(--surface)", border: "1px solid var(--border)", color: "var(--muted)" }}
            title={session.user.email || ""}
          >
            <span className="hidden sm:inline">{session.user.name || "O"}</span>
            <span className="sm:hidden">{(session.user.name || "O").charAt(0).toUpperCase()}</span>
            <span className="ml-1.5" style={{ fontSize: "9px" }}>{t("nav.signout")}</span>
          </button>
        )}
        <button onClick={toggleLang} className="text-[10px] font-semibold px-3 py-1.5 rounded-lg" style={{ background: "var(--surface)", border: "1px solid var(--border)", color: "var(--muted)" }}>{lang === "en" ? "TH" : "EN"}</button>
        <button onClick={toggle} className="theme-toggle">{theme === "light" ? "☾" : "☀"}</button>
        <div className="hidden sm:flex items-center gap-1.5">
          <Image src="/ossia-mark.svg?v2" alt="" width={20} height={20} className="micro-logo" />
          <span className="text-sm font-bold" style={{ color: "var(--foreground)", letterSpacing: "-0.04em" }}>ossia</span>
        </div>
      </nav>

      <main className="flex-1 max-w-6xl w-full mx-auto px-4 sm:px-6 py-4 sm:py-5 flex flex-col gap-3 sm:gap-4 min-h-0">
        <div style={{ animation: "slideUp 0.4s ease" }}>
          <h1 className="text-2xl font-semibold tracking-tight" style={{ letterSpacing: "-0.02em" }}>{t("cb.title")}</h1>
          <p className="text-[11px] mt-0.5" style={{ color: "var(--muted)" }}>{t("cb.subtitle")}</p>
        </div>

        {/* Upload Zones */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-3" style={{ animation: "slideUp 0.4s ease 0.05s forwards", opacity: 0 }}>
          <div className={`drop-zone ${excelDragOver ? "drag-over" : ""} ${excelFile ? "has-file" : ""}`} onDragOver={(e) => { e.preventDefault(); setExcelDragOver(true); }} onDragLeave={() => setExcelDragOver(false)} onDrop={handleExcelDrop} onClick={() => excelInputRef.current?.click()}>
            <input ref={excelInputRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={(e) => { if (e.target.files?.[0]) uploadExcel(e.target.files[0]); e.target.value = ""; }} />
            {excelLoading ? <p className="text-sm" style={{ color: "var(--accent)" }}>{t("cb.excel.parsing")}</p>
            : excelFile ? (
              <div style={{ animation: "fadeSlideIn 0.3s ease" }}>
                <div className="flex items-center justify-center gap-2">
                  <span style={{ color: "var(--accent)", fontSize: "16px" }}>✓</span>
                  <span className="text-sm font-medium" style={{ color: "var(--accent)" }}>{excelFile}</span>
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
            : <div><span className="text-[9px] font-semibold uppercase tracking-wider" style={{ color: "var(--accent)", opacity: 0.5 }}>{t("cb.step1")}</span><svg className="mx-auto mb-2.5 mt-1.5 opacity-25" width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" /><polyline points="14 2 14 8 20 8" /><line x1="16" y1="13" x2="8" y2="13" /><line x1="16" y1="17" x2="8" y2="17" /></svg><p className="text-sm font-medium mb-0.5">{t("cb.excel.title")}</p><p className="text-[11px]" style={{ color: "var(--muted)" }}>{t("cb.excel.hint")}</p></div>}
          </div>

          <div className={`drop-zone ${dstDragOver ? "drag-over" : ""} ${dstUploaded ? "has-file" : ""} ${!sessionId ? "opacity-25 pointer-events-none" : ""}`} onDragOver={(e) => { e.preventDefault(); setDstDragOver(true); }} onDragLeave={() => setDstDragOver(false)} onDrop={handleDstDrop} onClick={() => sessionId && zipInputRef.current?.click()}>
            <input ref={zipInputRef} type="file" accept=".zip" className="hidden" onChange={(e) => { if (e.target.files) uploadDst(e.target.files); e.target.value = ""; }} />
            {dstLoading ? <p className="text-sm" style={{ color: "var(--accent)" }}>{t("cb.dst.uploading")}</p>
            : dstUploaded && dstData ? <div style={{ animation: "fadeSlideIn 0.3s ease" }}><div className="flex items-center justify-center gap-2"><span style={{ color: dstData.all_matched ? "var(--accent)" : "var(--warning)", fontSize: "16px" }}>{dstData.all_matched ? "✓" : "⚠"}</span><span className="text-sm font-medium" style={{ color: dstData.all_matched ? "var(--accent)" : "var(--warning)" }}>{dstFileName}</span></div><p className="text-[11px] mt-1" style={{ color: "var(--muted)" }}>{dstData.needed_count > 0 ? <>{(dstData.needed_count - dstData.missing_programs.length)}/{dstData.needed_count} {t("cb.dst.matched")}{dstData.missing_programs.length > 0 && <span style={{ color: "var(--danger)" }}> · {dstData.missing_programs.length} {t("cb.dst.missing")}</span>}</> : <>{dstData.uploaded_count} {t("cb.dst.uploaded")}</>}</p></div>
            : <div><span className="text-[9px] font-semibold uppercase tracking-wider" style={{ color: "var(--accent)", opacity: 0.5 }}>{t("cb.step2")}</span><svg className="mx-auto mb-2.5 mt-1.5 opacity-25" width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" /><polyline points="17 8 12 3 7 8" /><line x1="12" y1="3" x2="12" y2="15" /></svg><p className="text-sm font-medium mb-0.5">{t("cb.dst.title")}</p><p className="text-[11px]" style={{ color: "var(--muted)" }}>{sessionId ? t("cb.dst.hint") : t("cb.dst.hint_disabled")}</p></div>}
          </div>
        </div>

        {/* Column Mapping Confirmation */}
        {detectData && !mappingConfirmed && (
          <div className="glass-panel p-4 sm:p-5" style={{ animation: "fadeSlideIn 0.3s ease" }}>
            <div className="flex items-center justify-between mb-3">
              <div>
                <h3 className="text-xs font-semibold">{t("cb.mapping.title")}</h3>
                <p className="text-[10px] mt-0.5" style={{ color: "var(--muted)" }}>
                  {detectData.confidence === "high" ? t("cb.mapping.auto") : t("cb.mapping.review")}
                </p>
              </div>
              {detectData.confidence !== "high" && (
                <span className="text-[9px] font-medium px-2 py-1 rounded-md" style={{ background: "rgba(245, 158, 11, 0.1)", color: "var(--warning)" }}>{t("cb.mapping.needs_review")}</span>
              )}
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-2">
              {FIELD_KEYS.map((field) => {
                const colIdx = columnMapping[field] ?? -1;
                const samples = detectData.preview_rows.slice(0, 3).map(r => colIdx >= 0 && colIdx < r.length ? r[colIdx] : null).filter(v => v !== null).map(v => String(v));
                return (
                  <div key={field} className="p-2.5 rounded-lg" style={{ background: "var(--surface)" }}>
                    <div className="flex items-center gap-2 mb-1">
                      <span className="text-[10px] font-semibold" style={{ color: "var(--accent)" }}>{t(`cb.mapping.field.${field}`)}</span>
                    </div>
                    <select
                      className="w-full text-[10px] px-2 py-1.5 rounded-lg bg-transparent mb-1"
                      style={{ border: "1px solid var(--border)", color: "var(--foreground)" }}
                      value={colIdx}
                      onChange={(e) => setColumnMapping(prev => ({ ...prev, [field]: parseInt(e.target.value) }))}
                    >
                      <option value={-1}>{t("cb.mapping.select")}</option>
                      {detectData.headers.map((h, i) => h && (
                        <option key={i} value={i}>{String.fromCharCode(65 + i)} — {h}</option>
                      ))}
                    </select>
                    <p className="text-[9px]" style={{ color: "var(--muted)" }}>{t(`cb.mapping.help.${field}`)}</p>
                    {samples.length > 0 && (
                      <p className="text-[9px] mt-0.5 font-mono truncate" style={{ color: "var(--foreground)" }}>{samples.join(", ")}</p>
                    )}
                  </div>
                );
              })}
            </div>
            <div className="flex justify-end mt-3">
              <button
                className="accent-btn text-xs"
                onClick={confirmMapping}
                disabled={excelLoading || Object.values(columnMapping).some(v => v < 0)}
              >
                {excelLoading ? t("cb.excel.parsing") : t("cb.mapping.confirm")}
              </button>
            </div>
          </div>
        )}

        {/* Excel Preview (collapsible) */}
        {parseData?.entries_preview && parseData.entries_preview.length > 0 && (
          <div style={{ animation: "fadeSlideIn 0.3s ease" }}>
            <button onClick={() => setShowExcelPreview(!showExcelPreview)} className="text-[11px] flex items-center gap-1.5 mb-1" style={{ color: "var(--accent)" }}>
              <span className="text-[9px]">{showExcelPreview ? "▼" : "▶"}</span>
              {showExcelPreview ? t("cb.excel.hide") : t("cb.excel.view")} ({parseData.entries_count} {t("cb.excel.rows")})
            </button>
            {showExcelPreview && (
              <div className="glass-panel p-3 overflow-auto custom-scroll" style={{ animation: "fadeSlideIn 0.2s ease", maxHeight: "320px" }}>
                <table className="w-full text-[10px]" style={{ fontFamily: "var(--font-geist-mono)" }}>
                  <thead className="sticky top-0" style={{ background: "var(--glass-strong)" }}>
                    <tr style={{ borderBottom: "1px solid var(--border)" }}>
                      {(["cb.table.program", "cb.table.name", "cb.table.title", "cb.table.qty", "cb.table.combo", "cb.table.machine", "cb.table.group"] as const).map((k) => (
                        <th key={k} className="text-left py-1.5 px-2 font-semibold" style={{ color: "var(--muted)" }}>{t(k)}</th>
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
                  <p className="text-[10px] mt-2 px-2" style={{ color: "var(--muted)" }}>{t("cb.excel.showing")} {parseData.entries_preview.length} {t("cb.excel.of")} {parseData.entries_count} {t("cb.excel.rows")}</p>
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
              {t("cb.settings")}
            </button>
            {showSettings && (
              <div className="glass-panel p-4 flex flex-col sm:flex-row gap-3 sm:gap-6" style={{ animation: "fadeSlideIn 0.2s ease" }}>
                <label className="flex items-center gap-2 text-[11px]">
                  <span style={{ color: "var(--muted)" }}>{t("cb.settings.vgap")}</span>
                  <input type="number" value={gapMm} onChange={(e) => setGapMm(Number(e.target.value))} min={0} max={20} step={0.5} className="w-14 text-center text-[11px] px-2 py-1 rounded-lg bg-transparent" style={{ border: "1px solid var(--border)", color: "var(--foreground)" }} />
                  <span style={{ color: "var(--muted)" }}>mm</span>
                </label>
                <label className="flex items-center gap-2 text-[11px]">
                  <span style={{ color: "var(--muted)" }}>{t("cb.settings.cgap")}</span>
                  <input type="number" value={columnGapMm} onChange={(e) => setColumnGapMm(Number(e.target.value))} min={0} max={30} step={0.5} className="w-14 text-center text-[11px] px-2 py-1 rounded-lg bg-transparent" style={{ border: "1px solid var(--border)", color: "var(--foreground)" }} />
                  <span style={{ color: "var(--muted)" }}>mm</span>
                </label>
              </div>
            )}
          </div>
        )}

        {/* Pipeline Stats */}
        {parseData && (
          <div className="grid grid-cols-2 gap-2 sm:flex sm:items-center sm:justify-center sm:gap-3" style={{ animation: "slideUp 0.4s ease 0.1s forwards", opacity: 0, overflow: "visible", paddingTop: "12px" }}>
            <StatCard value={parseData.entries_count} label={t("cb.stats.names")} delay={0.05} />
            <span className="stat-arrow hidden sm:block">→</span>
            <StatCard value={parseData.groups.length} label={t("cb.stats.groups")} delay={0.1} />
            <span className="stat-arrow hidden sm:block">→</span>
            <StatCard value={parseData.combo_count} label={t("cb.stats.output")} delay={0.15} />
            <span className="stat-arrow hidden sm:block">→</span>
            <StatCard value={parseData.total_slots} label={t("cb.stats.slots")} delay={0.2} />
          </div>
        )}

        {/* Two-Panel: Combo List + Slot Preview */}
        {parseData && parseData.groups.length > 0 && (
          <div className="flex-1 grid grid-cols-1 lg:grid-cols-[1fr_360px] gap-3 min-h-0" style={{ animation: "slideUp 0.4s ease 0.15s forwards", opacity: 0 }}>
            <div className="glass-panel overflow-hidden flex flex-col min-h-[200px] sm:min-h-[300px]">
              <div className="flex items-center justify-between px-4 py-2.5 shrink-0" style={{ borderBottom: "1px solid var(--border)" }}>
                <span className="text-xs font-semibold">{t("cb.files.title")} <span className="font-normal" style={{ color: "var(--muted)" }}>{selectedCombos.size}/{totalCombos}</span></span>
                <div className="flex gap-1"><button onClick={selectAll} className="glass-btn text-[10px]">{t("cb.files.all")}</button><button onClick={deselectAll} className="glass-btn text-[10px]">{t("cb.files.none")}</button></div>
              </div>
              <div className="flex-1 overflow-y-auto custom-scroll">
                {parseData.groups.map((group) => {
                  const gk = `${group.machine_program}_${group.com_no}`;
                  const exp = expandedGroups.has(gk);
                  const sel = group.combos.filter((c) => selectedCombos.has(c.filename)).length;
                  return (
                    <div key={gk} style={{ borderBottom: "1px solid var(--border)" }}>
                      <button className="w-full flex flex-col px-4 py-3 sm:py-2.5 text-left transition-colors" style={{ background: exp ? "var(--surface)" : "transparent" }} onClick={() => toggleGroup(gk)}>
                        <div className="flex items-center gap-2.5 w-full">
                          <span className="text-[10px] w-3" style={{ color: "var(--muted)" }}>{exp ? "▼" : "▶"}</span>
                          <span className="text-xs font-medium">{group.machine_program}<span style={{ color: "var(--muted)", fontWeight: 400 }}> / {t("Combo")} {group.com_no}</span></span>
                          <span className="text-[10px] ml-auto tabular-nums hidden sm:inline" style={{ color: "var(--muted)" }}>
                            {group.entry_count} {t("names")} → {group.combos.length} {group.combos.length === 1 ? t("file") : t("files")} ({group.total_slots} {t("slots")})
                          </span>
                          <span className="text-[10px] ml-auto tabular-nums sm:hidden" style={{ color: "var(--muted)" }}>
                            {group.combos.length}f
                          </span>
                        </div>
                      </button>
                      {exp && group.combos.map((combo, ci) => (
                        <div key={combo.filename} className="flex items-center gap-2.5 px-4 py-3 sm:py-2 cursor-pointer transition-all" style={{ background: previewCombo?.filename === combo.filename ? "var(--accent-glow)" : "var(--surface)", borderTop: "1px solid var(--border)", animation: `fadeSlideIn 0.15s ease ${ci * 0.02}s forwards`, opacity: 0 }} onClick={() => { setPreviewCombo(combo); setShowMobilePreview(true); }}>
                          <input type="checkbox" className="custom-checkbox" checked={selectedCombos.has(combo.filename)} onChange={() => toggleCombo(combo.filename)} onClick={(e) => e.stopPropagation()} />
                          <span className="text-[11px] font-mono">{combo.filename}</span>
                          <span className="text-[10px] ml-auto tabular-nums" style={{ color: "var(--muted)" }}>{combo.left_count}L{combo.right_count > 0 ? ` + ${combo.right_count}R` : ""}</span>
                        </div>
                      ))}
                    </div>
                  );
                })}
              </div>
            </div>

            {/* Desktop preview sidebar */}
            <div className="hidden lg:block glass-panel p-4 overflow-y-auto custom-scroll lg:sticky lg:top-0 lg:self-start min-h-[300px]" style={{ maxHeight: "calc(100vh - 340px)" }}>
              {previewCombo ? (
                <>
                  <h3 className="text-[11px] font-semibold font-mono mb-3 pb-2" style={{ borderBottom: "1px solid var(--border)", color: "var(--accent)" }}>{previewCombo.filename}</h3>
                  <div className="grid grid-cols-2 gap-4">
                    <div>
                      <p className="text-[9px] font-semibold uppercase tracking-wider mb-1.5" style={{ color: "var(--muted)" }}>{t("cb.preview.left")}</p>
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
                      <p className="text-[9px] font-semibold uppercase tracking-wider mb-1.5" style={{ color: "var(--muted)" }}>{t("cb.preview.right")}</p>
                      {previewCombo.slots.slice(10).map((s, i) => (
                        <div key={i} className="flex items-center gap-1 py-[2px]" style={{ borderBottom: "1px solid var(--border)" }}>
                          <span className="text-[9px] w-3.5 text-right tabular-nums" style={{ color: "var(--border-strong)" }}>{i + 11}</span>
                          <span className="text-[9px] font-mono w-6 tabular-nums" style={{ color: "var(--accent)" }}>{s.program}</span>
                          <span className="text-[9px] truncate">{s.name_line1}</span>
                        </div>
                      ))}
                      {previewCombo.right_count === 0 && <p className="text-[9px] py-2" style={{ color: "var(--border-strong)" }}>{t("cb.preview.no_right")}</p>}
                    </div>
                  </div>
                </>
              ) : (
                <div className="flex items-center justify-center h-full"><p className="text-[11px]" style={{ color: "var(--muted)" }}>{t("cb.preview.click")}</p></div>
              )}
            </div>
          </div>
        )}

        {/* Mobile preview overlay */}
        {showMobilePreview && previewCombo && (
          <div className="lg:hidden fixed inset-0 z-50 flex flex-col justify-end">
            <div className="absolute inset-0 bg-black/30" onClick={() => setShowMobilePreview(false)} />
            <div className="relative glass-panel rounded-t-2xl p-4 overflow-y-auto" style={{ maxHeight: "65vh", animation: "slideUp 0.2s ease" }}>
              <div className="flex items-center justify-between mb-3 pb-2" style={{ borderBottom: "1px solid var(--border)" }}>
                <h3 className="text-[11px] font-semibold font-mono" style={{ color: "var(--accent)" }}>{previewCombo.filename}</h3>
                <button onClick={() => setShowMobilePreview(false)} className="w-7 h-7 flex items-center justify-center rounded-lg" style={{ color: "var(--muted)", background: "var(--surface)" }}>✕</button>
              </div>
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <p className="text-[9px] font-semibold uppercase tracking-wider mb-1.5" style={{ color: "var(--muted)" }}>{t("cb.preview.left")}</p>
                  {previewCombo.slots.slice(0, 10).map((s, i) => (
                    <div key={i} className="flex items-center gap-1 py-[2px]" style={{ borderBottom: "1px solid var(--border)" }}>
                      <span className="text-[9px] w-3.5 text-right tabular-nums" style={{ color: "var(--border-strong)" }}>{i + 1}</span>
                      <span className="text-[9px] font-mono w-6 tabular-nums" style={{ color: "var(--accent)" }}>{s.program}</span>
                      <span className="text-[9px] truncate">{s.name_line1}</span>
                    </div>
                  ))}
                </div>
                <div>
                  <p className="text-[9px] font-semibold uppercase tracking-wider mb-1.5" style={{ color: "var(--muted)" }}>{t("cb.preview.right")}</p>
                  {previewCombo.slots.slice(10).map((s, i) => (
                    <div key={i} className="flex items-center gap-1 py-[2px]" style={{ borderBottom: "1px solid var(--border)" }}>
                      <span className="text-[9px] w-3.5 text-right tabular-nums" style={{ color: "var(--border-strong)" }}>{i + 11}</span>
                      <span className="text-[9px] font-mono w-6 tabular-nums" style={{ color: "var(--accent)" }}>{s.program}</span>
                      <span className="text-[9px] truncate">{s.name_line1}</span>
                    </div>
                  ))}
                  {previewCombo.right_count === 0 && <p className="text-[9px] py-2" style={{ color: "var(--border-strong)" }}>{t("cb.preview.no_right")}</p>}
                </div>
              </div>
            </div>
          </div>
        )}
      </main>

      {/* Export Bar */}
      {parseData && (
        <div className="shrink-0 px-4 sm:px-6 py-3" style={{ background: "var(--glass-strong)", backdropFilter: "blur(24px)", borderTop: "1px solid var(--glass-border)", animation: "slideUp 0.3s ease 0.2s forwards", opacity: 0 }}>
          <div className="max-w-6xl mx-auto flex flex-col sm:flex-row items-stretch sm:items-center gap-2 sm:gap-4">
            <div className="flex-1 min-w-0">
              {exporting && (
                <div>
                  <div className="flex items-center gap-2.5 mb-1.5">
                    <div className="relative w-5 h-5 shrink-0">
                      <div className="absolute inset-0 rounded-full" style={{ border: "2px solid var(--border)" }} />
                      <div className="absolute inset-0 rounded-full animate-spin" style={{ border: "2px solid transparent", borderTopColor: "var(--accent)" }} />
                    </div>
                    <p className="text-xs font-medium" style={{ color: "var(--foreground)" }}>
                      {exportProgress === -1 ? t("cb.export.downloading") : `${t("cb.export.exporting")}...`}
                    </p>
                    {exportProgress > 0 && exportProgress !== -1 && (
                      <span className="text-[11px] tabular-nums font-mono" style={{ color: "var(--accent)" }}>{exportProgress}{t("cb.export.elapsed")}</span>
                    )}
                  </div>
                  <div className="progress-bar" style={{ height: "6px", overflow: "hidden", borderRadius: "3px" }}>
                    <div style={{ width: "100%", height: "100%", background: `linear-gradient(90deg, var(--accent) 0%, var(--accent-light) 50%, var(--accent) 100%)`, backgroundSize: "200% 100%", animation: "shimmer 1.5s ease-in-out infinite", opacity: 0.6 }} />
                  </div>
                  <div className="flex items-center justify-between mt-1">
                    <p className="text-[10px]" style={{ color: "var(--muted)" }}>{t("cb.export.progress")}</p>
                    <p className="text-[10px]" style={{ color: "var(--muted)" }}>
                      {(() => { const avg = parseData ? parseData.total_slots / Math.max(1, parseData.combo_count) : 10; const spc = avg <= 10 ? 5 : 8; return `${t("cb.export.estimate")} ~${Math.max(1, Math.ceil(selectedCombos.size * spc / 60))} min`; })()}
                    </p>
                  </div>
                </div>
              )}
              {downloadUrl && !exporting && (
                <div className="flex items-center gap-2 flex-wrap">
                  <p className="text-[11px]" style={{ color: "var(--success)" }}>✓ {t("cb.export.done")} combos_{sessionName.replace(/\s+/g, "_")}.zip</p>
                  {exportElapsed > 0 && (
                    <span className="text-[10px]" style={{ color: "var(--muted)" }}>— {t("cb.export.completed_in")} {exportElapsed >= 60 ? `${Math.floor(exportElapsed / 60)}m ${exportElapsed % 60}s` : `${exportElapsed}s`}</span>
                  )}
                  <a href={downloadUrl} download={`combos_${sessionName.replace(/\s+/g, "_")}.zip`} className="text-[11px] underline" style={{ color: "var(--accent)" }}>{t("cb.export.again")}</a>
                </div>
              )}
              {!exporting && !downloadUrl && exported && (
                <p className="text-[11px]" style={{ color: "var(--success)" }}>✓ {t("cb.export.previous")}</p>
              )}
              {!exporting && !downloadUrl && !exported && !dstUploaded && <p className="text-[11px]" style={{ color: "var(--muted)" }}>{t("cb.export.need_dst")}</p>}
            </div>
            <div className="flex flex-col items-stretch sm:items-end gap-1">
              <span className="text-[9px] font-semibold uppercase tracking-wider text-center sm:text-right" style={{ color: "var(--accent)", opacity: 0.5 }}>{t("cb.step3")}</span>
              <button className="accent-btn w-full sm:w-auto" disabled={selectedCombos.size === 0 || !dstUploaded || exporting} onClick={handleExport}>
                {exporting ? t("cb.export.exporting") : `${t("cb.export.btn")} ${selectedCombos.size} ${selectedCombos.size !== 1 ? t("cb.export.files") : t("cb.export.file")}`}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
