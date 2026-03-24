"use client";

import Link from "next/link";
import Image from "next/image";
import { useState, useCallback, useRef, useEffect, useMemo } from "react";
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
const REQUIRED_FIELDS = ["program", "name_line1", "quantity", "com_no", "machine_program"] as const;
const ASSIGN_FIELD_KEYS = ["size", "fabric_colour", "frame_colour", "embroidery_colour"] as const;
interface DstResponse { session_id: string; uploaded_count: number; needed_count: number; missing_programs: number[]; all_matched: boolean; }
interface SessionSummary { session_id: string; name: string; created_at: string; updated_at: string; has_excel: boolean; entries_count: number; combo_count: number; dst_count: number; exported: boolean; }
interface FullSession {
  session_id: string; name: string; entries_count: number; total_slots: number;
  groups: Group[]; combo_count: number; entries_preview?: EntryPreview[];
  dst_count: number; dst_all_matched: boolean; missing_programs: number[];
  exported: boolean; exported_at: string | null; excel_filename: string | null;
  gap_mm: number; column_gap_mm: number; warnings: string[];
}

// Per-field color palette
const FIELD_COLORS: Record<string, { bg: string; text: string; border: string; cell: string; glow: string }> = {
  program:         { bg: "rgba(59,130,246,0.13)",  text: "#3b82f6", border: "rgba(59,130,246,0.4)",  cell: "rgba(59,130,246,0.06)",  glow: "rgba(59,130,246,0.2)"  },
  name_line1:      { bg: "rgba(34,197,94,0.13)",   text: "#16a34a", border: "rgba(34,197,94,0.4)",   cell: "rgba(34,197,94,0.06)",   glow: "rgba(34,197,94,0.2)"   },
  name_line2:      { bg: "rgba(20,184,166,0.13)",  text: "#0d9488", border: "rgba(20,184,166,0.4)",  cell: "rgba(20,184,166,0.06)",  glow: "rgba(20,184,166,0.2)"  },
  quantity:        { bg: "rgba(168,85,247,0.13)",  text: "#a855f7", border: "rgba(168,85,247,0.4)",  cell: "rgba(168,85,247,0.06)",  glow: "rgba(168,85,247,0.2)"  },
  com_no:          { bg: "rgba(249,115,22,0.13)",  text: "#f97316", border: "rgba(249,115,22,0.4)",  cell: "rgba(249,115,22,0.06)",  glow: "rgba(249,115,22,0.2)"  },
  machine_program: { bg: "rgba(239,68,68,0.13)",   text: "#ef4444", border: "rgba(239,68,68,0.4)",   cell: "rgba(239,68,68,0.06)",   glow: "rgba(239,68,68,0.2)"   },
};

const ASSIGN_FIELD_COLORS: Record<string, { bg: string; text: string; border: string; cell: string; glow: string }> = {
  size:              { bg: "rgba(245,158,11,0.13)",  text: "#f59e0b", border: "rgba(245,158,11,0.4)",  cell: "rgba(245,158,11,0.06)",  glow: "rgba(245,158,11,0.2)"  },
  fabric_colour:     { bg: "rgba(59,130,246,0.13)",  text: "#3b82f6", border: "rgba(59,130,246,0.4)",  cell: "rgba(59,130,246,0.06)",  glow: "rgba(59,130,246,0.2)"  },
  frame_colour:      { bg: "rgba(168,85,247,0.13)",  text: "#a855f7", border: "rgba(168,85,247,0.4)",  cell: "rgba(168,85,247,0.06)",  glow: "rgba(168,85,247,0.2)"  },
  embroidery_colour: { bg: "rgba(34,197,94,0.13)",   text: "#16a34a", border: "rgba(34,197,94,0.4)",   cell: "rgba(34,197,94,0.06)",   glow: "rgba(34,197,94,0.2)"   },
};

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

function StatCard({ value, label, description, delay = 0, icon }: { value: number; label: string; description?: string; delay?: number; icon?: React.ReactNode }) {
  const display = useCountUp(value);
  return (
    <div className="stat-card group relative" style={{ animation: `countUp 0.4s ease ${delay}s forwards`, opacity: 0 }}>
      {icon && <div className="flex justify-center mb-1.5 opacity-35">{icon}</div>}
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

// Visual step stepper
function StepIndicator({ steps, labels }: { steps: {done: boolean; active: boolean}[]; labels: string[] }) {
  return (
    <div className="flex items-center justify-center py-1" style={{ animation: "fadeIn 0.5s ease 0.1s forwards", opacity: 0 }}>
      {steps.map((s, i) => (
        <div key={i} className="flex items-center">
          {i > 0 && (
            <div style={{
              width: "36px", height: "2px", margin: "0 4px",
              background: steps[i - 1].done ? "var(--accent)" : "var(--border)",
              borderRadius: "1px",
              transition: "background 0.4s ease",
            }} />
          )}
          <div className="flex flex-col items-center gap-1">
            <div
              className={s.active ? "step-active-glow" : ""}
              style={{
                width: "26px", height: "26px", borderRadius: "50%",
                display: "flex", alignItems: "center", justifyContent: "center",
                fontSize: "10px", fontWeight: 700,
                background: s.done ? "var(--accent)" : s.active ? "transparent" : "var(--surface)",
                border: s.done ? "none" : s.active ? "2px solid var(--accent)" : "2px solid var(--border)",
                color: s.done ? "white" : s.active ? "var(--accent)" : "var(--muted)",
                transition: "all 0.35s cubic-bezier(.4,0,.2,1)",
              }}
            >
              {s.done ? (
                <svg width="11" height="11" viewBox="0 0 12 12" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                  <polyline points="2,6 5,9 10,3" />
                </svg>
              ) : i + 1}
            </div>
            <span style={{
              fontSize: "9px", whiteSpace: "nowrap", letterSpacing: "0.03em",
              color: s.done ? "var(--accent)" : s.active ? "var(--foreground)" : "var(--muted)",
              fontWeight: s.active ? 500 : 400,
              transition: "color 0.3s ease",
            }}>{labels[i]}</span>
          </div>
        </div>
      ))}
    </div>
  );
}

// Skeleton shimmer bars for Excel parsing
function ExcelSkeleton({ label }: { label: string }) {
  return (
    <div className="flex flex-col items-center gap-2.5 w-full px-6 max-w-xs mx-auto">
      <div className="flex gap-1.5 w-full">
        <div className="skeleton" style={{ height: "8px", borderRadius: "4px", flex: "0 0 40%" }} />
        <div className="skeleton" style={{ height: "8px", borderRadius: "4px", flex: "0 0 30%" }} />
        <div className="skeleton" style={{ height: "8px", borderRadius: "4px", flex: "0 0 20%" }} />
      </div>
      <div className="flex gap-1.5 w-full">
        <div className="skeleton" style={{ height: "8px", borderRadius: "4px", flex: "0 0 30%" }} />
        <div className="skeleton" style={{ height: "8px", borderRadius: "4px", flex: "0 0 45%" }} />
        <div className="skeleton" style={{ height: "8px", borderRadius: "4px", flex: "0 0 15%" }} />
      </div>
      <div className="flex gap-1.5 w-full">
        <div className="skeleton" style={{ height: "8px", borderRadius: "4px", flex: "0 0 50%" }} />
        <div className="skeleton" style={{ height: "8px", borderRadius: "4px", flex: "0 0 25%" }} />
        <div className="skeleton" style={{ height: "8px", borderRadius: "4px", flex: "0 0 15%" }} />
      </div>
      <p className="text-[10px] mt-0.5" style={{ color: "var(--accent)" }}>{label}</p>
    </div>
  );
}

export default function EmbroideryStacker() {
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
  // New: interactive column assignment
  const [activeField, setActiveField] = useState<string | null>(null);
  const [showHowItWorks, setShowHowItWorks] = useState(true);
  // Auto-assign MA/COM state
  const [assignMode, setAssignMode] = useState<"pending" | "detect" | "result" | "skipped">("pending");
  const [assignDetectData, setAssignDetectData] = useState<{headers: string[]; preview_rows: (string|number|null)[][]; detected_mapping: Record<string, number>; confidence: string} | null>(null);
  const [assignColumnMapping, setAssignColumnMapping] = useState<Record<string, number>>({});
  const [assignActiveField, setAssignActiveField] = useState<string | null>(null);
  const [assignResult, setAssignResult] = useState<{assignments_count: number; ma_summary: {ma: string; size: string; count: number}[]; com_summary: {ma: string; com: number; fabric_colour: string; frame_colour: string; embroidery_colour: string; count: number}[]; warnings: string[]} | null>(null);
  const [assignLoading, setAssignLoading] = useState(false);
  const [assignDownloading, setAssignDownloading] = useState(false);
  const [backendStatus, setBackendStatus] = useState<"connecting" | "ready" | "failed">("connecting");
  const assignExcelInputRef = useRef<HTMLInputElement>(null);
  const excelInputRef = useRef<HTMLInputElement>(null);
  const zipInputRef = useRef<HTMLInputElement>(null);

  const showToast = useCallback((message: string, type: "error" | "warning" | "success" = "error") => setToast({ message, type }), []);

  // Auto-expand "how it works" when confidence is low
  useEffect(() => {
    if (detectData) setShowHowItWorks(true);
  }, [detectData]);

  // Close active field on Escape
  useEffect(() => {
    const handler = (e: KeyboardEvent) => { if (e.key === "Escape") setActiveField(null); };
    window.addEventListener("keydown", handler);
    return () => window.removeEventListener("keydown", handler);
  }, []);

  // Per-field header pattern detection: "auto" vs "guessed"
  const HEADER_PATTERNS: Record<string, RegExp> = {
    program: /program|prog|โปรแกรม/i,
    name_line1: /name.*1|first.*name|ชื่อ.*1|name line/i,
    name_line2: /name.*2|last.*name|ชื่อ.*2|surname/i,
    quantity: /qty|quantity|amount|count|pcs|จำนวน/i,
    com_no: /com.*no|combo|คอมโบ/i,
    machine_program: /machine|^m$|ma|เครื่อง/i,
  };

  const fieldDetectionSource = useMemo(() => {
    if (!detectData) return {};
    const sources: Record<string, "auto" | "guessed"> = {};
    for (const field of FIELD_KEYS) {
      const colIdx = columnMapping[field];
      if (colIdx === undefined || colIdx < 0) continue;
      const header = detectData.headers[colIdx];
      if (header && HEADER_PATTERNS[field]?.test(header)) {
        sources[field] = "auto";
      } else {
        sources[field] = "guessed";
      }
    }
    return sources;
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [detectData, columnMapping]);

  // Proactive validation warnings on mapped columns
  const mappingWarnings = useMemo(() => {
    if (!detectData) return {};
    const warnings: Record<string, string> = {};
    const rows = detectData.preview_rows;

    const qtyCol = columnMapping.quantity;
    if (qtyCol !== undefined && qtyCol >= 0) {
      const hasHigh = rows.some(r => qtyCol < r.length && Number(r[qtyCol]) > 50);
      if (hasHigh) warnings.quantity = t("cb.mapping.warn.qty_high");
    }

    const progCol = columnMapping.program;
    if (progCol !== undefined && progCol >= 0) {
      const hasNonNum = rows.some(r => progCol < r.length && r[progCol] !== null && r[progCol] !== "" && isNaN(Number(r[progCol])));
      if (hasNonNum) warnings.program = t("cb.mapping.warn.program_format");
    }

    const comCol = columnMapping.com_no;
    if (comCol !== undefined && comCol >= 0) {
      const hasAlpha = rows.some(r => comCol < r.length && r[comCol] !== null && /[a-zA-Z]{3,}/.test(String(r[comCol])));
      if (hasAlpha) warnings.com_no = t("cb.mapping.warn.combo_format");
    }

    const maCol = columnMapping.machine_program;
    if (maCol !== undefined && maCol >= 0) {
      const hasBadMA = rows.some(r => maCol < r.length && r[maCol] !== null && r[maCol] !== "" && !/^MA/i.test(String(r[maCol])));
      if (hasBadMA) warnings.machine_program = t("cb.mapping.warn.ma_format");
    }

    return warnings;
  }, [detectData, columnMapping, t]);

  const fetchSessions = useCallback(async () => {
    try {
      const res = await authFetch(`${API}/api/sessions`);
      if (res.ok) setSavedSessions(await res.json());
    } catch { /* ignore */ }
  }, []);

  useEffect(() => {
    fetchSessions();
    warmupBackend(API, setBackendStatus);
  }, [fetchSessions]);

  // Clean up session when user leaves the page
  useEffect(() => {
    const cleanup = () => {
      if (sessionId) {
        fetch(`${API}/api/session/${sessionId}`, { method: "DELETE", keepalive: true });
      }
    };
    window.addEventListener("beforeunload", cleanup);
    return () => window.removeEventListener("beforeunload", cleanup);
  }, [sessionId]);

  const saveSessionName = useCallback(async (sid: string, name: string) => {
    const form = new FormData(); form.append("session_id", sid); form.append("name", name);
    try { await authFetch(`${API}/api/session/name`, { method: "POST", body: form }); } catch { /* ignore */ }
  }, []);

  const resetSession = useCallback(() => {
    setDetectData(null); setColumnMapping({}); setMappingConfirmed(false);
    setDstUploaded(false); setDstData(null); setDstFileName("");
    setSelectedCombos(new Set()); setExpandedGroups(new Set());
    setPreviewCombo(null); setDownloadUrl(""); setExportProgress(0);
    setShowExcelPreview(false); setExported(false); setActiveField(null);
    setAssignMode("pending"); setAssignDetectData(null); setAssignColumnMapping({}); setAssignResult(null); setAssignActiveField(null);
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
    setGapMm(data.gap_mm);
    setColumnGapMm(data.column_gap_mm);
    if (data.excel_filename) { setExcelFile(data.excel_filename); } else { setExcelFile(""); }
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
      setParseData(null); setSelectedCombos(new Set()); setExpandedGroups(new Set()); setPreviewCombo(null);
    }
    if (data.dst_count > 0) {
      setDstUploaded(true);
      setDstFileName(`${data.dst_count} DST files`);
      setDstData({ session_id: data.session_id, uploaded_count: data.dst_count, needed_count: data.entries_count, missing_programs: data.missing_programs || [], all_matched: data.dst_all_matched });
    } else {
      setDstUploaded(false); setDstData(null); setDstFileName("");
    }
    setDownloadUrl(""); setExportProgress(0); setShowExcelPreview(false); setShowSettings(false);
    setMappingConfirmed(data.entries_count > 0);
  }, []);

  const loadFullSession = useCallback(async (sid: string) => {
    setLoadingSession(true);
    try {
      const res = await authFetch(`${API}/api/session/${sid}/full`);
      if (!res.ok) throw new Error(`${t("err.load_session")}: ${res.status}`);
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
      if (!res.ok) throw new Error(t("err.delete_fail"));
      showToast(t("ok.session_deleted"), "success");
      fetchSessions();
      if (sid === sessionId) {
        resetSession(); setExcelFile(""); setParseData(null); setSessionStarted(false); setSessionId("");
      }
    } catch (e) { showToast(`${e instanceof Error ? e.message : t("err.delete_session")}`); }
    setDeleteConfirm(null);
  }, [sessionId, resetSession, showToast, fetchSessions, t]);

  const removeExcel = useCallback(async () => {
    if (!sessionId) return;
    try {
      const res = await authFetch(`${API}/api/session/${sessionId}/excel`, { method: "DELETE" });
      if (!res.ok) throw new Error(t("err.remove_excel_fail"));
      setExcelFile(""); setParseData(null); setSelectedCombos(new Set()); setExpandedGroups(new Set());
      setPreviewCombo(null); setDownloadUrl(""); setExportProgress(0); setShowExcelPreview(false);
      setExported(false); setDetectData(null); setColumnMapping({}); setMappingConfirmed(false);
      showToast(t("ok.excel_removed"), "success");
      fetchSessions();
    } catch (e) { showToast(`${e instanceof Error ? e.message : t("err.excel_read")}`); }
  }, [sessionId, showToast, fetchSessions, t]);

  const uploadExcel = useCallback(async (file: File) => {
    if (!file.name.match(/\.(xlsx|xls)$/i)) { showToast(t("err.excel_format")); return; }
    resetSession();
    setSessionId("");
    setExcelLoading(true); setExcelFile(file.name);
    const form = new FormData(); form.append("file", file);
    try {
      const res = await authFetch(`${API}/api/detect-columns`, { method: "POST", body: form });
      if (!res.ok) throw new Error(`${t("err.server")}: ${res.status}`);
      const data: DetectResponse = await res.json();
      setDetectData(data);
      setColumnMapping(data.detected_mapping);
      setSessionId(data.session_id);
      saveSessionName(data.session_id, sessionName);
      fetchSessions();
    } catch (e) { showToast(`${t("err.excel_read")}: ${e instanceof Error ? e.message : t("err.connection")}`); setExcelFile(""); }
    setExcelLoading(false);
  }, [sessionId, resetSession, showToast, sessionName, saveSessionName, fetchSessions, t]);

  const uploadAssignExcel = useCallback(async (file: File) => {
    if (!file.name.match(/\.(xlsx|xls)$/i)) { showToast(t("err.excel_format")); return; }
    setAssignLoading(true); setAssignResult(null); setExcelFile(file.name);
    const form = new FormData(); form.append("file", file);
    if (sessionId) form.append("session_id", sessionId);
    try {
      const res = await authFetch(`${API}/api/detect-assign-columns`, { method: "POST", body: form });
      if (!res.ok) throw new Error(`${t("err.server")}: ${res.status}`);
      const data = await res.json();
      setAssignDetectData(data);
      setAssignColumnMapping(data.detected_mapping);
      setSessionId(data.session_id);
      saveSessionName(data.session_id, sessionName);
      setAssignMode("detect");
    } catch (e) { showToast(`${t("err.excel_read")}: ${e instanceof Error ? e.message : t("err.connection")}`); setExcelFile(""); }
    setAssignLoading(false);
  }, [sessionId, showToast, sessionName, saveSessionName, t]);

  const runAutoAssign = useCallback(async () => {
    if (!sessionId) return;
    setAssignLoading(true);
    const form = new FormData();
    form.append("session_id", sessionId);
    form.append("column_map", JSON.stringify(assignColumnMapping));
    try {
      const res = await authFetch(`${API}/api/auto-assign`, { method: "POST", body: form });
      if (!res.ok) throw new Error(`${t("err.server")}: ${res.status}`);
      const data = await res.json();
      setAssignResult(data);
      setAssignMode("result");
      if (data.warnings?.length) showToast(`${data.warnings.length} ${t("ok.warnings")}`, "warning");
    } catch (e) { showToast(`${t("err.parse_fail")}: ${e instanceof Error ? e.message : t("err.connection")}`); }
    setAssignLoading(false);
  }, [sessionId, assignColumnMapping, showToast, t]);

  const downloadAssignedExcel = useCallback(async () => {
    if (!sessionId) return;
    setAssignDownloading(true);
    try {
      const form = new FormData(); form.append("session_id", sessionId);
      const res = await authFetch(`${API}/api/download-assigned-excel`, { method: "POST", body: form });
      if (!res.ok) throw new Error(`${t("err.server")}: ${res.status}`);
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a"); a.href = url;
      const disposition = res.headers.get("Content-Disposition");
      const filename = disposition?.match(/filename="(.+)"/)?.[1] || "order_with_MA_COM.xlsx";
      a.download = filename; a.click(); URL.revokeObjectURL(url);
    } catch (e) { showToast(`${t("err.export_fail")}: ${e instanceof Error ? e.message : t("err.connection")}`); }
    setAssignDownloading(false);
  }, [sessionId, showToast, t]);

  const proceedFromAssign = useCallback(async () => {
    if (!sessionId) return;
    setAssignLoading(true);
    try {
      const form = new FormData();
      form.append("session_id", sessionId);
      const res = await authFetch(`${API}/api/apply-assignments`, { method: "POST", body: form });
      if (!res.ok) throw new Error(`${t("err.server")}: ${res.status}`);
      const data = await res.json();
      if (data.entries_count === 0) { showToast(t("err.no_entries"), "warning"); }
      else {
        setMappingConfirmed(true);
        setAssignMode("skipped");
        applyParseData(data);
      }
    } catch (e) { showToast(`${t("err.parse_fail")}: ${e instanceof Error ? e.message : t("err.connection")}`); }
    setAssignLoading(false);
  }, [sessionId, applyParseData, showToast, t]);

  const handleAssignColumnClick = useCallback((colIdx: number) => {
    if (assignActiveField) {
      const cleaned = { ...assignColumnMapping };
      Object.entries(cleaned).forEach(([field, idx]) => {
        if (idx === colIdx && field !== assignActiveField) cleaned[field] = -1;
      });
      if (cleaned[assignActiveField] === colIdx) { cleaned[assignActiveField] = -1; }
      else { cleaned[assignActiveField] = colIdx; }
      setAssignColumnMapping(cleaned);
      setAssignActiveField(null);
    } else {
      const ownerField = Object.entries(assignColumnMapping).find(([, idx]) => idx === colIdx)?.[0];
      if (ownerField) setAssignActiveField(ownerField);
    }
  }, [assignActiveField, assignColumnMapping]);

  const confirmMapping = useCallback(async () => {
    if (!sessionId || !detectData) return;
    setExcelLoading(true);
    const form = new FormData();
    form.append("session_id", sessionId);
    form.append("column_map", JSON.stringify(columnMapping));
    try {
      const res = await authFetch(`${API}/api/parse-excel`, { method: "POST", body: form });
      if (!res.ok) throw new Error(`${t("err.server")}: ${res.status}`);
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
      if (!res.ok) throw new Error(`${t("err.server")}: ${res.status}`);
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
    const startTime = Date.now();
    const elapsedTimer = setInterval(() => {
      setExportProgress(Math.round((Date.now() - startTime) / 1000));
    }, 1000);
    try {
      const res = await authFetch(`${API}/api/export`, { method: "POST", body: form }, 300_000);
      clearInterval(elapsedTimer);
      setExportElapsed(Math.round((Date.now() - startTime) / 1000));
      if (!res.ok) throw new Error(await res.text().catch(() => t("err.export_fail")));
      setExportProgress(-1);
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      setDownloadUrl(url); setExportProgress(-2);
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
      if (!res.ok) throw new Error(t("err.api_not_running"));
      const data = await res.json();
      setExcelFile("nameorder_04032026-2 add column.xlsx");
      setMappingConfirmed(true);
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

  // Interactive column assignment
  const handleColumnClick = useCallback((colIdx: number) => {
    if (activeField) {
      // Deduplicate: remove this colIdx from any other field
      const cleanedMapping = { ...columnMapping };
      Object.entries(cleanedMapping).forEach(([field, idx]) => {
        if (idx === colIdx && field !== activeField) cleanedMapping[field] = -1;
      });
      // Toggle: clicking same col that activeField already has → unassign
      if (cleanedMapping[activeField] === colIdx) {
        cleanedMapping[activeField] = -1;
      } else {
        cleanedMapping[activeField] = colIdx;
      }
      setColumnMapping(cleanedMapping);
      setActiveField(null);
    } else {
      // Select the field that owns this column
      const ownerField = Object.entries(columnMapping).find(([, idx]) => idx === colIdx)?.[0];
      if (ownerField) setActiveField(ownerField);
    }
  }, [activeField, columnMapping]);

  const toggleCombo = (f: string) => setSelectedCombos((p) => { const n = new Set(p); n.has(f) ? n.delete(f) : n.add(f); return n; });
  const selectAll = () => { if (!parseData) return; const a = new Set<string>(); parseData.groups.forEach((g) => g.combos.forEach((c) => a.add(c.filename))); setSelectedCombos(a); };
  const deselectAll = () => setSelectedCombos(new Set());
  const toggleGroup = (k: string) => setExpandedGroups((p) => { const n = new Set(p); n.has(k) ? n.delete(k) : n.add(k); return n; });
  const totalCombos = parseData?.groups.reduce((s, g) => s + g.combos.length, 0) ?? 0;

  // Build reverse mapping: colIdx → fieldKey
  const colToField = Object.entries(columnMapping)
    .filter(([, v]) => v >= 0)
    .reduce((acc, [field, col]) => { acc[col] = field; return acc; }, {} as Record<number, string>);

  const retryWarmup = useCallback(() => {
    setBackendStatus("connecting");
    warmupBackend(API, setBackendStatus);
  }, []);

  const startSession = () => { if (!sessionName.trim()) return; setSessionStarted(true); };
  const startDemoSession = () => { setSessionName("Demo Session"); setSessionStarted(true); setTimeout(() => loadSampleData(), 100); };

  const connectingBanner = backendStatus !== "ready" ? (
    <div className="w-full text-center text-xs py-2 px-4" style={{ background: backendStatus === "failed" ? "var(--danger)" : "var(--accent)", color: "white" }}>
      {backendStatus === "connecting" && <>{t("cb.connecting") || "Connecting to server..."}</>}
      {backendStatus === "failed" && (
        <>{t("cb.connect_failed") || "Could not reach server."} <button onClick={retryWarmup} className="underline ml-1 font-medium">{t("cb.retry") || "Retry"}</button></>
      )}
    </div>
  ) : null;

  /* ── Session Picker Screen ── */
  if (!sessionStarted) {
    return (
      <div className="min-h-screen flex flex-col" style={{ background: "var(--background)" }}>
        {toast && <Toast message={toast.message} type={toast.type} onClose={() => setToast(null)} />}
        {connectingBanner}

        {/* Nav */}
        <nav className="flex items-center gap-4 px-6 py-3.5 shrink-0" style={{ borderBottom: "1px solid var(--border)" }}>
          <Link href="/"><Image src="/micro-logo.svg" alt="Micro" width={56} height={16} className="micro-logo opacity-40 hover:opacity-70 transition-opacity" /></Link>
          <div className="flex-1" />
          <button onClick={toggleLang} className="nav-btn">{lang === "en" ? "TH" : "EN"}</button>
          <button onClick={toggle} className="theme-toggle">{theme === "light" ? "☾" : "☀"}</button>
          <div className="hidden sm:flex items-center gap-1.5">
            <Image src="/ossia-mark.svg?v3" alt="" width={24} height={17} className="micro-logo" />
            <span className="text-sm font-normal" style={{ color: "var(--foreground)", letterSpacing: "-0.035em" }}>ossia</span>
          </div>
        </nav>

        <div className="flex-1 flex flex-col items-center px-4 sm:px-6 py-6 sm:py-10 overflow-y-auto">
          <div className="max-w-lg w-full" style={{ animation: "slideUp 0.5s ease" }}>
            <h1 className="text-2xl sm:text-3xl font-medium tracking-tight mb-2" style={{ letterSpacing: "-0.03em" }}>{t("cb.title")}</h1>
            <p className="text-sm mb-8" style={{ color: "var(--muted)" }}>{t("cb.session.subtitle")}</p>

            {/* New Session */}
            <div className="glass-panel p-6 mb-4">
              <label className="text-[11px] font-medium block mb-2" style={{ color: "var(--muted)" }}>{t("cb.session.name")}</label>
              <input
                type="text"
                value={sessionName}
                onChange={(e) => setSessionName(e.target.value)}
                onKeyDown={(e) => e.key === "Enter" && startSession()}
                placeholder={t("cb.session.date_placeholder")}
                className="w-full text-sm px-4 py-2.5 rounded-xl bg-transparent mb-4"
                style={{ border: "1px solid var(--border)", color: "var(--foreground)" }}
                autoFocus
              />
              <button className="accent-btn w-full" onClick={startSession} disabled={!sessionName.trim()}>
                {t("cb.session.start")}
              </button>
              <button
                className="w-full mt-2 text-[11px] py-2"
                style={{ color: "var(--muted)" }}
                onClick={startDemoSession}
              >
                {t("cb.session.demo")}
              </button>
            </div>

            {/* Existing Sessions */}
            {savedSessions.length > 0 && (
              <div>
                <h2 className="text-xs font-medium mb-3" style={{ color: "var(--muted)" }}>{t("cb.session.previous")}</h2>
                <div className="flex flex-col gap-2">
                  {savedSessions.map((s) => {
                    const dotColor = s.exported ? "var(--success)" : s.has_excel ? "var(--accent)" : "var(--border-strong)";
                    return (
                      <div
                        key={s.session_id}
                        className="glass-panel session-card px-4 py-3 flex items-center gap-3 cursor-pointer transition-all hover:scale-[1.01]"
                        style={{ animation: "fadeSlideIn 0.3s ease" }}
                        onClick={() => loadFullSession(s.session_id)}
                      >
                        {/* Status dot */}
                        <div className="w-2 h-2 rounded-full shrink-0 transition-colors" style={{ background: dotColor, boxShadow: s.exported ? `0 0 5px ${dotColor}` : "none" }} />
                        <div className="flex-1 min-w-0">
                          <div className="flex items-center gap-2">
                            <span className="text-sm font-medium truncate">{s.name}</span>
                            {s.exported && (
                              <span className="text-[9px] font-medium px-1.5 py-0.5 rounded-md shrink-0" style={{ background: "var(--success)", color: "white" }}>{t("cb.exported")}</span>
                            )}
                          </div>
                          <div className="flex items-center gap-3 mt-0.5">
                            <span className="text-[10px]" style={{ color: "var(--muted)" }}>{formatDate(s.created_at)}</span>
                            {s.entries_count > 0 && <span className="text-[10px]" style={{ color: "var(--muted)" }}>{s.entries_count} {t("names")}</span>}
                            {s.combo_count > 0 && <span className="text-[10px]" style={{ color: "var(--muted)" }}>{s.combo_count} {t("combos")}</span>}
                          </div>
                        </div>
                        {/* Delete button */}
                        <button
                          className="text-xs h-6 flex items-center justify-center rounded-lg shrink-0"
                          style={{
                            color: deleteConfirm === s.session_id ? "white" : "var(--muted)",
                            background: deleteConfirm === s.session_id ? "var(--danger)" : "transparent",
                            minWidth: deleteConfirm === s.session_id ? "70px" : "24px",
                            padding: deleteConfirm === s.session_id ? "0 8px" : "0",
                            transition: "all 0.2s ease",
                            fontSize: deleteConfirm === s.session_id ? "10px" : "12px",
                            fontWeight: deleteConfirm === s.session_id ? 500 : 400,
                          }}
                          onClick={(e) => {
                            e.stopPropagation();
                            if (deleteConfirm === s.session_id) { deleteSession(s.session_id); }
                            else { setDeleteConfirm(s.session_id); setTimeout(() => setDeleteConfirm(null), 3000); }
                          }}
                          title={deleteConfirm === s.session_id ? t("cb.session.delete_confirm") : t("cb.session.delete")}
                          aria-label={deleteConfirm === s.session_id ? t("cb.session.delete_confirm") : t("cb.session.delete")}
                        >
                          {deleteConfirm === s.session_id ? t("cb.session.delete_confirm_label") : "✕"}
                        </button>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}

            {loadingSession && (
              <div className="flex items-center justify-center py-8">
                <div className="flex items-center gap-2.5">
                  <div className="spinner" />
                  <p className="text-sm" style={{ color: "var(--accent)" }}>{t("cb.session.loading")}</p>
                </div>
              </div>
            )}
          </div>
        </div>
      </div>
    );
  }

  /* ── Main Workflow ── */
  const step0Done = assignMode === "skipped" || assignMode === "result";
  const step1Done = !!excelFile && mappingConfirmed;
  const step2Done = dstUploaded;
  const step3Done = exported;

  return (
    <div className="min-h-screen flex flex-col" style={{ background: "var(--background)" }}>
      {toast && <Toast message={toast.message} type={toast.type} onClose={() => setToast(null)} />}
      {connectingBanner}

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
              <div className="fixed inset-0 z-40" onClick={() => setShowSessionDropdown(false)} />
              <div className="absolute top-full left-0 mt-1.5 py-1.5 min-w-[280px] max-w-[calc(100vw-2rem)] z-50 rounded-xl overflow-hidden" style={{ background: "var(--glass-strong)", backdropFilter: "blur(24px)", border: "1px solid var(--glass-border)", boxShadow: "var(--shadow-float)", animation: "fadeSlideIn 0.15s ease" }}>
                {savedSessions.length > 0 && savedSessions.map((s) => {
                  const dotColor = s.exported ? "var(--success)" : s.has_excel ? "var(--accent)" : "var(--border-strong)";
                  return (
                    <button
                      key={s.session_id}
                      className="w-full text-left px-3.5 py-2 text-[11px] transition-colors flex items-center gap-2.5 hover:bg-[var(--surface-hover)]"
                      style={{ background: s.session_id === sessionId ? "var(--accent-glow)" : "transparent", color: s.session_id === sessionId ? "var(--accent)" : "var(--foreground)" }}
                      onClick={async () => { setShowSessionDropdown(false); if (s.session_id === sessionId) return; await loadFullSession(s.session_id); }}
                    >
                      <div className="w-1.5 h-1.5 rounded-full shrink-0" style={{ background: dotColor }} />
                      <div className="flex-1 flex items-center gap-1.5 min-w-0">
                        <span className="font-medium truncate">{s.name}</span>
                        {s.exported && <span className="text-[8px] px-1 py-0.5 rounded shrink-0" style={{ background: "var(--success)", color: "white" }}>{t("cb.exported")}</span>}
                      </div>
                      <span className="text-[10px] shrink-0" style={{ color: "var(--muted)" }}>
                        {s.entries_count > 0 ? `${s.entries_count} ${t("names")}` : t("empty")}
                      </span>
                    </button>
                  );
                })}
                {savedSessions.length === 0 && (
                  <p className="text-[10px] px-3.5 py-2" style={{ color: "var(--muted)" }}>{t("cb.session.no_sessions")}</p>
                )}
                <div style={{ borderTop: "1px solid var(--border)", marginTop: "4px", paddingTop: "4px" }}>
                  <button className="w-full text-left px-3.5 py-2 text-[11px] transition-colors hover:bg-[var(--surface-hover)]" style={{ color: "var(--muted)" }}
                    onClick={() => { setShowSessionDropdown(false); resetSession(); setExcelFile(""); setParseData(null); setSessionStarted(false); setSessionId(""); }}>
                    {t("cb.session.back")}
                  </button>
                  <button className="w-full text-left px-3.5 py-2 text-[11px] font-medium transition-colors hover:bg-[var(--surface-hover)]" style={{ color: "var(--accent)" }}
                    onClick={() => { setShowSessionDropdown(false); resetSession(); setExcelFile(""); setParseData(null); setSessionId(""); setSessionName(new Date().toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" })); }}>
                    {t("cb.session.new")}
                  </button>
                </div>
              </div>
            </>
          )}
        </div>
        <div className="flex-1" />
        {session?.user && (
          <button onClick={() => { clearAuthToken(); signOut({ callbackUrl: "/login" }); }} className="nav-btn" title={session.user.email || ""}>
            <span className="hidden sm:inline text-[10px]">{session.user.name || "O"}</span>
            <span className="sm:hidden text-[10px]">{(session.user.name || "O").charAt(0).toUpperCase()}</span>
            <span className="text-[9px]" style={{ opacity: 0.6 }}>{t("nav.signout")}</span>
          </button>
        )}
        <button onClick={toggleLang} className="nav-btn">{lang === "en" ? "TH" : "EN"}</button>
        <button onClick={toggle} className="theme-toggle">{theme === "light" ? "☾" : "☀"}</button>
        <div className="hidden sm:flex items-center gap-1.5">
          <Image src="/ossia-mark.svg?v3" alt="" width={24} height={17} className="micro-logo" />
          <span className="text-sm font-normal" style={{ color: "var(--foreground)", letterSpacing: "-0.035em" }}>ossia</span>
        </div>
      </nav>

      <main className="flex-1 max-w-6xl w-full mx-auto px-4 sm:px-6 py-4 sm:py-5 flex flex-col gap-3 sm:gap-4 min-h-0">

        {/* Title + Step Indicator */}
        <div style={{ animation: "slideUp 0.4s ease" }}>
          <h1 className="text-2xl font-medium tracking-tight mb-3" style={{ letterSpacing: "-0.02em" }}>{t("cb.title")}</h1>
          <StepIndicator
            steps={[
              { done: step0Done, active: !step0Done },
              { done: step1Done, active: step0Done && !step1Done },
              { done: step2Done, active: step1Done && !step2Done },
              { done: step3Done, active: step2Done && !step3Done },
            ]}
            labels={[t("cb.step.generate_ma_com"), t("cb.step.upload_order"), t("cb.step.upload_programs"), t("cb.step.export")]}
          />
        </div>

        {/* ── Step 0: Auto-assign MA & COM ── */}
        {assignMode !== "skipped" && !mappingConfirmed && (
          <div className="glass-panel overflow-hidden" style={{ animation: "fadeSlideIn 0.35s ease" }}>
            {/* Header */}
            <div className="px-5 py-4 flex items-center justify-between" style={{ borderBottom: "1px solid var(--border)" }}>
              <div>
                <h3 className="text-sm font-medium">{t("cb.assign.title")}</h3>
                <p className="text-[11px] mt-0.5" style={{ color: "var(--muted)" }}>
                  {assignMode === "pending" ? t("cb.assign.subtitle") : assignActiveField ? `→ ${t("cb.mapping.click_column")} (${t(`cb.assign.field.${assignActiveField}`)})` : t("cb.assign.detect_desc")}
                </p>
              </div>
              <button onClick={() => { setAssignMode("skipped"); }} className="glass-btn text-[10px]">{t("cb.assign.skip")}</button>
            </div>

            {/* Upload zone for assign step */}
            {assignMode === "pending" && (
              <div className="p-5">
                <div
                  className={`drop-zone ${excelFile ? "has-file" : ""}`}
                  onClick={() => assignExcelInputRef.current?.click()}
                  onDragOver={(e) => { e.preventDefault(); setExcelDragOver(true); }}
                  onDragLeave={() => setExcelDragOver(false)}
                  onDrop={(e) => { e.preventDefault(); setExcelDragOver(false); const file = e.dataTransfer.files[0]; if (file) uploadAssignExcel(file); }}
                >
                  <input ref={assignExcelInputRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={(e) => { if (e.target.files?.[0]) uploadAssignExcel(e.target.files[0]); e.target.value = ""; }} />
                  {assignLoading ? (
                    <ExcelSkeleton label={t("cb.excel.parsing")} />
                  ) : (
                    <div>
                      <svg className="mx-auto mb-2.5 mt-1.5 opacity-25" width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" /><polyline points="14 2 14 8 20 8" /><line x1="16" y1="13" x2="8" y2="13" /><line x1="16" y1="17" x2="8" y2="17" /></svg>
                      <p className="text-sm font-medium mb-0.5">{t("cb.excel.title")}</p>
                      <p className="text-[11px]" style={{ color: "var(--muted)" }}>{t("cb.assign.upload_hint")}</p>
                    </div>
                  )}
                </div>
                {/* How it works */}
                <div className="mt-4 p-3 rounded-lg text-[11px]" style={{ background: "var(--surface)", border: "1px solid var(--border)" }}>
                  <p className="font-medium mb-1.5" style={{ color: "var(--foreground)" }}>{t("cb.assign.how_title")}</p>
                  <ul className="flex flex-col gap-1" style={{ color: "var(--muted)" }}>
                    <li>• {t("cb.assign.how_ma")}</li>
                    <li>• {t("cb.assign.how_com")}</li>
                    <li>• {t("cb.assign.how_restart")}</li>
                  </ul>
                </div>
              </div>
            )}

            {/* Column picker for assign */}
            {assignMode === "detect" && assignDetectData && (
              <div className="p-5">
                {/* Field cards */}
                <div className="flex flex-wrap gap-2 mb-4">
                  {ASSIGN_FIELD_KEYS.map((field) => {
                    const fc = ASSIGN_FIELD_COLORS[field];
                    const colIdx = assignColumnMapping[field];
                    const isActive = assignActiveField === field;
                    const headerName = colIdx >= 0 && colIdx < assignDetectData.headers.length ? assignDetectData.headers[colIdx] : null;
                    return (
                      <button
                        key={field}
                        onClick={() => setAssignActiveField(isActive ? null : field)}
                        className="flex items-center gap-2 px-3 py-2 rounded-lg text-[11px] font-medium transition-all"
                        style={{
                          background: isActive ? fc.bg : colIdx >= 0 ? fc.cell : "var(--surface)",
                          border: `1.5px solid ${isActive ? fc.text : colIdx >= 0 ? fc.border : "var(--border)"}`,
                          color: colIdx >= 0 ? fc.text : "var(--muted)",
                          boxShadow: isActive ? `0 0 8px ${fc.glow}` : "none",
                        }}
                      >
                        <span>{t(`cb.assign.field.${field}`)}</span>
                        {headerName && <span className="text-[9px] opacity-70">({String.fromCharCode(65 + colIdx)}: {headerName})</span>}
                        {colIdx < 0 && <span className="text-[9px]">?</span>}
                      </button>
                    );
                  })}
                </div>

                {/* Spreadsheet preview */}
                <div className="overflow-x-auto custom-scroll rounded-lg mb-4" style={{ border: "1px solid var(--border)" }}>
                  <table className="text-[10px] border-collapse w-full" style={{ fontFamily: "var(--font-geist-mono)", minWidth: `${Math.max(assignDetectData.headers.length * 80, 400)}px` }}>
                    <thead>
                      <tr>
                        {assignDetectData.headers.map((header, colIdx) => {
                          const assignedField = Object.entries(assignColumnMapping).find(([, idx]) => idx === colIdx)?.[0];
                          const fc = assignedField ? ASSIGN_FIELD_COLORS[assignedField] : null;
                          return (
                            <th
                              key={colIdx}
                              onClick={() => handleAssignColumnClick(colIdx)}
                              style={{
                                padding: "8px 10px", textAlign: "center",
                                background: fc ? fc.bg : assignActiveField ? "var(--surface-hover)" : "var(--surface)",
                                borderRight: "1px solid var(--border)",
                                borderBottom: `3px solid ${fc ? fc.text : "transparent"}`,
                                color: fc ? fc.text : "var(--muted)",
                                cursor: assignActiveField ? "crosshair" : assignedField ? "pointer" : "default",
                                minWidth: "80px",
                              }}
                            >
                              <div className="font-bold text-[11px]">{String.fromCharCode(65 + colIdx)}</div>
                              {header && <div className="text-[8px] opacity-70 mt-0.5 truncate max-w-[70px]">{header}</div>}
                              {assignedField && fc && (
                                <div className="text-[7px] mt-1 font-sans font-medium uppercase tracking-wider" style={{ color: fc.text }}>
                                  {t(`cb.assign.field.${assignedField}`)}
                                </div>
                              )}
                            </th>
                          );
                        })}
                      </tr>
                    </thead>
                    <tbody>
                      {assignDetectData.preview_rows.slice(0, 4).map((row, rowIdx) => (
                        <tr key={rowIdx}>
                          {assignDetectData.headers.map((_, colIdx) => {
                            const assignedField = Object.entries(assignColumnMapping).find(([, idx]) => idx === colIdx)?.[0];
                            const fc = assignedField ? ASSIGN_FIELD_COLORS[assignedField] : null;
                            const val = colIdx < row.length ? row[colIdx] : null;
                            return (
                              <td key={colIdx} onClick={() => handleAssignColumnClick(colIdx)} style={{
                                padding: "5px 10px", textAlign: "center",
                                background: fc ? fc.cell : "transparent",
                                borderRight: "1px solid var(--border)",
                                borderBottom: rowIdx < 3 ? "1px solid var(--border)" : "none",
                                cursor: assignActiveField ? "crosshair" : "default",
                                color: fc ? fc.text : "var(--foreground)",
                              }}>
                                {val !== null ? String(val) : ""}
                              </td>
                            );
                          })}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                {/* Generate button */}
                <button
                  className="accent-btn w-full"
                  onClick={runAutoAssign}
                  disabled={assignLoading || !ASSIGN_FIELD_KEYS.every(f => assignColumnMapping[f] >= 0)}
                >
                  {assignLoading ? t("cb.assign.generating") : t("cb.assign.generate_btn")}
                </button>
              </div>
            )}

            {/* Results */}
            {assignMode === "result" && assignResult && (
              <div className="p-5">
                <h4 className="text-sm font-medium mb-3">{t("cb.assign.result_title")}</h4>

                {/* Summary stats */}
                <div className="grid grid-cols-3 gap-3 mb-4">
                  <div className="stat-card"><div className="stat-number">{assignResult.ma_summary.length}</div><div className="stat-label">{t("cb.assign.ma_groups")}</div></div>
                  <div className="stat-card"><div className="stat-number">{assignResult.com_summary.length}</div><div className="stat-label">{t("cb.assign.com_groups")}</div></div>
                  <div className="stat-card"><div className="stat-number">{assignResult.assignments_count}</div><div className="stat-label">{t("cb.assign.total_rows")}</div></div>
                </div>

                {/* MA Summary table */}
                <div className="overflow-x-auto custom-scroll rounded-lg mb-4" style={{ border: "1px solid var(--border)" }}>
                  <table className="text-[11px] border-collapse w-full">
                    <thead>
                      <tr style={{ background: "var(--surface)" }}>
                        <th className="px-3 py-2 text-left font-medium" style={{ borderBottom: "1px solid var(--border)", color: "var(--muted)" }}>{t("cb.assign.ma_label")}</th>
                        <th className="px-3 py-2 text-left font-medium" style={{ borderBottom: "1px solid var(--border)", color: "var(--muted)" }}>{t("cb.assign.field.size")}</th>
                        <th className="px-3 py-2 text-right font-medium" style={{ borderBottom: "1px solid var(--border)", color: "var(--muted)" }}>{t("cb.assign.entries")}</th>
                      </tr>
                    </thead>
                    <tbody>
                      {assignResult.ma_summary.map((ma) => (
                        <tr key={ma.ma}>
                          <td className="px-3 py-2 font-medium" style={{ color: "#f59e0b", borderBottom: "1px solid var(--border)" }}>{ma.ma}</td>
                          <td className="px-3 py-2" style={{ borderBottom: "1px solid var(--border)" }}>{ma.size}</td>
                          <td className="px-3 py-2 text-right" style={{ color: "var(--muted)", borderBottom: "1px solid var(--border)" }}>{ma.count}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                {/* COM Summary table */}
                <div className="overflow-x-auto custom-scroll rounded-lg mb-4" style={{ border: "1px solid var(--border)", maxHeight: "200px", overflowY: "auto" }}>
                  <table className="text-[11px] border-collapse w-full">
                    <thead>
                      <tr style={{ background: "var(--surface)", position: "sticky", top: 0 }}>
                        <th className="px-3 py-2 text-left font-medium" style={{ borderBottom: "1px solid var(--border)", color: "var(--muted)" }}>{t("cb.assign.ma_label")}</th>
                        <th className="px-3 py-2 text-left font-medium" style={{ borderBottom: "1px solid var(--border)", color: "var(--muted)" }}>{t("cb.assign.com_label")}</th>
                        <th className="px-3 py-2 text-left font-medium" style={{ borderBottom: "1px solid var(--border)", color: "var(--muted)" }}>{t("cb.assign.fabric_col")}</th>
                        <th className="px-3 py-2 text-left font-medium" style={{ borderBottom: "1px solid var(--border)", color: "var(--muted)" }}>{t("cb.assign.frame_col")}</th>
                        <th className="px-3 py-2 text-left font-medium" style={{ borderBottom: "1px solid var(--border)", color: "var(--muted)" }}>{t("cb.assign.embroidery_col")}</th>
                        <th className="px-3 py-2 text-right font-medium" style={{ borderBottom: "1px solid var(--border)", color: "var(--muted)" }}>{t("cb.assign.entries")}</th>
                      </tr>
                    </thead>
                    <tbody>
                      {assignResult.com_summary.map((c, i) => (
                        <tr key={i}>
                          <td className="px-3 py-1.5 font-medium" style={{ color: "#f59e0b", borderBottom: "1px solid var(--border)" }}>{c.ma}</td>
                          <td className="px-3 py-1.5 font-medium" style={{ color: "var(--accent)", borderBottom: "1px solid var(--border)" }}>{c.com}</td>
                          <td className="px-3 py-1.5" style={{ borderBottom: "1px solid var(--border)" }}>{c.fabric_colour}</td>
                          <td className="px-3 py-1.5" style={{ borderBottom: "1px solid var(--border)" }}>{c.frame_colour}</td>
                          <td className="px-3 py-1.5" style={{ borderBottom: "1px solid var(--border)" }}>{c.embroidery_colour}</td>
                          <td className="px-3 py-1.5 text-right" style={{ color: "var(--muted)", borderBottom: "1px solid var(--border)" }}>{c.count}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                {/* Action buttons */}
                <div className="flex flex-col sm:flex-row gap-2">
                  <button className="glass-btn flex-1 text-[11px] py-2.5" onClick={downloadAssignedExcel} disabled={assignDownloading}>
                    {assignDownloading ? t("cb.assign.downloading") : t("cb.assign.download_excel")}
                  </button>
                  <button className="accent-btn flex-1" onClick={proceedFromAssign} disabled={assignLoading}>
                    {assignLoading ? t("cb.assign.generating") : t("cb.assign.proceed")}
                  </button>
                </div>
              </div>
            )}
          </div>
        )}

        {/* Upload Zones */}
        {(assignMode === "skipped" || mappingConfirmed) && (<>
        <div className="grid grid-cols-1 md:grid-cols-2 gap-3" style={{ animation: "slideUp 0.4s ease 0.05s forwards", opacity: 0 }}>
          {/* Excel drop zone */}
          <div
            className={`drop-zone ${excelDragOver ? "drag-over" : ""} ${excelFile ? "has-file" : ""}`}
            onDragOver={(e) => { e.preventDefault(); setExcelDragOver(true); }}
            onDragLeave={() => setExcelDragOver(false)}
            onDrop={handleExcelDrop}
            onClick={() => excelInputRef.current?.click()}
          >
            <input ref={excelInputRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={(e) => { if (e.target.files?.[0]) uploadExcel(e.target.files[0]); e.target.value = ""; }} />
            {excelLoading && !detectData ? (
              <ExcelSkeleton label={t("cb.excel.parsing")} />
            ) : excelFile ? (
              <div style={{ animation: "fadeSlideIn 0.3s ease" }}>
                <div className="flex items-center justify-center gap-2">
                  <svg width="14" height="14" viewBox="0 0 12 12" fill="none" stroke="var(--accent)" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="2,6 5,9 10,3" /></svg>
                  <span className="text-sm font-medium" style={{ color: "var(--accent)" }}>{excelFile}</span>
                  <button
                    className="w-5 h-5 flex items-center justify-center rounded-md text-[10px] transition-colors"
                    style={{ color: "var(--muted)", background: "var(--surface)" }}
                    onClick={(e) => { e.stopPropagation(); removeExcel(); }}
                    title={t("cb.excel.remove")}
                  >✕</button>
                </div>
              </div>
            ) : (
              <div>
                <span className="text-[9px] font-medium font-mono uppercase tracking-wider" style={{ color: "var(--accent)", opacity: 0.5 }}>01</span>
                <svg className="mx-auto mb-2.5 mt-1.5 opacity-25" width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" /><polyline points="14 2 14 8 20 8" /><line x1="16" y1="13" x2="8" y2="13" /><line x1="16" y1="17" x2="8" y2="17" /></svg>
                <p className="text-sm font-medium mb-0.5">{t("cb.excel.title")}</p>
                <p className="text-[11px]" style={{ color: "var(--muted)" }}>{t("cb.excel.hint")}</p>
              </div>
            )}
          </div>

          {/* DST drop zone */}
          {!mappingConfirmed ? (
            <div className="drop-zone-locked">
              <svg className="mx-auto mb-2.5 mt-1.5" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" style={{ color: "var(--muted)", opacity: 0.6 }}>
                <rect x="3" y="11" width="18" height="11" rx="2" ry="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/>
              </svg>
              <p className="text-sm font-medium mb-0.5" style={{ color: "var(--muted)" }}>{t("cb.dst.title")}</p>
              <p className="text-[11px]" style={{ color: "var(--muted)", opacity: 0.7 }}>{t("cb.dst.hint_disabled")}</p>
            </div>
          ) : (
            <div
              className={`drop-zone ${dstDragOver ? "drag-over" : ""} ${dstUploaded ? "has-file" : ""}`}
              onDragOver={(e) => { e.preventDefault(); setDstDragOver(true); }}
              onDragLeave={() => setDstDragOver(false)}
              onDrop={handleDstDrop}
              onClick={() => zipInputRef.current?.click()}
            >
              <input ref={zipInputRef} type="file" accept=".zip" className="hidden" onChange={(e) => { if (e.target.files) uploadDst(e.target.files); e.target.value = ""; }} />
              {dstLoading ? (
                <div className="flex flex-col items-center gap-3" style={{ animation: "fadeIn 0.2s ease" }}>
                  <div className="spinner" />
                  <p className="text-sm" style={{ color: "var(--accent)" }}>{t("cb.dst.uploading")}</p>
                </div>
              ) : dstUploaded && dstData ? (
                <div style={{ animation: "fadeSlideIn 0.3s ease" }}>
                  <div className="flex items-center justify-center gap-2">
                    <span style={{ color: dstData.all_matched ? "var(--accent)" : "var(--warning)", fontSize: "16px" }}>{dstData.all_matched ? "✓" : "⚠"}</span>
                    <span className="text-sm font-medium" style={{ color: dstData.all_matched ? "var(--accent)" : "var(--warning)" }}>{dstFileName}</span>
                  </div>
                  <p className="text-[11px] mt-1" style={{ color: "var(--muted)" }}>
                    {dstData.needed_count > 0 ? <>{(dstData.needed_count - dstData.missing_programs.length)}/{dstData.needed_count} {t("cb.dst.matched")}{dstData.missing_programs.length > 0 && <span style={{ color: "var(--danger)" }}> · {dstData.missing_programs.length} {t("cb.dst.missing")}</span>}</> : <>{dstData.uploaded_count} {t("cb.dst.uploaded")}</>}
                  </p>
                </div>
              ) : (
                <div>
                  <span className="text-[9px] font-medium font-mono uppercase tracking-wider" style={{ color: "var(--accent)", opacity: 0.5 }}>02</span>
                  <svg className="mx-auto mb-2.5 mt-1.5 opacity-25" width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" /><polyline points="17 8 12 3 7 8" /><line x1="12" y1="3" x2="12" y2="15" /></svg>
                  <p className="text-sm font-medium mb-0.5">{t("cb.dst.title")}</p>
                  <p className="text-[11px]" style={{ color: "var(--muted)" }}>{t("cb.dst.hint")}</p>
                </div>
              )}
            </div>
          )}
        </div>

        {/* ── Column Mapping (Interactive) ── */}
        {detectData && !mappingConfirmed && (
          <div className="glass-panel overflow-hidden" style={{ animation: "fadeSlideIn 0.35s ease" }}>

            {/* Header */}
            <div className="px-5 py-4 flex items-center justify-between" style={{ borderBottom: "1px solid var(--border)" }}>
              <div>
                <h3 className="text-sm font-medium">{t("cb.mapping.title")}</h3>
                <p className="text-[11px] mt-0.5" style={{ color: activeField ? "var(--accent)" : "var(--muted)" }}>
                  {activeField
                    ? `→ ${t("cb.mapping.click_column")} (${t(`cb.mapping.field.${activeField}`)})`
                    : detectData.confidence === "high" ? t("cb.mapping.auto") : t("cb.mapping.review")}
                </p>
              </div>
              <div className="flex items-center gap-2">
                {activeField && (
                  <button onClick={() => setActiveField(null)} className="glass-btn text-[10px]">{t("cb.mapping.esc")}</button>
                )}
                <span className="flex items-center gap-1.5 text-[10px] font-medium px-2.5 py-1 rounded-full" style={{
                  background: detectData.confidence === "high" ? "rgba(22,163,74,0.1)"
                    : detectData.confidence === "medium" ? "rgba(245,158,11,0.1)"
                    : "rgba(239,68,68,0.1)",
                  color: detectData.confidence === "high" ? "var(--success)"
                    : detectData.confidence === "medium" ? "var(--warning)"
                    : "var(--danger)",
                }}>
                  <span className="w-2 h-2 rounded-full" style={{
                    background: "currentColor",
                    animation: detectData.confidence === "low" ? "pulse 1.5s infinite" : "none",
                  }} />
                  {t(`cb.mapping.confidence_${detectData.confidence}`)}
                </span>
              </div>
            </div>

            {/* Confidence warning banners */}
            {detectData.confidence === "low" && (
              <div className="px-5 py-3.5 text-[11px] flex items-start gap-3" style={{ background: "rgba(239, 68, 68, 0.06)", borderBottom: "1px solid rgba(239, 68, 68, 0.2)", borderLeft: "4px solid var(--danger)", color: "var(--danger)" }}>
                <span className="text-base leading-none shrink-0">⚠</span>
                <div>
                  <p className="font-semibold mb-0.5">{t("cb.mapping.warn_low_title")}</p>
                  <p style={{ opacity: 0.85 }}>{t("cb.mapping.warn_low")}</p>
                </div>
              </div>
            )}
            {detectData.confidence === "medium" && (
              <div className="px-5 py-3.5 text-[11px] flex items-start gap-3" style={{ background: "rgba(245, 158, 11, 0.06)", borderBottom: "1px solid rgba(245, 158, 11, 0.2)", borderLeft: "4px solid var(--warning)", color: "var(--warning)" }}>
                <span className="text-base leading-none shrink-0">⚠</span>
                <div>
                  <p className="font-semibold mb-0.5">{t("cb.mapping.warn_medium_title")}</p>
                  <p style={{ opacity: 0.85 }}>{t("cb.mapping.warn_medium")}</p>
                </div>
              </div>
            )}

            {/* ── Spreadsheet Preview ── */}
            <div className="px-5 pt-4 pb-2">
              <p className="text-[10px] font-medium uppercase tracking-wider mb-2" style={{ color: "var(--muted)", letterSpacing: "0.1em" }}>
                {t("cb.mapping.spreadsheet_preview")}
                {activeField && <span style={{ color: "var(--accent)" }}> — {t("cb.mapping.click_to_assign")}</span>}
              </p>
              <div className="overflow-x-auto custom-scroll rounded-lg" style={{ border: "1px solid var(--border)" }}>
                <table className="text-[10px] border-collapse w-full" style={{ fontFamily: "var(--font-geist-mono)", minWidth: `${Math.max(detectData.headers.length * 80, 400)}px` }}>
                  <thead>
                    <tr>
                      {detectData.headers.map((header, colIdx) => {
                        const assignedField = colToField[colIdx];
                        const fc = assignedField ? FIELD_COLORS[assignedField] : null;
                        const isAssignedToActive = activeField && columnMapping[activeField] === colIdx;
                        return (
                          <th
                            key={colIdx}
                            onClick={() => handleColumnClick(colIdx)}
                            title={header || undefined}
                            style={{
                              padding: "8px 10px",
                              textAlign: "center",
                              background: fc ? fc.bg : activeField ? "var(--surface-hover)" : "var(--surface)",
                              borderRight: "1px solid var(--border)",
                              borderBottom: `3px solid ${fc ? fc.text : "transparent"}`,
                              color: fc ? fc.text : "var(--muted)",
                              cursor: activeField ? "crosshair" : assignedField ? "pointer" : "default",
                              transition: "all 0.15s ease",
                              minWidth: "80px",
                              position: "relative",
                              outline: isAssignedToActive ? `2px solid ${FIELD_COLORS[activeField!]?.text}` : fc && activeField ? `2px solid ${fc.text}` : "none",
                              outlineOffset: "-2px",
                            }}
                          >
                            <div className="font-bold text-[11px]">{String.fromCharCode(65 + colIdx)}</div>
                            {header && <div className="text-[8px] opacity-70 mt-0.5 truncate max-w-[70px]">{header}</div>}
                            {assignedField && (
                              <div className="text-[7px] mt-1 font-sans font-medium uppercase tracking-wider" style={{ color: fc?.text }}>
                                {t(`cb.mapping.field.${assignedField}`)}
                              </div>
                            )}
                          </th>
                        );
                      })}
                    </tr>
                  </thead>
                  <tbody>
                    {detectData.preview_rows.slice(0, 4).map((row, rowIdx) => (
                      <tr key={rowIdx}>
                        {detectData.headers.map((_, colIdx) => {
                          const assignedField = colToField[colIdx];
                          const fc = assignedField ? FIELD_COLORS[assignedField] : null;
                          const val = colIdx < row.length ? row[colIdx] : null;
                          return (
                            <td
                              key={colIdx}
                              onClick={() => handleColumnClick(colIdx)}
                              style={{
                                padding: "5px 10px",
                                textAlign: "center",
                                background: fc ? fc.cell : "transparent",
                                borderRight: "1px solid var(--border)",
                                borderBottom: "1px solid var(--border)",
                                color: fc ? fc.text : "var(--foreground)",
                                cursor: activeField ? "crosshair" : "default",
                                transition: "background 0.15s ease",
                                opacity: val === null ? 0.3 : 1,
                              }}
                            >
                              {val !== null ? String(val) : "·"}
                            </td>
                          );
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            {/* ── Field Cards ── */}
            <div className="px-5 py-4 grid grid-cols-2 sm:grid-cols-3 gap-2.5">
              {FIELD_KEYS.map((field) => {
                const colIdx = columnMapping[field] ?? -1;
                const assigned = colIdx >= 0;
                const isActive = activeField === field;
                const fc = FIELD_COLORS[field];
                const colLetter = assigned ? String.fromCharCode(65 + colIdx) : null;
                const isOptional = field === "name_line2";
                const samples = assigned
                  ? detectData.preview_rows.slice(0, 3).map(r => colIdx < r.length ? r[colIdx] : null).filter(v => v !== null && v !== "").map(v => String(v))
                  : [];
                const isRequired = !isOptional;
                const isMissing = isRequired && !assigned;

                return (
                  <div
                    key={field}
                    className={`field-card ${isActive ? "field-card-active" : ""}`}
                    onClick={() => setActiveField(isActive ? null : field)}
                    tabIndex={0}
                    onKeyDown={(e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); setActiveField(isActive ? null : field); } }}
                    style={{
                      padding: "12px",
                      borderRadius: "12px",
                      cursor: "pointer",
                      transition: "all 0.18s cubic-bezier(.4,0,.2,1)",
                      background: isActive ? fc.bg : assigned ? `${fc.bg}` : "var(--surface)",
                      border: isActive
                        ? `2px solid ${fc.text}`
                        : assigned
                          ? `1.5px solid ${fc.border}`
                          : isMissing
                            ? "1.5px dashed rgba(239,68,68,0.35)"
                            : "1.5px dashed var(--border-strong)",
                      boxShadow: isActive ? `0 0 0 4px ${fc.glow}, 0 4px 16px ${fc.glow}` : "none",
                    }}
                  >
                    {/* Label row */}
                    <div className="flex items-center gap-1.5 mb-2">
                      <div className="w-2 h-2 rounded-full shrink-0" style={{ background: assigned || isActive ? fc.text : isMissing ? "rgba(239,68,68,0.5)" : "var(--border-strong)" }} />
                      <span className="text-[10px] font-semibold" style={{ color: assigned || isActive ? fc.text : isMissing ? "var(--danger)" : "var(--muted)", letterSpacing: "0.02em" }}>
                        {t(`cb.mapping.field.${field}`)}
                      </span>
                      {assigned && fieldDetectionSource[field] && (
                        <span className="text-[7px] px-1.5 py-0.5 rounded" style={{
                          background: fieldDetectionSource[field] === "auto" ? "rgba(22,163,74,0.1)" : "rgba(245,158,11,0.1)",
                          color: fieldDetectionSource[field] === "auto" ? "var(--success)" : "var(--warning)",
                        }}>
                          {t(`cb.mapping.${fieldDetectionSource[field] === "auto" ? "auto_detected" : "position_guessed"}`)}
                        </span>
                      )}
                      {isOptional && (
                        <span className="ml-auto text-[7px] px-1.5 py-0.5 rounded" style={{ background: "var(--surface)", color: "var(--muted)", border: "1px solid var(--border)" }}>opt</span>
                      )}
                    </div>

                    {/* Assignment */}
                    {assigned ? (
                      <div>
                        <div className="text-[22px] font-bold font-mono leading-none mb-1" style={{ color: fc.text, letterSpacing: "-0.03em" }}>
                          {colLetter}
                        </div>
                        {samples.length > 0 && (
                          <div className="text-[9px] font-mono leading-relaxed" style={{ color: "var(--muted)" }}>
                            {samples.map((s, i) => <span key={i}>{i > 0 ? ", " : ""}<span style={{ color: fc.text, opacity: 0.8 }}>{s}</span></span>)}
                          </div>
                        )}
                      </div>
                    ) : (
                      <div className="text-[9px] leading-snug" style={{ color: isActive ? fc.text : "var(--muted)" }}>
                        {isActive ? (
                          <span className="font-medium">↑ click a column</span>
                        ) : (
                          <span>{t(`cb.mapping.help.${field}`)}</span>
                        )}
                      </div>
                    )}
                    {mappingWarnings[field] && (
                      <div className="mt-1.5 text-[9px] px-2 py-1 rounded" style={{
                        background: "rgba(245, 158, 11, 0.08)",
                        color: "var(--warning)",
                        border: "1px solid rgba(245, 158, 11, 0.2)",
                      }}>
                        ⚠ {mappingWarnings[field]}
                      </div>
                    )}
                  </div>
                );
              })}
            </div>

            {/* How stacking works (collapsible) */}
            <div className="px-5 pb-4">
              <button
                onClick={() => setShowHowItWorks(!showHowItWorks)}
                className="flex items-center gap-1.5 text-[11px] mb-2 transition-colors"
                style={{ color: "var(--muted)" }}
              >
                <span className="text-[9px]">{showHowItWorks ? "▼" : "▶"}</span>
                {t("cb.mapping.how_it_works")}
              </button>
              {showHowItWorks && (
                <div className="p-3 rounded-lg text-[10px]" style={{ background: "var(--surface)", border: "1px solid var(--border)", animation: "fadeSlideIn 0.2s ease" }}>
                  <ul className="space-y-1" style={{ color: "var(--muted)" }}>
                    <li>· {t("cb.mapping.how_grouping")}</li>
                    <li>· {t("cb.mapping.how_quantity")}</li>
                    <li>· {t("cb.mapping.how_program")}</li>
                    <li>· {t("cb.mapping.how_names")}</li>
                  </ul>
                </div>
              )}
            </div>

            {/* Confirm button */}
            <div className="px-5 pb-5 flex items-center justify-between gap-3" style={{ borderTop: "1px solid var(--border)", paddingTop: "16px" }}>
              <p className="text-[10px]" style={{ color: "var(--muted)" }}>
                {REQUIRED_FIELDS.filter(f => (columnMapping[f] ?? -1) < 0).length === 0
                  ? <span style={{ color: "var(--success)" }}>✓ {t("cb.mapping.all_assigned")}</span>
                  : <span>{REQUIRED_FIELDS.filter(f => (columnMapping[f] ?? -1) < 0).length} {t("cb.mapping.fields_remaining")}</span>
                }
              </p>
              <button
                className="accent-btn text-xs"
                onClick={confirmMapping}
                disabled={excelLoading || REQUIRED_FIELDS.some(f => (columnMapping[f] ?? -1) < 0)}
              >
                {excelLoading ? (
                  <span className="flex items-center gap-2"><span className="spinner" style={{ width: "12px", height: "12px", borderWidth: "1.5px" }} />{t("cb.excel.parsing")}</span>
                ) : t("cb.mapping.confirm")}
              </button>
            </div>
          </div>
        )}
        </>)}

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
                        <th key={k} className="text-left py-1.5 px-2 font-medium" style={{ color: "var(--muted)" }}>{t(k)}</th>
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
            <StatCard
              value={parseData.entries_count} label={t("cb.stats.names")} delay={0.05}
              icon={<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>}
            />
            <span className="stat-arrow hidden sm:block">→</span>
            <StatCard
              value={parseData.groups.length} label={t("cb.stats.groups")} delay={0.1}
              icon={<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/></svg>}
            />
            <span className="stat-arrow hidden sm:block">→</span>
            <StatCard
              value={parseData.combo_count} label={t("cb.stats.output")} delay={0.15}
              icon={<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>}
            />
            <span className="stat-arrow hidden sm:block">→</span>
            <StatCard
              value={parseData.total_slots} label={t("cb.stats.slots")} delay={0.2}
              icon={<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><polygon points="12 2 2 7 12 12 22 7 12 2"/><polyline points="2 17 12 22 22 17"/><polyline points="2 12 12 17 22 12"/></svg>}
            />
          </div>
        )}

        {/* Two-Panel: Combo List + Slot Preview */}
        {parseData && parseData.groups.length > 0 && (
          <div className="flex-1 grid grid-cols-1 lg:grid-cols-[1fr_360px] gap-3 min-h-0" style={{ animation: "slideUp 0.4s ease 0.15s forwards", opacity: 0 }}>
            <div className="glass-panel overflow-hidden flex flex-col min-h-[200px] sm:min-h-[300px]">
              <div className="flex items-center justify-between px-4 py-2.5 shrink-0" style={{ borderBottom: "1px solid var(--border)" }}>
                <span className="text-xs font-medium">{t("cb.files.title")} <span className="font-normal" style={{ color: "var(--muted)" }}>{selectedCombos.size}/{totalCombos}</span></span>
                <div className="flex gap-1">
                  <button onClick={selectAll} className="glass-btn text-[10px]">{t("cb.files.all")}</button>
                  <button onClick={deselectAll} className="glass-btn text-[10px]">{t("cb.files.none")}</button>
                </div>
              </div>
              <div className="flex-1 overflow-y-auto custom-scroll">
                {parseData.groups.map((group) => {
                  const gk = `${group.machine_program}_${group.com_no}`;
                  const exp = expandedGroups.has(gk);
                  return (
                    <div key={gk} style={{ borderBottom: "1px solid var(--border)" }}>
                      <button className="w-full flex flex-col px-4 py-3 sm:py-2.5 text-left transition-colors" style={{ background: exp ? "var(--surface)" : "transparent" }} onClick={() => toggleGroup(gk)}>
                        <div className="flex items-center gap-2.5 w-full">
                          <span className="text-[10px] w-3" style={{ color: "var(--muted)" }}>{exp ? "▼" : "▶"}</span>
                          <span className="text-xs font-medium">{group.machine_program}<span style={{ color: "var(--muted)", fontWeight: 400 }}> / {t("Combo")} {group.com_no}</span></span>
                          <span className="text-[10px] ml-auto tabular-nums hidden sm:inline" style={{ color: "var(--muted)" }}>
                            {group.entry_count} {t("names")} → {group.combos.length} {group.combos.length === 1 ? t("file") : t("files")} ({group.total_slots} {t("slots")})
                          </span>
                          <span className="text-[10px] ml-auto tabular-nums sm:hidden" style={{ color: "var(--muted)" }}>{group.combos.length}f</span>
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
                  <h3 className="text-[11px] font-medium font-mono mb-3 pb-2" style={{ borderBottom: "1px solid var(--border)", color: "var(--accent)" }}>{previewCombo.filename}</h3>
                  <div className="grid grid-cols-2 gap-4">
                    <div>
                      <p className="text-[9px] font-medium uppercase tracking-wider mb-1.5" style={{ color: "var(--muted)" }}>{t("cb.preview.left")}</p>
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
                      <p className="text-[9px] font-medium uppercase tracking-wider mb-1.5" style={{ color: "var(--muted)" }}>{t("cb.preview.right")}</p>
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
                <h3 className="text-[11px] font-medium font-mono" style={{ color: "var(--accent)" }}>{previewCombo.filename}</h3>
                <button onClick={() => setShowMobilePreview(false)} className="w-7 h-7 flex items-center justify-center rounded-lg" style={{ color: "var(--muted)", background: "var(--surface)" }}>✕</button>
              </div>
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <p className="text-[9px] font-medium uppercase tracking-wider mb-1.5" style={{ color: "var(--muted)" }}>{t("cb.preview.left")}</p>
                  {previewCombo.slots.slice(0, 10).map((s, i) => (
                    <div key={i} className="flex items-center gap-1 py-[2px]" style={{ borderBottom: "1px solid var(--border)" }}>
                      <span className="text-[9px] w-3.5 text-right tabular-nums" style={{ color: "var(--border-strong)" }}>{i + 1}</span>
                      <span className="text-[9px] font-mono w-6 tabular-nums" style={{ color: "var(--accent)" }}>{s.program}</span>
                      <span className="text-[9px] truncate">{s.name_line1}</span>
                    </div>
                  ))}
                </div>
                <div>
                  <p className="text-[9px] font-medium uppercase tracking-wider mb-1.5" style={{ color: "var(--muted)" }}>{t("cb.preview.right")}</p>
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
                <div className="flex items-center gap-4">
                  {/* ASCII sewing machine mascot */}
                  <pre className="text-[10px] leading-[1.15] shrink-0 hidden sm:block" style={{ fontFamily: "var(--font-geist-mono)", color: "var(--accent)", animation: "bobble 2s ease-in-out infinite" }}>{
                    exportProgress === -1
                      ? `  ,___,\n  (o.o)\n  / > >\n  \\_|_/\n   ~~~`
                      : exportProgress % 2 === 0
                        ? `  ,___,\n  (o.o)\n  />  >\n  \\_|_/\n   ~ ~`
                        : `  ,___,\n  (o.o)\n  <  <\\\n  \\_|_/\n   ~ ~`
                  }</pre>
                  <div className="flex-1 min-w-0">
                    <div className="flex items-center gap-2 mb-1">
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
              <span className="text-[9px] font-medium uppercase tracking-wider text-center sm:text-right" style={{ color: "var(--accent)", opacity: 0.5 }}>03</span>
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
