"use client";

import Link from "next/link";
import Image from "next/image";
import { useState, useCallback, useRef, useEffect, useMemo } from "react";
import { useSession as _useSession, signOut } from "next-auth/react";

const IS_LOCAL_MODE = process.env.NEXT_PUBLIC_LOCAL_MODE === "true";
// In local mode, skip useSession (no auth)
const useSession = IS_LOCAL_MODE ? () => ({ data: null }) as ReturnType<typeof _useSession> : _useSession;
import { useTheme } from "../theme-provider";
import { useLanguage } from "../i18n";
import { authFetch, clearAuthToken, warmupBackend } from "@/lib/api";

const API = process.env.NEXT_PUBLIC_API_URL || "http://localhost:8000";

interface Slot { program: number; name_line1: string; name_line2: string; quantity: number; }
interface ComboFile { filename: string; part_number: number; total_parts: number; slot_count: number; left_count: number; right_count: number; slots: Slot[]; head_mode?: string; }
interface Group { machine_program: string; com_no: string; entry_count: number; total_slots: number; combos: ComboFile[]; }
interface EntryPreview { program: number; name_line1: string; name_line2: string; quantity: number; com_no: string; machine_program: string; }
interface ParseResponse { session_id: string; entries_count: number; total_slots: number; optimized_slots?: number; slots_saved?: number; even_file_count?: number; odd_file_count?: number; groups: Group[]; combo_count: number; warnings: string[]; entries_preview?: EntryPreview[]; }
interface DetectResponse { session_id: string; excel_filename: string; headers: string[]; preview_rows: (string | number | null)[][]; detected_mapping: Record<string, number>; confidence: string; }
const FIELD_KEYS = ["program", "name_line1", "name_line2", "quantity", "com_no", "machine_program"] as const;
const REQUIRED_FIELDS = ["program", "name_line1", "quantity", "com_no", "machine_program"] as const;
const ASSIGN_FIELD_KEYS = ["size", "fabric_colour", "frame_colour", "embroidery_colour"] as const;
interface DstResponse { session_id: string; uploaded_count: number; needed_count: number; missing_programs: number[]; all_matched: boolean; }
interface SessionSummary { session_id: string; name: string; created_at: string; updated_at: string; expires_at: string | null; has_excel: boolean; entries_count: number; combo_count: number; dst_count: number; exported: boolean; }
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

function getTimeRemaining(expiresAt: string | null): { text: string; color: string; urgency: "ok" | "warn" | "danger" } | null {
  if (!expiresAt) return null;
  try {
    const now = Date.now();
    const expires = new Date(expiresAt).getTime();
    const hoursLeft = (expires - now) / (1000 * 60 * 60);
    if (hoursLeft <= 0) return { text: "Expired", color: "var(--danger)", urgency: "danger" };
    if (hoursLeft < 2) return { text: `${Math.ceil(hoursLeft * 60)}m left`, color: "var(--danger)", urgency: "danger" };
    if (hoursLeft < 12) return { text: `${Math.round(hoursLeft)}h left`, color: "var(--warning)", urgency: "warn" };
    return { text: `${Math.round(hoursLeft)}h left`, color: "var(--muted)", urgency: "ok" };
  } catch { return null; }
}

// Visual step stepper — completed steps are clickable for back navigation
function StepIndicator({ steps, labels, descriptions, onStepClick }: {
  steps: {done: boolean; active: boolean}[];
  labels: string[];
  descriptions?: string[];
  onStepClick?: (step: number) => void;
}) {
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
          <div
            className="flex flex-col items-center gap-1"
            style={{ cursor: s.done ? "pointer" : "default" }}
            onClick={() => s.done && onStepClick?.(i)}
            title={s.done ? `Go to: ${labels[i]}` : undefined}
          >
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
                ...(s.done ? { filter: "brightness(1)", transform: "scale(1)" } : {}),
              }}
              onMouseEnter={(e) => { if (s.done) { e.currentTarget.style.transform = "scale(1.15)"; e.currentTarget.style.filter = "brightness(1.2)"; }}}
              onMouseLeave={(e) => { if (s.done) { e.currentTarget.style.transform = "scale(1)"; e.currentTarget.style.filter = "brightness(1)"; }}}
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
            {descriptions?.[i] && (
              <span className="hidden sm:block" style={{
                fontSize: "8px", whiteSpace: "nowrap",
                color: "var(--muted)", opacity: 0.7,
                marginTop: "-2px",
              }}>{descriptions[i]}</span>
            )}
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
  const [removeExcelConfirm, setRemoveExcelConfirm] = useState(false);
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
  const [optimizeHeads, setOptimizeHeads] = useState(true);
  // New: interactive column assignment
  const [activeField, setActiveField] = useState<string | null>(null);
  const [showHowItWorks, setShowHowItWorks] = useState(true);
  // Auto-assign MA/COM state
  const [assignMode, setAssignMode] = useState<"pending" | "detect" | "result" | "skipped">("pending");
  const [assignDetectData, setAssignDetectData] = useState<{headers: string[]; preview_rows: (string|number|null)[][]; detected_mapping: Record<string, number>; confidence: string} | null>(null);
  const [assignColumnMapping, setAssignColumnMapping] = useState<Record<string, number>>({});
  const [assignActiveField, setAssignActiveField] = useState<string | null>(null);
  const [assignResult, setAssignResult] = useState<{assignments_count: number; ma_summary: {ma: string; size: string; count: number; is_new?: boolean}[]; com_summary: {ma: string; com: number; fabric_colour: string; frame_colour: string; embroidery_colour: string; count: number; is_new?: boolean}[]; warnings: string[]} | null>(null);
  const [assignLoading, setAssignLoading] = useState(false);
  const [assignDownloading, setAssignDownloading] = useState(false);
  // MA Reference table state
  const [maReference, setMaReference] = useState<{id?: number; size_normalized: string; size_display: string; ma_number: string}[] | null>(null);
  const [maRefLoading, setMaRefLoading] = useState(false);
  const [maRefExpanded, setMaRefExpanded] = useState(false);
  const maRefInputRef = useRef<HTMLInputElement>(null);
  // COM Reference state
  const [comReference, setComReference] = useState<{id?: number; ma_number: string; com_number: number; fabric_colour: string; embroidery_colour: string; frame_colour: string}[] | null>(null);
  const [expandedMaRows, setExpandedMaRows] = useState<Set<string>>(new Set());
  // Track which items have been added to reference from assign results
  const [addedMaRefs, setAddedMaRefs] = useState<Set<string>>(new Set());
  const [addedComRefs, setAddedComRefs] = useState<Set<string>>(new Set());
  const [backendStatus, setBackendStatus] = useState<"connecting" | "ready" | "failed">(IS_LOCAL_MODE ? "ready" : "connecting");
  const [updateInfo, setUpdateInfo] = useState<{ latest: string; update_url: string; installer_url?: string } | null>(null);
  const [currentVersion, setCurrentVersion] = useState("1.0.0");
  const [updateDismissed, setUpdateDismissed] = useState(false);
  const [updateChecking, setUpdateChecking] = useState(false);
  const [updateStatus, setUpdateStatus] = useState<"idle" | "downloading" | "installing" | "up-to-date" | "error">("idle");
  const [tourStep, setTourStep] = useState(-1); // -1 = hidden
  const assignExcelInputRef = useRef<HTMLInputElement>(null);
  const excelInputRef = useRef<HTMLInputElement>(null);
  const zipInputRef = useRef<HTMLInputElement>(null);

  // Step section refs for scroll-to-step navigation
  const step0Ref = useRef<HTMLDivElement>(null);
  const step1Ref = useRef<HTMLDivElement>(null);
  const step3Ref = useRef<HTMLDivElement>(null);

  const handleStepClick = (stepIndex: number) => {
    const refs = [step0Ref, step1Ref, step1Ref, step3Ref]; // Step 2 shares step1 zone
    refs[stepIndex]?.current?.scrollIntoView({ behavior: "smooth", block: "start" });
  };

  const showToast = useCallback((message: string, type: "error" | "warning" | "success" = "error") => setToast({ message, type }), []);

  const checkForUpdates = useCallback(async () => {
    setUpdateChecking(true);
    setUpdateStatus("idle");
    try {
      const res = await fetch(`${API}/api/version`);
      const data = await res.json();
      if (data.version) setCurrentVersion(data.version);
      if (data.update_available && data.latest && data.update_url) {
        setUpdateInfo({ latest: data.latest, update_url: data.update_url, installer_url: data.installer_url });
        setUpdateDismissed(false);
        localStorage.removeItem("update-dismissed");
      } else {
        setUpdateStatus("up-to-date");
        setTimeout(() => setUpdateStatus("idle"), 3000);
      }
    } catch {
      setUpdateStatus("error");
      setTimeout(() => setUpdateStatus("idle"), 3000);
    } finally {
      setUpdateChecking(false);
    }
  }, []);

  const handleDesktopUpdate = useCallback(async () => {
    setUpdateStatus("downloading");
    try {
      const dlRes = await fetch(`${API}/api/update/download`);
      if (!dlRes.ok) throw new Error("Download failed");
      setUpdateStatus("installing");
      await fetch(`${API}/api/update/install`, { method: "POST" });
    } catch {
      setUpdateStatus("error");
      setTimeout(() => setUpdateStatus("idle"), 3000);
    }
  }, []);

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
    if (!IS_LOCAL_MODE) warmupBackend(API, setBackendStatus);

    // Show guided tour on first visit
    if (!localStorage.getItem("tour-completed")) {
      setTourStep(0);
    }

    // Check for app updates (non-blocking)
    const dismissed = localStorage.getItem("update-dismissed");
    if (dismissed && Date.now() - parseInt(dismissed) < 24 * 60 * 60 * 1000) {
      setUpdateDismissed(true);
    } else {
      fetch(`${API}/api/version`).then(r => r.json()).then(data => {
        if (data.update_available && data.latest && data.update_url) {
          setUpdateInfo({ latest: data.latest, update_url: data.update_url, installer_url: data.installer_url });
        }
      }).catch(() => {/* no internet or endpoint not ready — skip */});
    }
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
    if (!file.name.match(/\.(xlsx|xls)$/i)) { showToast(t("err.invalid_excel")); return; }
    if (file.size > 10 * 1024 * 1024) { showToast(t("err.file_too_large")); return; }
    if (excelLoading) return; // prevent double upload
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

  // --- MA Reference management ---
  const fetchMaReference = useCallback(async () => {
    try {
      const res = await authFetch(`${API}/api/ma-reference`);
      if (!res.ok) return;
      const data = await res.json();
      setMaReference(data.mappings || []);
    } catch { /* ignore */ }
  }, []);

  const fetchComReference = useCallback(async () => {
    try {
      const res = await authFetch(`${API}/api/com-reference`);
      if (!res.ok) return;
      const data = await res.json();
      setComReference(data.entries || []);
    } catch { /* ignore */ }
  }, []);

  // Load MA + COM reference on mount
  useEffect(() => { fetchMaReference(); fetchComReference(); }, [fetchMaReference, fetchComReference]);

  const uploadMaReference = useCallback(async (file: File) => {
    if (!file.name.match(/\.(xlsx|xls)$/i)) { showToast(t("err.invalid_excel")); return; }
    setMaRefLoading(true);
    try {
      const form = new FormData();
      form.append("file", file);
      const res = await authFetch(`${API}/api/ma-reference/upload`, { method: "POST", body: form });
      if (!res.ok) { const err = await res.json().catch(() => ({})); throw new Error(err.detail || res.statusText); }
      const data = await res.json();
      setMaReference(data.mappings || []);
      showToast(`${data.count} size → MA mappings loaded`, "success");
      if (data.warnings?.length) data.warnings.forEach((w: string) => showToast(w, "warning"));
    } catch (e) { showToast(`Failed to upload MA reference: ${e instanceof Error ? e.message : "Unknown error"}`); }
    setMaRefLoading(false);
  }, [showToast, t]);

  const clearMaReference = useCallback(async () => {
    try {
      const res = await authFetch(`${API}/api/ma-reference`, { method: "DELETE" });
      if (res.ok) { setMaReference([]); showToast("MA reference cleared", "success"); }
    } catch { showToast("Failed to clear MA reference"); }
  }, [showToast]);

  // Add individual MA entry to reference
  const addMaToRef = useCallback(async (size: string, ma: string) => {
    try {
      const form = new FormData();
      form.append("size_normalized", size);
      form.append("size_display", size);
      form.append("ma_number", ma);
      const res = await authFetch(`${API}/api/ma-reference/add`, { method: "POST", body: form });
      if (!res.ok) throw new Error("Failed");
      setAddedMaRefs(prev => new Set(prev).add(ma));
      fetchMaReference();
      showToast(`${ma} added to reference`, "success");
    } catch { showToast("Failed to add MA to reference"); }
  }, [fetchMaReference, showToast]);

  // Add individual COM entry to reference
  const addComToRef = useCallback(async (ma: string, com: number, fabric: string, embroidery: string, frame: string) => {
    try {
      const form = new FormData();
      form.append("ma_number", ma);
      form.append("com_number", String(com));
      form.append("fabric_colour", fabric);
      form.append("embroidery_colour", embroidery);
      form.append("frame_colour", frame);
      const res = await authFetch(`${API}/api/com-reference/add`, { method: "POST", body: form });
      if (!res.ok) throw new Error("Failed");
      const key = `${ma}-${com}`;
      setAddedComRefs(prev => new Set(prev).add(key));
      fetchComReference();
      showToast(`COM ${com} (${ma}) added to reference`, "success");
    } catch { showToast("Failed to add COM to reference"); }
  }, [fetchComReference, showToast]);

  const uploadAssignExcel = useCallback(async (file: File) => {
    if (!file.name.match(/\.(xlsx|xls)$/i)) { showToast(t("err.invalid_excel")); return; }
    if (file.size > 10 * 1024 * 1024) { showToast(t("err.file_too_large")); return; }
    if (assignLoading) return; // prevent double upload
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
      form.append("optimize_heads", String(optimizeHeads));
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
    form.append("optimize_heads", String(optimizeHeads));
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
  }, [sessionId, detectData, columnMapping, optimizeHeads, applyParseData, showToast, t]);

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

  const updateBanner = updateInfo && !updateDismissed ? (
    <div className="w-full flex items-center justify-center gap-3 text-[11px] py-2.5 px-4" style={{ background: "linear-gradient(135deg, var(--accent), #4a5bb8)", color: "white" }}>
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
      <span className="font-medium">{t("update.available")} — v{updateInfo.latest}</span>
      {IS_LOCAL_MODE && updateInfo.installer_url ? (
        <button
          onClick={handleDesktopUpdate}
          disabled={updateStatus === "downloading" || updateStatus === "installing"}
          className="px-3 py-0.5 rounded-full text-[10px] font-semibold transition-all"
          style={{ background: "rgba(255,255,255,0.25)", backdropFilter: "blur(4px)" }}
        >
          {updateStatus === "downloading" ? "⏳ Downloading..." : updateStatus === "installing" ? t("update.installing") : t("update.download")}
        </button>
      ) : (
        <a href={updateInfo.update_url} target="_blank" rel="noopener noreferrer"
          className="px-3 py-0.5 rounded-full text-[10px] font-semibold transition-all"
          style={{ background: "rgba(255,255,255,0.25)", backdropFilter: "blur(4px)" }}>
          {t("update.download")}
        </a>
      )}
      <button onClick={() => { setUpdateDismissed(true); localStorage.setItem("update-dismissed", String(Date.now())); }}
        className="ml-1 w-5 h-5 rounded-full flex items-center justify-center transition-all"
        style={{ background: "rgba(255,255,255,0.15)" }}>✕</button>
    </div>
  ) : null;

  /* ── Session Picker Screen ── */
  if (!sessionStarted) {
    return (
      <div className="min-h-screen flex flex-col" style={{ background: "var(--background)" }}>
        {toast && <Toast message={toast.message} type={toast.type} onClose={() => setToast(null)} />}
        {connectingBanner}
        {updateBanner}

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
                    const timer = getTimeRemaining(s.expires_at);
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
                            {timer && <span className="text-[10px]" style={{ color: timer.color }}>⏱ {timer.text}</span>}
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

        {/* Guided Tour Overlay (session picker) */}
        {tourStep >= 0 && (
          <div className="fixed inset-0 z-50 flex items-center justify-center" style={{ background: "rgba(0,0,0,0.6)", backdropFilter: "blur(4px)" }}>
            <div className="glass-panel p-6 mx-4" style={{ maxWidth: "420px", width: "100%", animation: "fadeSlideIn 0.3s ease" }}>
              <div className="flex items-center gap-1.5 mb-4">
                {[0, 1, 2, 3].map(i => (
                  <div key={i} className="h-1 rounded-full flex-1 transition-all" style={{ background: i <= tourStep ? "var(--accent)" : "var(--border)" }} />
                ))}
              </div>
              <div className="w-10 h-10 rounded-xl flex items-center justify-center mb-3" style={{ background: "var(--accent)", color: "white" }}>
                <span className="text-lg font-bold">{tourStep + 1}</span>
              </div>
              <h3 className="text-base font-semibold mb-1.5">{t(`tour.step${tourStep + 1}.title`)}</h3>
              <p className="text-[13px] leading-relaxed mb-6" style={{ color: "var(--muted)" }}>{t(`tour.step${tourStep + 1}.desc`)}</p>
              <div className="flex items-center justify-between">
                <button className="text-[12px] px-3 py-1.5 rounded-lg transition-colors" style={{ color: "var(--muted)" }}
                  onClick={() => { setTourStep(-1); localStorage.setItem("tour-completed", "1"); }}>{t("tour.skip")}</button>
                <button className="accent-btn !py-2 !px-5 !text-[13px] !min-h-0"
                  onClick={() => { if (tourStep < 3) setTourStep(tourStep + 1); else { setTourStep(-1); localStorage.setItem("tour-completed", "1"); } }}>
                  {tourStep < 3 ? t("tour.next") : t("tour.done")}
                </button>
              </div>
            </div>
          </div>
        )}
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
      {updateBanner}

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
                <div style={{ borderTop: "1px solid var(--border)", marginTop: "4px", paddingTop: "4px" }}>
                  <button
                    className="w-full text-left px-3.5 py-2 text-[11px] transition-colors hover:bg-[var(--surface-hover)] flex items-center gap-2"
                    style={{ color: "var(--muted)" }}
                    disabled={updateChecking}
                    onClick={() => { checkForUpdates(); }}
                  >
                    <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
                    {updateChecking ? t("update.checking") : updateStatus === "up-to-date" ? `✓ ${t("update.latest")}` : updateStatus === "error" ? t("update.no_internet") : t("update.check")}
                  </button>
                  <div className="px-3.5 py-1.5 text-[9px]" style={{ color: "var(--border-strong)" }}>
                    {t("update.version")} {currentVersion}
                  </div>
                </div>
              </div>
            </>
          )}
        </div>
        <div className="flex-1" />
        {!IS_LOCAL_MODE && session?.user && (
          <button onClick={() => { clearAuthToken(); signOut({ callbackUrl: "/login" }); }} className="nav-btn" title={session.user.email || ""}>
            <span className="hidden sm:inline text-[10px]">{session.user.name || "O"}</span>
            <span className="sm:hidden text-[10px]">{(session.user.name || "O").charAt(0).toUpperCase()}</span>
            <span className="text-[9px]" style={{ opacity: 0.6 }}>{t("nav.signout")}</span>
          </button>
        )}
        <button onClick={() => setTourStep(0)} className="nav-btn" title={t("tour.help")}>?</button>
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
            descriptions={[t("cb.step.desc.generate_ma_com"), t("cb.step.desc.upload_order"), t("cb.step.desc.upload_programs"), t("cb.step.desc.export")]}
            onStepClick={handleStepClick}
          />
        </div>

        {/* ── Step 0: Auto-assign MA & COM ── */}
        {assignMode !== "skipped" && !mappingConfirmed && (
          <div ref={step0Ref} className="glass-panel overflow-hidden" style={{ animation: "fadeSlideIn 0.35s ease" }}>
            {/* Header */}
            <div className="px-5 py-4 flex items-center justify-between" style={{ borderBottom: "1px solid var(--border)" }}>
              <div>
                <h3 className="text-sm font-medium">{t("cb.assign.title")}</h3>
                <p className="text-[11px] mt-0.5" style={{ color: "var(--muted)" }}>
                  {assignMode === "pending" ? t("cb.assign.subtitle") : assignActiveField ? `→ ${t("cb.mapping.click_column")} (${t(`cb.assign.field.${assignActiveField}`)})` : t("cb.assign.detect_desc")}
                </p>
              </div>
              {assignMode !== "pending" && (
                <button onClick={() => { setAssignMode("skipped"); }} className="glass-btn text-[10px]">{t("cb.assign.skip")}</button>
              )}
            </div>

            {/* Two-path upload: Generate MA/COM vs I already have MA/COM */}
            {assignMode === "pending" && (
              <div className="p-5">
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  {/* Left card: Generate MA & COM (recommended) */}
                  <div className="glass-panel p-4 flex flex-col" style={{ borderBottom: "2px solid var(--accent)" }}>
                    <div className="flex items-center gap-2 mb-1">
                      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="var(--accent)" strokeWidth="1.5"><path d="M3 3h7v7H3zM14 3h7v7h-7zM3 14h7v7H3zM14 14h7v7h-7z"/></svg>
                      <span className="text-sm font-medium">{t("cb.assign.generate_title")}</span>
                    </div>
                    <span className="text-[9px] font-medium px-1.5 py-0.5 rounded-md mb-3 w-fit" style={{ background: "var(--accent)", color: "white" }}>{t("cb.assign.recommended")}</span>
                    <p className="text-[11px] mb-4" style={{ color: "var(--muted)" }}>{t("cb.assign.generate_desc")}</p>
                    <div className="flex-1" />
                    <div
                      className="drop-zone !py-5 !rounded-xl"
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
                          <svg className="mx-auto mb-2 opacity-25" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" /><polyline points="14 2 14 8 20 8" /><line x1="12" y1="18" x2="12" y2="12" /><polyline points="9 15 12 12 15 15" /></svg>
                          <p className="text-[11px]" style={{ color: "var(--muted)" }}>{t("cb.assign.upload_hint")}</p>
                        </div>
                      )}
                    </div>
                  </div>

                  {/* Right card: I already have MA & COM */}
                  <div className="glass-panel p-4 flex flex-col">
                    <div className="flex items-center gap-2 mb-1">
                      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="var(--success)" strokeWidth="1.5"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>
                      <span className="text-sm font-medium">{t("cb.assign.skip")}</span>
                    </div>
                    <p className="text-[11px] mb-4 mt-2" style={{ color: "var(--muted)" }}>{t("cb.assign.skip_desc")}</p>
                    <div className="flex-1" />
                    <div
                      className="drop-zone !py-5 !rounded-xl"
                      onClick={() => { setAssignMode("skipped"); excelInputRef.current?.click(); }}
                    >
                      <div>
                        <svg className="mx-auto mb-2 opacity-25" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" /><polyline points="14 2 14 8 20 8" /><line x1="12" y1="18" x2="12" y2="12" /><polyline points="9 15 12 12 15 15" /></svg>
                        <p className="text-[11px]" style={{ color: "var(--muted)" }}>{t("cb.assign.upload_hint")}</p>
                      </div>
                    </div>
                  </div>
                </div>

                {/* MA Reference Table */}
                <div className="mt-4 p-3 rounded-lg text-[11px]" style={{ background: "var(--surface)", border: "1px solid var(--border)" }}>
                  <div className="flex items-center justify-between mb-2">
                    <div className="flex items-center gap-2">
                      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke={maReference && maReference.length > 0 ? "var(--success)" : "var(--muted)"} strokeWidth="2"><path d="M4 7V4a2 2 0 0 1 2-2h8.5L20 7.5V20a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2v-3"/><polyline points="14 2 14 8 20 8"/><line x1="2" y1="15" x2="12" y2="15"/><polyline points="9 12 12 15 9 18"/></svg>
                      <span className="font-medium" style={{ color: "var(--foreground)" }}>MA Reference Table</span>
                      {maReference && maReference.length > 0 && (
                        <span className="text-[9px] font-medium px-1.5 py-0.5 rounded-md" style={{ background: "rgba(34,197,94,0.15)", color: "var(--success)" }}>
                          {maReference.length} sizes loaded
                        </span>
                      )}
                    </div>
                    <div className="flex items-center gap-1.5">
                      {maReference && maReference.length > 0 && (
                        <>
                          <button onClick={() => setMaRefExpanded(!maRefExpanded)} className="glass-btn text-[9px] px-2 py-1">
                            {maRefExpanded ? "Hide" : "View"}
                          </button>
                          <button onClick={clearMaReference} className="glass-btn text-[9px] px-2 py-1" style={{ color: "var(--error)" }}>Clear</button>
                        </>
                      )}
                      <button
                        onClick={() => maRefInputRef.current?.click()}
                        className="glass-btn text-[9px] px-2 py-1"
                        disabled={maRefLoading}
                        style={{ color: "var(--accent)" }}
                      >
                        {maRefLoading ? "Uploading..." : maReference && maReference.length > 0 ? "Replace" : "Upload"}
                      </button>
                      <input ref={maRefInputRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={(e) => { if (e.target.files?.[0]) uploadMaReference(e.target.files[0]); e.target.value = ""; }} />
                    </div>
                  </div>
                  {(!maReference || maReference.length === 0) && (
                    <p style={{ color: "var(--muted)" }}>Upload your Size → MA reference Excel to use real MA numbers. Without it, MA1, MA2... will be generated.</p>
                  )}
                  {maRefExpanded && maReference && maReference.length > 0 && (
                    <div className="mt-2">
                      {/* Total counts */}
                      <p className="text-[9px] mb-1.5" style={{ color: "var(--muted)" }}>
                        {t("cb.ref.total").replace("{ma}", String(maReference.length)).replace("{com}", String(comReference?.length || 0))}
                      </p>
                      <div className="overflow-x-auto overflow-y-auto custom-scroll rounded-lg" style={{ border: "1px solid var(--border)", maxHeight: "300px" }}>
                        <table className="text-[10px] border-collapse w-full" style={{ fontFamily: "var(--font-geist-mono)" }}>
                          <thead>
                            <tr style={{ background: "var(--surface)", position: "sticky", top: 0, zIndex: 1 }}>
                              <th className="px-3 py-1.5 text-left font-medium" style={{ borderBottom: "1px solid var(--border)", color: "var(--muted)", width: "20px" }}></th>
                              <th className="px-3 py-1.5 text-left font-medium" style={{ borderBottom: "1px solid var(--border)", color: "var(--muted)" }}>Size</th>
                              <th className="px-3 py-1.5 text-left font-medium" style={{ borderBottom: "1px solid var(--border)", color: "var(--muted)" }}>MA Number</th>
                              <th className="px-3 py-1.5 text-right font-medium" style={{ borderBottom: "1px solid var(--border)", color: "var(--muted)" }}>COMs</th>
                            </tr>
                          </thead>
                          <tbody>
                            {maReference.map((m, i) => {
                              const maComs = comReference?.filter(c => c.ma_number === m.ma_number) || [];
                              const isExpanded = expandedMaRows.has(m.ma_number);
                              return (
                                <>
                                  <tr
                                    key={`ma-${i}`}
                                    style={{ cursor: maComs.length > 0 ? "pointer" : "default" }}
                                    onClick={() => {
                                      if (maComs.length === 0) return;
                                      setExpandedMaRows(prev => {
                                        const next = new Set(prev);
                                        if (next.has(m.ma_number)) next.delete(m.ma_number);
                                        else next.add(m.ma_number);
                                        return next;
                                      });
                                    }}
                                  >
                                    <td className="px-2 py-1 text-center" style={{ borderBottom: (!isExpanded && i < maReference.length - 1) ? "1px solid var(--border)" : isExpanded ? "none" : "none", color: "var(--muted)", fontSize: "8px" }}>
                                      {maComs.length > 0 && (isExpanded ? "▼" : "▶")}
                                    </td>
                                    <td className="px-3 py-1" style={{ borderBottom: (!isExpanded && i < maReference.length - 1) ? "1px solid var(--border)" : isExpanded ? "none" : "none" }}>{m.size_display}</td>
                                    <td className="px-3 py-1 font-medium" style={{ borderBottom: (!isExpanded && i < maReference.length - 1) ? "1px solid var(--border)" : isExpanded ? "none" : "none", color: "#f59e0b" }}>{m.ma_number}</td>
                                    <td className="px-3 py-1 text-right" style={{ borderBottom: (!isExpanded && i < maReference.length - 1) ? "1px solid var(--border)" : isExpanded ? "none" : "none" }}>
                                      {maComs.length > 0 ? (
                                        <span className="inline-block px-1.5 py-0.5 rounded text-[9px]" style={{ background: "rgba(38,57,122,0.1)", color: "var(--accent)" }}>{maComs.length}</span>
                                      ) : (
                                        <span className="text-[9px]" style={{ color: "var(--muted)" }}>{t("cb.ref.no_coms")}</span>
                                      )}
                                    </td>
                                  </tr>
                                  {isExpanded && maComs.length > 0 && (
                                    <tr key={`coms-${i}`}>
                                      <td colSpan={4} style={{ padding: 0, borderBottom: i < maReference.length - 1 ? "1px solid var(--border)" : "none" }}>
                                        <div className="ml-6 mr-2 mb-1.5 rounded" style={{ background: "var(--surface)", border: "1px solid var(--border)" }}>
                                          <table className="text-[9px] border-collapse w-full">
                                            <thead>
                                              <tr>
                                                <th className="px-2 py-1 text-left font-medium" style={{ color: "var(--muted)", borderBottom: "1px solid var(--border)" }}>COM</th>
                                                <th className="px-2 py-1 text-left font-medium" style={{ color: "var(--muted)", borderBottom: "1px solid var(--border)" }}>Fabric</th>
                                                <th className="px-2 py-1 text-left font-medium" style={{ color: "var(--muted)", borderBottom: "1px solid var(--border)" }}>Frame</th>
                                                <th className="px-2 py-1 text-left font-medium" style={{ color: "var(--muted)", borderBottom: "1px solid var(--border)" }}>Embroidery</th>
                                              </tr>
                                            </thead>
                                            <tbody>
                                              {maComs.map((c, ci) => (
                                                <tr key={ci}>
                                                  <td className="px-2 py-0.5 font-medium" style={{ color: "var(--accent)", borderBottom: ci < maComs.length - 1 ? "1px solid var(--border)" : "none" }}>{c.com_number}</td>
                                                  <td className="px-2 py-0.5" style={{ borderBottom: ci < maComs.length - 1 ? "1px solid var(--border)" : "none" }}>{c.fabric_colour}</td>
                                                  <td className="px-2 py-0.5" style={{ borderBottom: ci < maComs.length - 1 ? "1px solid var(--border)" : "none" }}>{c.frame_colour}</td>
                                                  <td className="px-2 py-0.5" style={{ borderBottom: ci < maComs.length - 1 ? "1px solid var(--border)" : "none" }}>{c.embroidery_colour}</td>
                                                </tr>
                                              ))}
                                            </tbody>
                                          </table>
                                        </div>
                                      </td>
                                    </tr>
                                  )}
                                </>
                              );
                            })}
                          </tbody>
                        </table>
                      </div>
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

            {/* Results — Grouped Card Layout */}
            {assignMode === "result" && assignResult && (
              <div className="p-5">
                <h4 className="text-sm font-medium mb-3">{t("cb.assign.result_title")}</h4>

                {/* Summary stats */}
                <div className="grid grid-cols-3 gap-3 mb-4">
                  <div className="stat-card"><div className="stat-number">{assignResult.ma_summary.length}</div><div className="stat-label">{t("cb.assign.ma_groups")}</div></div>
                  <div className="stat-card"><div className="stat-number">{assignResult.com_summary.length}</div><div className="stat-label">{t("cb.assign.com_groups")}</div></div>
                  <div className="stat-card"><div className="stat-number">{assignResult.assignments_count}</div><div className="stat-label">{t("cb.assign.total_rows")}</div></div>
                </div>

                {/* Legend */}
                <div className="flex flex-wrap gap-4 mb-4 text-[10px] px-1" style={{ color: "var(--muted)" }}>
                  <span><span className="inline-block w-2 h-2 rounded-full mr-1" style={{ background: "#f59e0b" }} />{t("cb.assign.legend_ma")}</span>
                  <span><span className="inline-block w-2 h-2 rounded-full mr-1" style={{ background: "var(--accent)" }} />{t("cb.assign.legend_com")}</span>
                </div>

                {/* Grouped MA → COM cards */}
                <div className="flex flex-col gap-3 mb-4">
                  {assignResult.ma_summary.map((ma) => {
                    const comRows = assignResult.com_summary.filter(c => c.ma === ma.ma);
                    return (
                      <div key={ma.ma} className="glass-panel overflow-hidden">
                        {/* MA Header */}
                        <div className="px-4 py-3 flex items-center gap-3" style={{ borderBottom: "1px solid var(--border)" }}>
                          <span className="text-[10px] font-bold px-2 py-0.5 rounded-md" style={{
                            background: /^MA\d+$/.test(ma.ma) && maReference && maReference.length > 0 ? "rgba(239,68,68,0.15)" : "rgba(245,158,11,0.15)",
                            color: /^MA\d+$/.test(ma.ma) && maReference && maReference.length > 0 ? "#ef4444" : "#f59e0b",
                          }}>{ma.ma}{/^MA\d+$/.test(ma.ma) && maReference && maReference.length > 0 && <span title="Not found in MA reference"> ⚠</span>}</span>
                          <span className="text-[12px] font-medium">{ma.size}</span>
                          <div className="flex-1" />
                          <span className="text-[10px]" style={{ color: "var(--muted)" }}>{ma.count} {t("cb.assign.entries")}</span>
                          {ma.is_new && !addedMaRefs.has(ma.ma) && (
                            <button
                              className="glass-btn text-[9px] px-2 py-0.5 ml-1"
                              style={{ color: "var(--accent)", borderColor: "var(--accent)" }}
                              onClick={(e) => { e.stopPropagation(); addMaToRef(ma.size, ma.ma); }}
                            >{t("cb.ref.add_ma")}</button>
                          )}
                          {ma.is_new && addedMaRefs.has(ma.ma) && (
                            <span className="text-[9px] px-2 py-0.5 ml-1 font-medium" style={{ color: "var(--success)" }}>{t("cb.ref.added")} ✓</span>
                          )}
                        </div>
                        {/* COM Table */}
                        <div className="overflow-x-auto">
                          <table className="text-[11px] border-collapse w-full">
                            <thead>
                              <tr style={{ background: "var(--surface)" }}>
                                <th className="px-3 py-1.5 text-left font-medium" style={{ borderBottom: "1px solid var(--border)", color: "var(--muted)", width: "50px" }}>{t("cb.assign.com_label")}</th>
                                <th className="px-3 py-1.5 text-left font-medium" style={{ borderBottom: "1px solid var(--border)", color: "var(--muted)", background: "rgba(59,130,246,0.04)" }}>{t("cb.assign.fabric_col")}</th>
                                <th className="px-3 py-1.5 text-left font-medium" style={{ borderBottom: "1px solid var(--border)", color: "var(--muted)", background: "rgba(168,85,247,0.04)" }}>{t("cb.assign.frame_col")}</th>
                                <th className="px-3 py-1.5 text-left font-medium" style={{ borderBottom: "1px solid var(--border)", color: "var(--muted)", background: "rgba(34,197,94,0.04)" }}>{t("cb.assign.embroidery_col")}</th>
                                <th className="px-3 py-1.5 text-right font-medium" style={{ borderBottom: "1px solid var(--border)", color: "var(--muted)", width: "60px" }}>{t("cb.assign.entries")}</th>
                                <th className="px-2 py-1.5 text-center font-medium" style={{ borderBottom: "1px solid var(--border)", color: "var(--muted)", width: "70px" }}></th>
                              </tr>
                            </thead>
                            <tbody>
                              {comRows.map((c, i) => {
                                const comKey = `${ma.ma}-${c.com}`;
                                const isAdded = addedComRefs.has(comKey);
                                return (
                                <tr key={i}>
                                  <td className="px-3 py-1.5 font-medium" style={{ color: "var(--accent)", borderBottom: i < comRows.length - 1 ? "1px solid var(--border)" : "none" }}>
                                    <span className="inline-block px-1.5 py-0.5 rounded text-[10px]" style={{ background: "rgba(38,57,122,0.1)" }}>{c.com}</span>
                                  </td>
                                  <td className="px-3 py-1.5" style={{ borderBottom: i < comRows.length - 1 ? "1px solid var(--border)" : "none", background: "rgba(59,130,246,0.03)" }}>{c.fabric_colour}</td>
                                  <td className="px-3 py-1.5" style={{ borderBottom: i < comRows.length - 1 ? "1px solid var(--border)" : "none", background: "rgba(168,85,247,0.03)" }}>{c.frame_colour}</td>
                                  <td className="px-3 py-1.5" style={{ borderBottom: i < comRows.length - 1 ? "1px solid var(--border)" : "none", background: "rgba(34,197,94,0.03)" }}>{c.embroidery_colour}</td>
                                  <td className="px-3 py-1.5 text-right" style={{ color: "var(--muted)", borderBottom: i < comRows.length - 1 ? "1px solid var(--border)" : "none" }}>{c.count}</td>
                                  <td className="px-2 py-1.5 text-center" style={{ borderBottom: i < comRows.length - 1 ? "1px solid var(--border)" : "none" }}>
                                    {c.is_new && !isAdded && (
                                      <button
                                        className="glass-btn text-[9px] px-1.5 py-0.5"
                                        style={{ color: "var(--accent)", borderColor: "var(--accent)" }}
                                        onClick={() => addComToRef(ma.ma, c.com, c.fabric_colour, c.embroidery_colour, c.frame_colour)}
                                      >{t("cb.ref.add_to_ref")}</button>
                                    )}
                                    {c.is_new && isAdded && (
                                      <span className="text-[9px] font-medium" style={{ color: "var(--success)" }}>{t("cb.ref.added")} ✓</span>
                                    )}
                                  </td>
                                </tr>
                                );
                              })}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    );
                  })}
                </div>

                {/* Action buttons */}
                <div className="flex flex-col sm:flex-row gap-2">
                  <button
                    className="flex-1 text-[11px] py-2.5 rounded-[10px] font-medium transition-all flex items-center justify-center gap-2"
                    style={{ border: "1.5px solid var(--accent)", color: "var(--accent)", background: "transparent" }}
                    onMouseEnter={(e) => { e.currentTarget.style.background = "rgba(38,57,122,0.06)"; }}
                    onMouseLeave={(e) => { e.currentTarget.style.background = "transparent"; }}
                    onClick={downloadAssignedExcel}
                    disabled={assignDownloading}
                  >
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
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
        <div ref={step1Ref} className="grid grid-cols-1 md:grid-cols-2 gap-3" style={{ animation: "slideUp 0.4s ease 0.05s forwards", opacity: 0 }}>
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
                    className="flex items-center justify-center rounded-md text-[10px] transition-all"
                    style={{
                      color: removeExcelConfirm ? "white" : "var(--muted)",
                      background: removeExcelConfirm ? "var(--danger)" : "var(--surface)",
                      minWidth: removeExcelConfirm ? "80px" : "20px",
                      height: "20px",
                      padding: removeExcelConfirm ? "0 6px" : "0",
                      fontSize: removeExcelConfirm ? "9px" : "10px",
                      fontWeight: removeExcelConfirm ? 500 : 400,
                    }}
                    onClick={(e) => {
                      e.stopPropagation();
                      if (removeExcelConfirm) { removeExcel(); setRemoveExcelConfirm(false); }
                      else { setRemoveExcelConfirm(true); setTimeout(() => setRemoveExcelConfirm(false), 3000); }
                    }}
                    title={removeExcelConfirm ? t("err.remove_excel_confirm") : t("cb.excel.remove")}
                  >{removeExcelConfirm ? t("cb.session.delete_confirm_label") : "✕"}</button>
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
                <label className="flex items-center gap-2 text-[11px] group relative">
                  <span style={{ color: "var(--muted)" }}>{t("cb.settings.vgap")}</span>
                  <span className="cursor-help" style={{ color: "var(--muted)", fontSize: "10px" }} title={t("cb.settings.vgap.help")}>&#9432;</span>
                  <input type="number" value={gapMm} onChange={(e) => setGapMm(Number(e.target.value))} min={0} max={20} step={0.5} className="w-14 text-center text-[11px] px-2 py-1 rounded-lg bg-transparent" style={{ border: "1px solid var(--border)", color: "var(--foreground)" }} />
                  <span style={{ color: "var(--muted)" }}>mm</span>
                </label>
                <label className="flex items-center gap-2 text-[11px] group relative">
                  <span style={{ color: "var(--muted)" }}>{t("cb.settings.cgap")}</span>
                  <span className="cursor-help" style={{ color: "var(--muted)", fontSize: "10px" }} title={t("cb.settings.cgap.help")}>&#9432;</span>
                  <input type="number" value={columnGapMm} onChange={(e) => setColumnGapMm(Number(e.target.value))} min={0} max={30} step={0.5} className="w-14 text-center text-[11px] px-2 py-1 rounded-lg bg-transparent" style={{ border: "1px solid var(--border)", color: "var(--foreground)" }} />
                  <span style={{ color: "var(--muted)" }}>mm</span>
                </label>
                <label className="flex items-center gap-2 text-[11px] cursor-pointer">
                  <input type="checkbox" checked={optimizeHeads} onChange={(e) => setOptimizeHeads(e.target.checked)} className="custom-checkbox" />
                  <span style={{ color: "var(--muted)" }}>{t("cb.settings.2head")}</span>
                  <span className="cursor-help" style={{ color: "var(--muted)", fontSize: "10px" }} title={t("cb.settings.2head.help")}>&#9432;</span>
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
              value={parseData.optimized_slots ?? parseData.total_slots} label={t("cb.stats.slots")} delay={0.2}
              icon={<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><polygon points="12 2 2 7 12 12 22 7 12 2"/><polyline points="2 17 12 22 22 17"/><polyline points="2 12 12 17 22 12"/></svg>}
            />
            {(parseData.slots_saved ?? 0) > 0 && (<>
              <span className="stat-arrow hidden sm:block">→</span>
              <StatCard
                value={parseData.slots_saved!} label={t("cb.stats.slots_saved")} delay={0.25}
                icon={<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#16a34a" strokeWidth="2"><polyline points="20 6 9 17 4 12"/></svg>}
              />
            </>)}
          </div>
        )}

        {/* Two-Panel: Combo List + Slot Preview */}
        {parseData && parseData.groups.length > 0 && (
          <div ref={step3Ref} className="flex-1 grid grid-cols-1 lg:grid-cols-[1fr_360px] gap-3 min-h-0" style={{ animation: "slideUp 0.4s ease 0.15s forwards", opacity: 0 }}>
            <div className="glass-panel overflow-hidden flex flex-col min-h-[200px] sm:min-h-[300px]">
              <div className="flex items-center justify-between px-4 py-2.5 shrink-0" style={{ borderBottom: "1px solid var(--border)" }}>
                <div className="flex flex-col">
                  <span className="text-xs font-medium">{t("cb.files.title")} <span className="font-normal" style={{ color: "var(--muted)" }}>{selectedCombos.size}/{totalCombos}</span></span>
                  <span style={{ fontSize: "8px", color: "var(--muted)", opacity: 0.7 }}>{t("cb.files.sort_order")}</span>
                </div>
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
                          {combo.head_mode === "2-HEAD" && <span className="text-[9px] px-1.5 py-0.5 rounded font-bold" style={{ background: "rgba(34,197,94,0.15)", color: "#16a34a", border: "1px solid rgba(34,197,94,0.3)" }}>2-HEAD</span>}
                          {combo.head_mode === "1-HEAD" && <span className="text-[9px] px-1.5 py-0.5 rounded" style={{ background: "rgba(168,85,247,0.1)", color: "#a855f7", border: "1px solid rgba(168,85,247,0.2)" }}>1-HEAD</span>}
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
              {dstData && dstData.missing_programs.length > 0 && !exporting && (
                <div className="mt-2 px-3 py-2 rounded-lg text-[11px] flex items-center gap-2" style={{ background: "rgba(245,158,11,0.1)", border: "1px solid rgba(245,158,11,0.3)", color: "var(--warning)" }}>
                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>
                  <span>{dstData.missing_programs.length} DST file{dstData.missing_programs.length !== 1 ? "s" : ""} missing: {dstData.missing_programs.slice(0, 8).join(", ")}{dstData.missing_programs.length > 8 ? " ..." : ""}</span>
                </div>
              )}
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

      {/* Guided Tour Overlay */}
      {tourStep >= 0 && (
        <div className="fixed inset-0 z-50 flex items-center justify-center" style={{ background: "rgba(0,0,0,0.6)", backdropFilter: "blur(4px)" }}>
          <div className="glass-panel p-6 mx-4" style={{ maxWidth: "420px", width: "100%", animation: "fadeSlideIn 0.3s ease" }}>
            {/* Step indicator */}
            <div className="flex items-center gap-1.5 mb-4">
              {[0, 1, 2, 3].map(i => (
                <div key={i} className="h-1 rounded-full flex-1 transition-all" style={{ background: i <= tourStep ? "var(--accent)" : "var(--border)" }} />
              ))}
            </div>

            {/* Step icon */}
            <div className="w-10 h-10 rounded-xl flex items-center justify-center mb-3" style={{ background: "var(--accent)", color: "white" }}>
              <span className="text-lg font-bold">{tourStep + 1}</span>
            </div>

            {/* Step content */}
            <h3 className="text-base font-semibold mb-1.5">
              {t(`tour.step${tourStep + 1}.title`)}
            </h3>
            <p className="text-[13px] leading-relaxed mb-6" style={{ color: "var(--muted)" }}>
              {t(`tour.step${tourStep + 1}.desc`)}
            </p>

            {/* Buttons */}
            <div className="flex items-center justify-between">
              <button
                className="text-[12px] px-3 py-1.5 rounded-lg transition-colors"
                style={{ color: "var(--muted)" }}
                onClick={() => { setTourStep(-1); localStorage.setItem("tour-completed", "1"); }}
              >
                {t("tour.skip")}
              </button>
              <button
                className="accent-btn !py-2 !px-5 !text-[13px] !min-h-0"
                onClick={() => {
                  if (tourStep < 3) { setTourStep(tourStep + 1); }
                  else { setTourStep(-1); localStorage.setItem("tour-completed", "1"); }
                }}
              >
                {tourStep < 3 ? t("tour.next") : t("tour.done")}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
