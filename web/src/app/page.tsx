"use client";

import Link from "next/link";
import Image from "next/image";
import { useTheme } from "./theme-provider";
import { useLanguage } from "./i18n";

export default function Home() {
  const { theme, toggle } = useTheme();
  const { lang, toggle: toggleLang, t } = useLanguage();

  return (
    <div className="min-h-screen flex flex-col items-center justify-center px-6 relative">
      {/* Top-right controls */}
      <div className="fixed top-6 right-6 flex items-center gap-2 z-50" style={{ animation: "fadeIn 0.5s ease 0.4s forwards", opacity: 0 }}>
        <button onClick={toggleLang} className="nav-btn">
          {lang === "en" ? "TH" : "EN"}
        </button>
        <button onClick={toggle} className="theme-toggle">{theme === "light" ? "☾" : "☀"}</button>
      </div>

      {/* Header */}
      <div className="text-center mb-8 sm:mb-14" style={{ animation: "slideUp 0.6s ease forwards" }}>
        <div className="flex items-center justify-center gap-4 mb-4">
          <Image src="/micro-logo.svg" alt="Micro" width={100} height={28} className="micro-logo" />
          <div style={{ width: "1px", height: "28px", background: "var(--border-strong)", opacity: 0.4 }} />
          <div className="flex items-center gap-2.5">
            <Image src="/ossia-mark.svg?v3" alt="" width={30} height={21} className="micro-logo" />
            <span className="text-xl font-normal" style={{ color: "var(--foreground)", letterSpacing: "-0.035em" }}>ossia</span>
          </div>
        </div>
        <p className="text-[11px] tracking-widest uppercase" style={{ color: "var(--border-strong)", letterSpacing: "0.2em" }}>
          {t("home.subtitle")}
        </p>
      </div>

      {/* Product Grid */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-3 sm:gap-5 max-w-3xl w-full" style={{ animation: "slideUp 0.6s ease 0.1s forwards", opacity: 0 }}>
        <Link href="/stacker" className="block group">
          <div className="glass-card p-7 cursor-pointer">
            <div className="w-10 h-10 rounded-xl flex items-center justify-center mb-5" style={{ background: "var(--accent)", color: "white" }}>
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><polygon points="12 2 2 7 12 12 22 7 12 2"/><polyline points="2 17 12 22 22 17"/><polyline points="2 12 12 17 22 12"/></svg>
            </div>
            <h2 className="text-base font-medium mb-1.5" style={{ color: "var(--foreground)" }}>{t("home.combo.title")}</h2>
            <p className="text-xs leading-relaxed mb-4" style={{ color: "var(--muted)" }}>{t("home.combo.desc")}</p>
            <span className="text-[10px] font-medium px-2.5 py-1 rounded-full uppercase tracking-wider" style={{ background: "var(--accent-glow)", color: "var(--accent)" }}>{t("home.combo.available")}</span>
          </div>
        </Link>

        <div className="glass-card p-7 cursor-not-allowed" style={{ opacity: 0.38 }}>
          <div className="w-10 h-10 rounded-xl flex items-center justify-center mb-5" style={{ background: "var(--surface)", color: "var(--muted)" }}>
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><line x1="4" y1="9" x2="20" y2="9" /><line x1="4" y1="15" x2="20" y2="15" /><line x1="10" y1="3" x2="8" y2="21" /><line x1="16" y1="3" x2="14" y2="21" /></svg>
          </div>
          <h2 className="text-base font-medium mb-1.5">{t("home.stitch.title")}</h2>
          <p className="text-xs leading-relaxed mb-4" style={{ color: "var(--muted)" }}>{t("home.stitch.desc")}</p>
          <span className="flex items-center gap-1.5 w-fit text-[10px] font-medium px-2.5 py-1 rounded-full" style={{ background: "var(--surface)", color: "var(--muted)", border: "1px solid var(--border)" }}>
            <svg width="9" height="9" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><rect x="3" y="11" width="18" height="11" rx="2" ry="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></svg>
            {t("home.coming_soon")}
          </span>
        </div>

        <div className="glass-card p-7 cursor-not-allowed" style={{ opacity: 0.38 }}>
          <div className="w-10 h-10 rounded-xl flex items-center justify-center mb-5" style={{ background: "var(--surface)", color: "var(--muted)" }}>
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="11" cy="11" r="8" /><line x1="21" y1="21" x2="16.65" y2="16.65" /><line x1="8" y1="11" x2="14" y2="11" /><line x1="11" y1="8" x2="11" y2="14" /></svg>
          </div>
          <h2 className="text-base font-medium mb-1.5">{t("home.batch.title")}</h2>
          <p className="text-xs leading-relaxed mb-4" style={{ color: "var(--muted)" }}>{t("home.batch.desc")}</p>
          <span className="flex items-center gap-1.5 w-fit text-[10px] font-medium px-2.5 py-1 rounded-full" style={{ background: "var(--surface)", color: "var(--muted)", border: "1px solid var(--border)" }}>
            <svg width="9" height="9" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><rect x="3" y="11" width="18" height="11" rx="2" ry="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></svg>
            {t("home.coming_soon")}
          </span>
        </div>
      </div>

      {/* Footer removed — branding already in header */}
    </div>
  );
}
