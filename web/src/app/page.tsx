"use client";

import Link from "next/link";
import Image from "next/image";
import { useTheme } from "./theme-provider";

export default function Home() {
  const { theme, toggle } = useTheme();

  return (
    <div className="min-h-screen flex flex-col items-center justify-center px-6 relative">
      {/* Theme toggle */}
      <button
        onClick={toggle}
        className="theme-toggle fixed top-6 right-6"
        style={{ zIndex: 50, animation: "fadeIn 0.5s ease 0.4s forwards", opacity: 0 }}
      >
        {theme === "light" ? "☾" : "☀"}
      </button>

      {/* Header */}
      <div
        className="text-center mb-14"
        style={{ animation: "slideUp 0.6s ease forwards" }}
      >
        <div className="flex items-center justify-center gap-4 mb-4">
          <Image
            src="/micro-logo.svg"
            alt="Micro"
            width={100}
            height={28}
            className="micro-logo"
          />
          <div style={{ width: "1px", height: "28px", background: "var(--border-strong)", opacity: 0.4 }} />
          <div className="flex items-center gap-2.5">
            <Image
              src="/ossia-mark.svg?v2"
              alt=""
              width={26}
              height={26}
              className="micro-logo"
            />
            <span
              className="text-xl font-semibold"
              style={{ color: "var(--foreground)", letterSpacing: "-0.04em" }}
            >
              ossia
            </span>
          </div>
        </div>
        <p
          className="text-[11px] tracking-widest uppercase"
          style={{ color: "var(--border-strong)", letterSpacing: "0.2em" }}
        >
          automation tools
        </p>
      </div>

      {/* Product Grid */}
      <div
        className="grid grid-cols-1 md:grid-cols-3 gap-5 max-w-3xl w-full"
        style={{ animation: "slideUp 0.6s ease 0.1s forwards", opacity: 0 }}
      >
        <Link href="/combo-builder" className="block group">
          <div className="glass-card p-7 cursor-pointer">
            <div
              className="w-10 h-10 rounded-xl flex items-center justify-center mb-5"
              style={{ background: "var(--accent)", color: "white" }}
            >
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                <path d="M13 2L3 14h9l-1 8 10-12h-9l1-8z" />
              </svg>
            </div>
            <h2 className="text-base font-semibold mb-1.5" style={{ color: "var(--foreground)" }}>
              Combo Builder
            </h2>
            <p className="text-xs leading-relaxed mb-4" style={{ color: "var(--muted)" }}>
              Combine embroidery name programs into production-ready combo files.
            </p>
            <span
              className="text-[10px] font-semibold px-2.5 py-1 rounded-full uppercase tracking-wider"
              style={{ background: "var(--accent-glow)", color: "var(--accent)" }}
            >
              Available
            </span>
          </div>
        </Link>

        <div className="glass-card p-7 opacity-30 cursor-not-allowed">
          <div className="w-10 h-10 rounded-xl flex items-center justify-center mb-5" style={{ background: "var(--surface)", color: "var(--muted)" }}>
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <line x1="4" y1="9" x2="20" y2="9" /><line x1="4" y1="15" x2="20" y2="15" /><line x1="10" y1="3" x2="8" y2="21" /><line x1="16" y1="3" x2="14" y2="21" />
            </svg>
          </div>
          <h2 className="text-base font-semibold mb-1.5">Stitch Count Predictor</h2>
          <p className="text-xs leading-relaxed mb-4" style={{ color: "var(--muted)" }}>Predict stitch counts and production time from design files.</p>
          <span className="text-[10px] font-medium px-2.5 py-1 rounded-full" style={{ background: "var(--surface)", color: "var(--muted)" }}>Coming Soon</span>
        </div>

        <div className="glass-card p-7 opacity-30 cursor-not-allowed">
          <div className="w-10 h-10 rounded-xl flex items-center justify-center mb-5" style={{ background: "var(--surface)", color: "var(--muted)" }}>
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <circle cx="11" cy="11" r="8" /><line x1="21" y1="21" x2="16.65" y2="16.65" /><line x1="8" y1="11" x2="14" y2="11" /><line x1="11" y1="8" x2="11" y2="14" />
            </svg>
          </div>
          <h2 className="text-base font-semibold mb-1.5">Batch Inspector</h2>
          <p className="text-xs leading-relaxed mb-4" style={{ color: "var(--muted)" }}>Validate and inspect DST/NGS files before production runs.</p>
          <span className="text-[10px] font-medium px-2.5 py-1 rounded-full" style={{ background: "var(--surface)", color: "var(--muted)" }}>Coming Soon</span>
        </div>
      </div>

      <p
        className="mt-16 text-[10px] uppercase tracking-widest"
        style={{ color: "var(--border-strong)", animation: "fadeIn 0.6s ease 0.4s forwards", opacity: 0 }}
      >
        Production tools for Micro Embroidery Co.
      </p>
    </div>
  );
}
