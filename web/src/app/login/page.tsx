"use client";

import { signIn } from "next-auth/react";
import { useSearchParams, useRouter } from "next/navigation";
import { Suspense, useState, useEffect } from "react";
import Image from "next/image";
import { useLanguage } from "../i18n";
import { useTheme } from "../theme-provider";

const IS_LOCAL_MODE = process.env.NEXT_PUBLIC_LOCAL_MODE === "true";

function LoginContent() {
  const router = useRouter();

  // In local mode, skip login entirely
  useEffect(() => {
    if (IS_LOCAL_MODE) router.replace("/stacker");
  }, [router]);
  if (IS_LOCAL_MODE) return null;
  const searchParams = useSearchParams();
  const error = searchParams.get("error");
  const callbackUrl = searchParams.get("callbackUrl") || "/";
  const [password, setPassword] = useState("");
  const [loading, setLoading] = useState(false);
  const [authError, setAuthError] = useState(error);
  const { lang, toggle, t } = useLanguage();
  const { theme, toggle: toggleTheme } = useTheme();

  const handlePasswordLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!password.trim()) return;
    setLoading(true);
    setAuthError(null);
    const result = await signIn("credentials", {
      password,
      redirect: false,
      callbackUrl,
    });
    if (result?.error) {
      setAuthError("CredentialsSignin");
      setLoading(false);
    } else if (result?.url) {
      window.location.href = result.url;
    }
  };

  return (
    <div className="min-h-screen flex flex-col items-center justify-center px-6">
      {/* Top-right controls */}
      <div className="fixed top-4 right-4 z-50 flex items-center gap-2">
        <button onClick={toggle} className="nav-btn">
          {lang === "en" ? "TH" : "EN"}
        </button>
        <button onClick={toggleTheme} className="theme-toggle">{theme === "light" ? "☾" : "☀"}</button>
      </div>

      {/* Logo */}
      <div className="text-center mb-10" style={{ animation: "slideUp 0.6s ease forwards" }}>
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

      {/* Sign-in card */}
      <div className="glass-card p-8 sm:p-10 w-full max-w-sm text-center" style={{ animation: "slideUp 0.6s ease 0.1s forwards", opacity: 0 }}>
        <h1 className="text-lg font-medium mb-2" style={{ color: "var(--foreground)" }}>{t("login.welcome")}</h1>
        <p className="text-xs mb-6" style={{ color: "var(--muted)" }}>{t("login.subtitle")}</p>

        {authError && (
          <div className="mb-5 px-4 py-3 rounded-xl text-xs font-medium" style={{ background: "rgba(220, 38, 38, 0.08)", color: "var(--danger)", border: "1px solid rgba(220, 38, 38, 0.15)" }}>
            {authError === "AccessDenied" ? t("login.error.denied") : authError === "CredentialsSignin" ? t("login.error.password") : t("login.error.generic")}
          </div>
        )}

        <form onSubmit={handlePasswordLogin} className="mb-4">
          <input
            type="password"
            value={password}
            onChange={(e) => { setPassword(e.target.value); setAuthError(null); }}
            placeholder={t("login.password")}
            className="w-full text-sm px-4 py-3 rounded-xl bg-transparent mb-3"
            style={{ border: "1px solid var(--border)", color: "var(--foreground)" }}
            autoFocus
          />
          <button type="submit" className="accent-btn w-full" disabled={!password.trim() || loading}>
            {loading ? t("login.signing_in") : t("login.signin")}
          </button>
        </form>

        {/* Google sign-in — hidden until Google OAuth is configured */}
      </div>

      <p className="mt-12 text-[10px] uppercase tracking-widest" style={{ color: "var(--border-strong)", animation: "fadeIn 0.6s ease 0.4s forwards", opacity: 0 }}>
        {t("login.footer")}
      </p>
    </div>
  );
}

export default function LoginPage() {
  return (
    <Suspense>
      <LoginContent />
    </Suspense>
  );
}
