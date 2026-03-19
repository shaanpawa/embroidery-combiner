"use client";

import { signIn } from "next-auth/react";
import { useSearchParams } from "next/navigation";
import { Suspense, useState } from "react";
import Image from "next/image";

function LoginContent() {
  const searchParams = useSearchParams();
  const error = searchParams.get("error");
  const callbackUrl = searchParams.get("callbackUrl") || "/";
  const [password, setPassword] = useState("");
  const [loading, setLoading] = useState(false);
  const [authError, setAuthError] = useState(error);

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
      {/* Logo */}
      <div
        className="text-center mb-10"
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
          <div
            style={{
              width: "1px",
              height: "28px",
              background: "var(--border-strong)",
              opacity: 0.4,
            }}
          />
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

      {/* Sign-in card */}
      <div
        className="glass-card p-8 sm:p-10 w-full max-w-sm text-center"
        style={{ animation: "slideUp 0.6s ease 0.1s forwards", opacity: 0 }}
      >
        <h1
          className="text-lg font-semibold mb-2"
          style={{ color: "var(--foreground)" }}
        >
          Welcome back
        </h1>
        <p className="text-xs mb-6" style={{ color: "var(--muted)" }}>
          Sign in to access Micro Automation
        </p>

        {/* Error message */}
        {authError && (
          <div
            className="mb-5 px-4 py-3 rounded-xl text-xs font-medium"
            style={{
              background: "rgba(220, 38, 38, 0.08)",
              color: "var(--danger)",
              border: "1px solid rgba(220, 38, 38, 0.15)",
            }}
          >
            {authError === "AccessDenied"
              ? "Access denied. Your email may not be whitelisted."
              : authError === "CredentialsSignin"
              ? "Incorrect password. Please try again."
              : "Something went wrong. Please try again."}
          </div>
        )}

        {/* Password form */}
        <form onSubmit={handlePasswordLogin} className="mb-4">
          <input
            type="password"
            value={password}
            onChange={(e) => { setPassword(e.target.value); setAuthError(null); }}
            placeholder="Enter password"
            className="w-full text-sm px-4 py-3 rounded-xl bg-transparent mb-3"
            style={{ border: "1px solid var(--border)", color: "var(--foreground)" }}
            autoFocus
          />
          <button
            type="submit"
            className="accent-btn w-full"
            disabled={!password.trim() || loading}
          >
            {loading ? "Signing in..." : "Sign In"}
          </button>
        </form>

        {/* Google sign-in (optional) */}
        <button
          onClick={() => signIn("google", { callbackUrl })}
          className="glass-btn w-full flex items-center justify-center gap-3 py-3 px-5 text-xs"
          style={{
            borderRadius: "12px",
            color: "var(--muted)",
            border: "1px solid var(--glass-border)",
            background: "var(--glass)",
            backdropFilter: "blur(12px)",
          }}
        >
          <svg width="16" height="16" viewBox="0 0 24 24">
            <path
              d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92a5.06 5.06 0 0 1-2.2 3.32v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.1z"
              fill="#4285F4"
            />
            <path
              d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z"
              fill="#34A853"
            />
            <path
              d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.07H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.93l2.85-2.22.81-.62z"
              fill="#FBBC05"
            />
            <path
              d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.07l3.66 2.84c.87-2.6 3.3-4.53 6.16-4.53z"
              fill="#EA4335"
            />
          </svg>
          or sign in with Google
        </button>
      </div>

      {/* Footer */}
      <p
        className="mt-12 text-[10px] uppercase tracking-widest"
        style={{
          color: "var(--border-strong)",
          animation: "fadeIn 0.6s ease 0.4s forwards",
          opacity: 0,
        }}
      >
        Micro Automation by Ossia
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
