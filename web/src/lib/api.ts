const IS_LOCAL_MODE = process.env.NEXT_PUBLIC_LOCAL_MODE === "true";

let cachedToken: string | null = null;
let tokenFetchedAt = 0;
const TOKEN_TTL_MS = 5 * 60 * 1000; // 5 minutes

async function getToken(): Promise<string | null> {
  // Local mode: no auth needed, backend has AUTH_DISABLED=true
  if (IS_LOCAL_MODE) return null;

  if (cachedToken && Date.now() - tokenFetchedAt < TOKEN_TTL_MS) {
    return cachedToken;
  }
  try {
    const res = await fetch("/api/auth/token");
    if (!res.ok) return null;
    const { token } = await res.json();
    cachedToken = token;
    tokenFetchedAt = Date.now();
    return token;
  } catch {
    return null;
  }
}

const MAX_RETRIES = 3;
const RETRY_DELAYS = [1000, 3000]; // delays before attempt 2 and 3

function isRetryable(e: unknown): boolean {
  if (e instanceof DOMException && e.name === "AbortError") return true;
  if (e instanceof TypeError) return true; // network failure / DNS / connection refused
  return false;
}

export async function authFetch(
  url: string,
  options: RequestInit = {},
  timeoutMs: number = 120_000, // 2 minute default timeout
): Promise<Response> {
  const token = await getToken();
  const headers = new Headers(options.headers);
  if (token) {
    headers.set("Authorization", `Bearer ${token}`);
  }

  let lastError: Error | null = null;

  for (let attempt = 0; attempt < MAX_RETRIES; attempt++) {
    if (attempt > 0) {
      await new Promise((r) => setTimeout(r, RETRY_DELAYS[attempt - 1]));
    }

    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);

    try {
      const res = await fetch(url, { ...options, headers, signal: controller.signal });
      clearTimeout(timer);

      if (res.status === 401 || res.status === 403) {
        cachedToken = null;
        tokenFetchedAt = 0;
        if (!IS_LOCAL_MODE && typeof window !== "undefined") {
          window.location.href = "/login";
        }
      }

      return res;
    } catch (e) {
      clearTimeout(timer);
      if (e instanceof DOMException && e.name === "AbortError") {
        lastError = new Error("Request timed out — server may be starting up.");
      } else {
        lastError = new Error("Cannot reach server — it may be starting up.");
      }
      if (!isRetryable(e)) break; // non-retryable error, stop immediately
    }
  }

  throw lastError ?? new Error("Cannot reach server. Please try again.");
}

export function clearAuthToken() {
  cachedToken = null;
  tokenFetchedAt = 0;
}

/**
 * Ping the backend to wake it up (Render free tier sleeps after 15min).
 * Polls up to `maxAttempts` times with `intervalMs` delay between attempts.
 * Calls `onStatus` so the UI can show connection state.
 */
export async function warmupBackend(
  apiUrl: string,
  onStatus?: (status: "connecting" | "ready" | "failed") => void,
  onProgress?: (elapsed: number) => void,
  maxAttempts = 30,
  intervalMs = 2000,
): Promise<boolean> {
  onStatus?.("connecting");
  const start = Date.now();
  for (let i = 0; i < maxAttempts; i++) {
    onProgress?.(Math.round((Date.now() - start) / 1000));
    try {
      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), 5000);
      const res = await fetch(`${apiUrl}/api/health`, {
        method: "GET",
        mode: "cors",
        signal: controller.signal,
      });
      clearTimeout(timer);
      if (res.ok) {
        onStatus?.("ready");
        return true;
      }
    } catch {
      // server not up yet — retry
    }
    if (i < maxAttempts - 1) {
      await new Promise((r) => setTimeout(r, intervalMs));
    }
  }
  onStatus?.("failed");
  return false;
}

/**
 * Ensure the backend is awake before making a request.
 * If the server was sleeping (Render free tier), wakes it up first.
 * Returns true if ready, false if warmup failed.
 */
export async function ensureBackendAwake(
  apiUrl: string,
  onStatus?: (status: "connecting" | "ready" | "failed") => void,
  onProgress?: (elapsed: number) => void,
): Promise<boolean> {
  try {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 5000);
    const res = await fetch(`${apiUrl}/api/health`, { signal: controller.signal });
    clearTimeout(timer);
    if (res.ok) return true;
  } catch {
    // Server not responding — run full warmup
  }
  return warmupBackend(apiUrl, onStatus, onProgress);
}
