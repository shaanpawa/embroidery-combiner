let cachedToken: string | null = null;
let tokenFetchedAt = 0;
const TOKEN_TTL_MS = 5 * 60 * 1000; // 5 minutes

async function getToken(): Promise<string | null> {
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

  // Timeout via AbortController — prevents infinite hangs if backend is cold/down
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);

  let res: Response;
  try {
    res = await fetch(url, { ...options, headers, signal: controller.signal });
  } catch (e) {
    clearTimeout(timer);
    if (e instanceof DOMException && e.name === "AbortError") {
      throw new Error("Request timed out — server may be starting up. Try again in 30 seconds.");
    }
    throw new Error("Cannot reach server. Check your internet connection.");
  }
  clearTimeout(timer);

  if (res.status === 401 || res.status === 403) {
    cachedToken = null;
    tokenFetchedAt = 0;
    if (typeof window !== "undefined") {
      window.location.href = "/login";
    }
  }

  return res;
}

export function clearAuthToken() {
  cachedToken = null;
  tokenFetchedAt = 0;
}

/**
 * Ping the backend to wake it up (Render free tier sleeps after 15min).
 * Call this early (e.g. on page load) so the backend is warm by the time
 * the user uploads files.
 */
export async function warmupBackend(apiUrl: string): Promise<void> {
  try {
    await fetch(`${apiUrl}/api/health`, { method: "GET", mode: "cors" });
  } catch {
    // Ignore — best effort
  }
}
