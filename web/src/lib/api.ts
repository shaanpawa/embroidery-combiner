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
  options: RequestInit = {}
): Promise<Response> {
  const token = await getToken();
  const headers = new Headers(options.headers);
  if (token) {
    headers.set("Authorization", `Bearer ${token}`);
  }

  const res = await fetch(url, { ...options, headers });

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
