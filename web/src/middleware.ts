import { auth } from "@/auth";

// Auth is active when either Google OAuth or password is configured
const isAuthConfigured =
  (process.env.GOOGLE_CLIENT_ID &&
    process.env.GOOGLE_CLIENT_ID !== "replace-with-google-client-id") ||
  !!process.env.ADMIN_PASSWORD;

export default auth((req) => {
  // If auth not configured, allow all requests (dev mode)
  if (!isAuthConfigured) return;

  const { pathname } = req.nextUrl;

  // Allow auth API routes, login page, and static assets
  if (
    pathname.startsWith("/api/auth") ||
    pathname === "/login" ||
    pathname.startsWith("/_next") ||
    pathname.startsWith("/favicon") ||
    pathname.match(/\.(svg|png|jpg|ico|css|js)$/)
  ) {
    return;
  }

  // Redirect unauthenticated users to login
  if (!req.auth) {
    const loginUrl = new URL("/login", req.url);
    loginUrl.searchParams.set("callbackUrl", pathname);
    return Response.redirect(loginUrl);
  }
});

export const config = {
  matcher: ["/((?!_next/static|_next/image).*)"],
};
