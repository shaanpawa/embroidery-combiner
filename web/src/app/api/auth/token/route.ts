import { auth } from "@/auth";
import { SignJWT } from "jose";

const isAuthConfigured =
  process.env.GOOGLE_CLIENT_ID &&
  process.env.GOOGLE_CLIENT_ID !== "replace-with-google-client-id";

const secret = new TextEncoder().encode(
  process.env.NEXTAUTH_SECRET || "dev-secret-not-for-production"
);

async function signToken(email: string, name: string): Promise<string> {
  return new SignJWT({ email, name })
    .setProtectedHeader({ alg: "HS256" })
    .setExpirationTime("24h")
    .sign(secret);
}

export async function GET() {
  // Dev mode: return a dev token without requiring session
  if (!isAuthConfigured) {
    const token = await signToken("local@dev", "Local Dev");
    return Response.json({ token });
  }

  const session = await auth();
  if (!session?.user?.email) {
    return Response.json({ error: "Not authenticated" }, { status: 401 });
  }

  const token = await signToken(session.user.email, session.user.name || "");
  return Response.json({ token });
}
