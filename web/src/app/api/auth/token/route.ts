import { auth } from "@/auth";
import { SignJWT } from "jose";

export const dynamic = "force-dynamic";

const rawSecret = process.env.NEXTAUTH_SECRET;
if (!rawSecret) {
  throw new Error("NEXTAUTH_SECRET environment variable is required");
}
const secret = new TextEncoder().encode(rawSecret);

async function signToken(email: string, name: string): Promise<string> {
  return new SignJWT({ email, name })
    .setProtectedHeader({ alg: "HS256" })
    .setExpirationTime("24h")
    .sign(secret);
}

export async function GET() {
  const session = await auth();
  if (!session?.user?.email) {
    return Response.json({ error: "Not authenticated" }, { status: 401 });
  }

  const token = await signToken(session.user.email, session.user.name || "");
  return Response.json({ token });
}
