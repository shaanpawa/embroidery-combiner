import NextAuth from "next-auth";
import Google from "next-auth/providers/google";
import Credentials from "next-auth/providers/credentials";

const isGoogleConfigured =
  process.env.GOOGLE_CLIENT_ID &&
  process.env.GOOGLE_CLIENT_ID !== "replace-with-google-client-id";

const providers = [];

if (isGoogleConfigured) {
  providers.push(
    Google({
      clientId: process.env.GOOGLE_CLIENT_ID!,
      clientSecret: process.env.GOOGLE_CLIENT_SECRET!,
    })
  );
}

// Always register Credentials provider — password check happens at runtime
providers.push(
  Credentials({
    name: "Password",
    credentials: {
      password: { label: "Password", type: "password" },
    },
    async authorize(credentials) {
      const expected = process.env.ADMIN_PASSWORD;
      if (!expected) return null; // No password configured = reject all
      const pw = typeof credentials?.password === "string" ? credentials.password : "";
      if (pw === expected) {
        return { id: "operator", email: "operator@micro", name: "Operator" };
      }
      return null;
    },
  })
);

export const { handlers, auth, signIn, signOut } = NextAuth({
  providers,
  secret: process.env.NEXTAUTH_SECRET || "dev-secret-not-for-production",
  session: {
    strategy: "jwt",
  },
  pages: {
    signIn: "/login",
  },
  callbacks: {
    jwt({ token, user }) {
      if (user) {
        token.email = user.email;
        token.name = user.name;
      }
      return token;
    },
    session({ session, token }) {
      if (session.user) {
        session.user.email = token.email as string;
        session.user.name = token.name as string;
      }
      return session;
    },
  },
});
