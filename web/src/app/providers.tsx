"use client";

import { SessionProvider } from "next-auth/react";
import { LanguageProvider } from "./i18n";

const IS_LOCAL_MODE = process.env.NEXT_PUBLIC_LOCAL_MODE === "true";

export function Providers({ children }: { children: React.ReactNode }) {
  // In local mode, skip SessionProvider (no auth needed)
  if (IS_LOCAL_MODE) {
    return <LanguageProvider>{children}</LanguageProvider>;
  }

  return (
    <SessionProvider>
      <LanguageProvider>{children}</LanguageProvider>
    </SessionProvider>
  );
}
