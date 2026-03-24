"use client";

import { createContext, useContext, useEffect, useState } from "react";

type Theme = "light" | "dark";

const ThemeContext = createContext<{
  theme: Theme;
  toggle: () => void;
}>({ theme: "light", toggle: () => {} });

export function useTheme() {
  return useContext(ThemeContext);
}

export function ThemeProvider({ children }: { children: React.ReactNode }) {
  const [theme, setTheme] = useState<Theme>("light");

  useEffect(() => {
    const saved = localStorage.getItem("theme") as Theme | null;
    if (saved) {
      setTheme(saved);
      document.documentElement.setAttribute("data-theme", saved);
      // Force Safari to repaint backdrop-filter elements after CSS variable change
      requestAnimationFrame(() => {
        document.documentElement.style.transform = "translateZ(0)";
        requestAnimationFrame(() => {
          document.documentElement.style.transform = "";
        });
      });
    }
  }, []);

  const toggle = () => {
    const next = theme === "light" ? "dark" : "light";
    setTheme(next);
    localStorage.setItem("theme", next);
    document.documentElement.setAttribute("data-theme", next);
    // Force Safari to repaint backdrop-filter elements after CSS variable change
    requestAnimationFrame(() => {
      document.documentElement.style.transform = "translateZ(0)";
      requestAnimationFrame(() => {
        document.documentElement.style.transform = "";
      });
    });
  };

  return (
    <ThemeContext.Provider value={{ theme, toggle }}>
      {children}
    </ThemeContext.Provider>
  );
}
