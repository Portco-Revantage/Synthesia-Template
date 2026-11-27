import { useEffect } from "react";
import { Clapperboard } from "lucide-react";
import { Toaster } from "@/components/ui/sonner";
import { HomePage } from "@/pages/home";

export function AppShell() {
  // The glass theme is dark; this turns on the dark variants of status colours too.
  useEffect(() => {
    document.documentElement.classList.add("dark");
  }, []);
  return (
    <div className="flex min-h-screen flex-col">
      <a
        href="#work"
        className="sr-only focus:not-sr-only focus:absolute focus:left-3 focus:top-3 focus:z-50 focus:rounded-md focus:bg-background focus:px-3 focus:py-2 focus:text-sm focus:outline-2 focus:outline-ring"
      >
        Skip to workspace
      </a>
      <header className="glass sticky top-3 z-30 mx-4 mt-3 flex h-16 items-center gap-3 rounded-2xl px-5">
        <span className="grid size-9 place-items-center rounded-xl bg-gradient-to-br from-violet-500 via-indigo-500 to-cyan-500 shadow-lg shadow-violet-900/50">
          <Clapperboard className="size-5 text-white" aria-hidden="true" />
        </span>
        <span className="glow-text text-lg font-semibold tracking-tight">Synthesia Template Builder</span>
      </header>
      <main className="flex flex-1 flex-col">
        <HomePage />
      </main>
      <Toaster />
    </div>
  );
}
