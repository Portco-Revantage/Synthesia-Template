import { Clapperboard } from "lucide-react";
import { Toaster } from "@/components/ui/sonner";
import { HomePage } from "@/pages/home";

export function AppShell() {
  return (
    <div className="flex min-h-screen flex-col">
      <a
        href="#work"
        className="sr-only focus:not-sr-only focus:absolute focus:left-3 focus:top-3 focus:z-50 focus:rounded-md focus:bg-background focus:px-3 focus:py-2 focus:text-sm focus:outline-2 focus:outline-ring"
      >
        Skip to workspace
      </a>
      <header className="sticky top-0 z-30 flex h-14 items-center gap-3 border-b bg-primary px-4 text-primary-foreground">
        <Clapperboard className="size-5" aria-hidden="true" />
        <span className="text-base font-semibold tracking-tight">Synthesia Template Builder</span>
      </header>
      <main className="flex flex-1 flex-col">
        <HomePage />
      </main>
      <Toaster />
    </div>
  );
}
