import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { ConfirmDialogProvider } from "@/components/confirm-dialog";
import { AppShell } from "@/components/app-shell";
import "./index.css";

createRoot(document.getElementById("root")!).render(
  <StrictMode>
    <ConfirmDialogProvider>
      <AppShell />
    </ConfirmDialogProvider>
  </StrictMode>,
);
