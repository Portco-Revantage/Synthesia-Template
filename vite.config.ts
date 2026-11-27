import path from "path";
import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import tailwindcss from "@tailwindcss/vite";

// base "./" lets the built site run from any folder, including a GitHub Pages project URL.
export default defineConfig({
  base: "./",
  plugins: [react(), tailwindcss()],
  resolve: { alias: { "@": path.resolve(__dirname, "./src") } },
  // The PDF reader is large but only loads when a PDF is uploaded.
  build: { chunkSizeWarningLimit: 2000 },
});
