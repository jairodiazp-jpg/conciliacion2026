import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

const backendTarget = process.env.VITE_PROXY_TARGET || "http://127.0.0.1:8000";

export default defineConfig({
  plugins: [react()],
  server: {
    host: true,
    port: 5173,
    strictPort: false,
    allowedHosts: true,
    proxy: {
      "/procesar": { target: backendTarget, changeOrigin: true },
      "/pse": { target: backendTarget, changeOrigin: true },
      "/health": { target: backendTarget, changeOrigin: true },
      "/version": { target: backendTarget, changeOrigin: true },
    },
  },
  preview: {
    host: true,
    port: 4173,
    proxy: {
      "/procesar": { target: backendTarget, changeOrigin: true },
      "/pse": { target: backendTarget, changeOrigin: true },
      "/health": { target: backendTarget, changeOrigin: true },
      "/version": { target: backendTarget, changeOrigin: true },
    },
  },
});
