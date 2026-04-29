import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// Allow the localtunnel host and configure HMR for tunneled development
export default defineConfig({
  plugins: [react()],
  server: {
    host: true,
    port: 5173,
    strictPort: false,
    // Permit the tunneling hostname so Vite accepts requests from it
    allowedHosts: ["conciliacion2026.loca.lt", "localhost"],
    hmr: {
      host: "conciliacion2026.loca.lt",
    },
  },
});
