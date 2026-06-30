import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";

// Office requiere HTTPS incluso en desarrollo. `vite --https` genera un
// certificado autofirmado; en algunos equipos conviene usar `office-addin-dev-certs`.
export default defineConfig({
  plugins: [react()],
  server: {
    port: 3000,
    https: true,
    cors: true,
  },
});
