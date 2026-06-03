import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// Project is served from https://<user>.github.io/fleet-data-analyzer-script/
export default defineConfig({
  plugins: [react()],
  base: "/fleet-data-analyzer-script/",
});
