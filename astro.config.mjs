// @ts-check
import node from "@astrojs/node";
import tailwindcss from "@tailwindcss/vite";
import { defineConfig } from "astro/config";

export default defineConfig({
  site: "https://github.com/evergladesfoundation/PPTX-Brand-Converter",
  output: "server",
  adapter: node({ mode: "standalone" }),
  server: {
    host: "127.0.0.1",
  },
  vite: {
    plugins: [tailwindcss()],
  },
});
