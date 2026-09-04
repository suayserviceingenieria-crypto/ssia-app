import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  // Necesario para GitHub Pages: la app se publica bajo
  // https://suayserviceingenieria-crypto.github.io/ssia-app/ (una subcarpeta,
  // no la raíz del dominio) — sin este "base", los archivos JS/CSS generados
  // se buscarían en la raíz del dominio y no cargarían.
  base: "/ssia-app/",
});
