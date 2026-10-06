import { existsSync, readFileSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import type { Plugin } from "vite";
import { defineConfig } from "vitest/config";

// El manifiesto de producción también se publica en GitHub Pages para poder
// descargarlo directamente (útil en la Mac).
function publicarManifiesto(): Plugin {
  return {
    name: "publicar-manifiesto",
    apply: "build",
    generateBundle() {
      this.emitFile({
        type: "asset",
        fileName: "manifest.xml",
        source: readFileSync("manifest.xml", "utf8"),
      });
    },
  };
}

// Certificados de desarrollo de Office creados con `npm run certs`. Solo se leen:
// instalarlos abre un diálogo del sistema, así que nunca se hace automáticamente.
function httpsDev() {
  const dir = join(homedir(), ".office-addin-dev-certs");
  const cert = join(dir, "localhost.crt");
  const key = join(dir, "localhost.key");
  if (!existsSync(cert) || !existsSync(key)) return undefined;
  return { cert: readFileSync(cert), key: readFileSync(key) };
}

export default defineConfig(({ command }) => ({
  base: command === "build" ? "/excel-agent/" : "/",
  plugins: [publicarManifiesto()],
  build: {
    target: ["es2020", "safari14"],
    rollupOptions: {
      input: {
        taskpane: "taskpane.html",
        commands: "commands.html",
      },
    },
  },
  server: {
    port: 3000,
    strictPort: true,
    https: command === "serve" && !process.env.VITEST ? httpsDev() : undefined,
  },
  test: {
    include: ["tests/**/*.test.ts"],
  },
}));
