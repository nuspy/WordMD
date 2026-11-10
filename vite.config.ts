import { defineConfig } from "vite";
import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { resolve } from "node:path";

// Word carica il pannello solo via HTTPS: in sviluppo si usano i certificati di office-addin-dev-certs.
async function httpsOptions() {
  if (process.argv.includes("build")) return undefined;
  const { getHttpsServerOptions } = await import("office-addin-dev-certs");
  return getHttpsServerOptions();
}

/** In sviluppo riceve l'esito dell'autotest eseguito dentro Word e lo salva in .selftest/result.json. */
const reportPlugin = () => ({
  name: "wordmd-selftest-report",
  configureServer(server: { middlewares: { use: (path: string, fn: (req: NodeJS.ReadableStream, res: { end: (s?: string) => void }) => void) => void } }) {
    // L'autotest parte solo se esiste .selftest/run (si consuma alla prima lettura).
    server.middlewares.use("/__selftest-flag", (_req, res) => {
      const flag = resolve(__dirname, ".selftest/run");
      const text = existsSync(flag) ? readFileSync(flag, "utf8").trim() : "";
      if (text) rmSync(flag);
      res.end(text);
    });
    server.middlewares.use("/__report", (req, res) => {
      let body = "";
      req.on("data", (c) => (body += c));
      req.on("end", () => {
        mkdirSync(resolve(__dirname, ".selftest"), { recursive: true });
        writeFileSync(resolve(__dirname, ".selftest/result.json"), body);
        res.end("ok");
      });
    });
  },
});

export default defineConfig(async () => ({
  plugins: [reportPlugin()],
  root: resolve(__dirname, "src/taskpane"),
  publicDir: resolve(__dirname, "assets"),
  base: "./",
  server: { port: 3000, strictPort: true, https: await httpsOptions() },
  build: { outDir: resolve(__dirname, "dist"), emptyOutDir: true, chunkSizeWarningLimit: 4000 },
}));
