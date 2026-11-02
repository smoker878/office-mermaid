import { readFile, readdir } from "node:fs/promises";
import { fileURLToPath, URL } from "node:url";
import { defineConfig, type Plugin, type UserConfig } from "vite";
import devCerts from "office-addin-dev-certs";

const root = fileURLToPath(new URL(".", import.meta.url));
const urlDev = "https://localhost:3000/";
const urlProd = "https://www.contoso.com/"; // Set to the production deployment location.

function officePages(browser: boolean, development: boolean): Plugin {
  return {
    name: "office-pages",
    enforce: "post",
    configureServer(server) {
      server.middlewares.use((req, res, next) => {
        const pathname = req.url?.split("?")[0];
        if (
          !browser &&
          (pathname === "/tests.html" ||
            pathname === "/office-tests.html" ||
            pathname?.startsWith("/tests/"))
        ) {
          res.statusCode = 404;
          res.end("Not found");
          return;
        }
        // Keep the browser URL unchanged: the editor uses it to skip Office.js.
        if (pathname === "/preview.html")
          req.url = req.url!.replace("/preview.html", "/taskpane.html");
        next();
      });
    },
    async generateBundle(_options, bundle) {
      const taskpane = bundle["taskpane.html"];
      if (taskpane?.type !== "asset") throw new Error("Missing taskpane.html build output");
      this.emitFile({ type: "asset", fileName: "preview.html", source: taskpane.source });
      for (const name of await readdir(new URL("./assets/", import.meta.url))) {
        this.emitFile({
          type: "asset",
          fileName: `assets/${name}`,
          source: await readFile(new URL(`./assets/${name}`, import.meta.url)),
        });
      }
      for (const name of (await readdir(root)).filter((name) => /^manifest.*\.xml$/.test(name))) {
        const manifest = await readFile(new URL(name, import.meta.url), "utf8");
        this.emitFile({
          type: "asset",
          fileName: name,
          source: development ? manifest : manifest.split(urlDev).join(urlProd),
        });
      }
    },
  };
}

export default defineConfig(async ({ command, mode, isPreview }): Promise<UserConfig> => {
  const browser = mode === "browser";
  const development = mode !== "production";
  return {
    root,
    // This app does not use .env files; do not load local secrets into the build.
    envDir: false,
    publicDir: false,
    appType: "mpa",
    plugins: [officePages(browser, development)],
    server: {
      port: 3000,
      strictPort: true,
      headers: { "Access-Control-Allow-Origin": "*" },
      https:
        command === "serve" && !isPreview && !browser
          ? await devCerts.getHttpsServerOptions()
          : undefined,
    },
    build: {
      sourcemap: true,
      minify: !development,
      rolldownOptions: {
        input: {
          taskpane: fileURLToPath(new URL("./taskpane.html", import.meta.url)),
          commands: fileURLToPath(new URL("./commands.html", import.meta.url)),
        },
      },
    },
  };
});
