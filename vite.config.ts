import { fileURLToPath, URL } from "node:url";
import { defineConfig } from "vitest/config";
import vue from "@vitejs/plugin-vue";
import { VitePWA } from "vite-plugin-pwa";
import wasm from "vite-plugin-wasm";

function createResolveAliases() {
  const aliases: Record<string, string> = {
    "@": fileURLToPath(new URL("./src", import.meta.url)),
  };

  if (process.env.LOCAL_MINDOODB === "1") {
    aliases["mindoodb/browser"] = fileURLToPath(new URL("../mindoodb/src/browser/index.ts", import.meta.url));
    aliases["mindoodb/core"] = fileURLToPath(new URL("../mindoodb/src/core/index.ts", import.meta.url));
    aliases.mindoodb = fileURLToPath(new URL("../mindoodb/src/core/index.ts", import.meta.url));
    aliases["mindoodb-app-sdk"] = fileURLToPath(new URL("../mindoodb-app-sdk/src/index.ts", import.meta.url));
    aliases["mindoodb-view-language"] = fileURLToPath(new URL("../mindoodb-view-language/src/index.ts", import.meta.url));
  }

  return aliases;
}

export default defineConfig({
  plugins: [
    wasm(),
    vue(),
    VitePWA({
      strategies: "injectManifest",
      srcDir: "src",
      filename: "sw.ts",
      injectRegister: false,
      manifest: false,
      injectManifest: {
        globPatterns: ["**/*.{css,html,ico,js,png,svg,ttf,wasm,webmanifest,woff,woff2}"],
        maximumFileSizeToCacheInBytes: 12 * 1024 * 1024,
      },
    }),
  ],
  resolve: {
    alias: createResolveAliases(),
  },
  server: {
    host: "127.0.0.1",
    port: 4207,
  },
  test: {
    environment: "jsdom",
    setupFiles: ["./src/test/setup.ts"],
  },
});
