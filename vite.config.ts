import { fileURLToPath, URL } from "node:url";
import { defineConfig } from "vitest/config";
import vue from "@vitejs/plugin-vue";
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
  }

  return aliases;
}

export default defineConfig({
  plugins: [wasm(), vue()],
  resolve: {
    alias: createResolveAliases(),
  },
  server: {
    host: "127.0.0.1",
    port: 4207,
  },
  test: {
    environment: "jsdom",
  },
});
