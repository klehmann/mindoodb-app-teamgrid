import react from '@vitejs/plugin-react'
import { fileURLToPath } from 'node:url'
import { defineConfig } from 'vite'
import wasm from 'vite-plugin-wasm'

const genofficePackages = fileURLToPath(new URL('./vendor/genoffice/packages', import.meta.url))
const nodeStubs = fileURLToPath(new URL('./src/node-stubs/index.ts', import.meta.url))

export default defineConfig({
  base: './',
  // wasm(): Automerge in the SDK's mock Haven (/__haven-test/) imports its .wasm as an ES module.
  plugins: [wasm(), react()],
  resolve: {
    alias: [
      // GenOffice workspace packages: `exports` maps "." to src/index.ts and
      // "./x" to src/x(.ts), so the vendored sources resolve the same way.
      { find: /^@genoffice\/([a-z0-9-]+)$/, replacement: `${genofficePackages}/$1/src/index.ts` },
      { find: /^@genoffice\/([a-z0-9-]+)\/(.*)$/, replacement: `${genofficePackages}/$1/src/$2` },
      // Browser builds only; tests run in Node and get the real modules.
      ...(process.env.VITEST ? [] : [{ find: /^node:(crypto|fs|fs\/promises|os|path)$/, replacement: nodeStubs }]),
    ],
  },
  // Same host as the classic TeamGrid dev server (4207), next to it on 4208.
  server: {
    host: '127.0.0.1',
    port: 4208,
  },
  build: {
    target: 'es2022',
    chunkSizeWarningLimit: 20_000,
  },
})
