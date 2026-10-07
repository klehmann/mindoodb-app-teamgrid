import react from '@vitejs/plugin-react'
import { fileURLToPath } from 'node:url'
import { defineConfig } from 'vite'
import wasm from 'vite-plugin-wasm'

const genofficePackages = fileURLToPath(new URL('./vendor/genoffice/packages', import.meta.url))
const nodeStubs = fileURLToPath(new URL('./src/node-stubs/index.ts', import.meta.url))
// LOCAL_SDK=1: the mock Haven (/__haven-test/) from the sibling mindoodb-app-sdk
// checkout, e.g. to try unreleased test-host features.
const localSdkTesting =
  process.env.LOCAL_SDK === '1'
    ? [{ find: /^mindoodb-app-sdk\/testing$/, replacement: fileURLToPath(new URL('../../mindoodb-app-sdk/src/testing/index.ts', import.meta.url)) }]
    : []

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
      { find: /^node:(crypto|fs|fs\/promises|os|path)$/, replacement: nodeStubs },
      ...localSdkTesting,
    ],
  },
  build: {
    target: 'es2022',
    chunkSizeWarningLimit: 20_000,
  },
})
