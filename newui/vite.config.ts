import react from '@vitejs/plugin-react'
import { fileURLToPath } from 'node:url'
import { defineConfig } from 'vite'

const genofficePackages = fileURLToPath(new URL('./vendor/genoffice/packages', import.meta.url))
const nodeStubs = fileURLToPath(new URL('./src/node-stubs/index.ts', import.meta.url))

export default defineConfig({
  base: './',
  plugins: [react()],
  resolve: {
    alias: [
      // GenOffice workspace packages: `exports` maps "." to src/index.ts and
      // "./x" to src/x(.ts), so the vendored sources resolve the same way.
      { find: /^@genoffice\/([a-z0-9-]+)$/, replacement: `${genofficePackages}/$1/src/index.ts` },
      { find: /^@genoffice\/([a-z0-9-]+)\/(.*)$/, replacement: `${genofficePackages}/$1/src/$2` },
      { find: /^node:(crypto|fs|fs\/promises|os|path)$/, replacement: nodeStubs },
    ],
  },
  build: {
    target: 'es2022',
    chunkSizeWarningLimit: 20_000,
  },
})
