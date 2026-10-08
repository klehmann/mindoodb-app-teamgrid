import react from '@vitejs/plugin-react'
import { havenBundle } from 'mindoodb-app-sdk/vite'
import { fileURLToPath } from 'node:url'
import { defineConfig, type Alias } from 'vite'
import wasm from 'vite-plugin-wasm'

const here = (path: string) => fileURLToPath(new URL(path, import.meta.url))
const genofficePackages = here('./vendor/genoffice/packages')
const nodeStubs = here('./src/node-stubs/index.ts')

/** `LOCAL_MINDOODB=1`: the first-party packages from their sibling checkouts' sources. */
function localMindooDBAliases(): Alias[] {
  if (process.env.LOCAL_MINDOODB !== '1') return []
  return [
    { find: /^mindoodb\/browser$/, replacement: here('../mindoodb/src/browser/index.ts') },
    { find: /^mindoodb\/core$/, replacement: here('../mindoodb/src/core/index.ts') },
    { find: /^mindoodb$/, replacement: here('../mindoodb/src/core/index.ts') },
    { find: /^mindoodb-app-sdk\/testing$/, replacement: here('../mindoodb-app-sdk/src/testing/index.ts') },
    { find: /^mindoodb-app-sdk\/vite$/, replacement: here('../mindoodb-app-sdk/src/vite/index.ts') },
    { find: /^mindoodb-app-sdk$/, replacement: here('../mindoodb-app-sdk/src/index.ts') },
    { find: /^mindoodb-view-language$/, replacement: here('../mindoodb-view-language/src/index.ts') },
  ]
}

/**
 * `/__haven-test/` frames the app with a mock Haven (see `src/testHost/main.ts`). `vite dev`
 * serves it anyway; a build only includes it with `HAVEN_TEST_HOST=1`, for preview
 * deployments, so the production URL never exposes a mock-data page to end users.
 */
function buildInputs(): Record<string, string> {
  return {
    main: here('./index.html'),
    ...(process.env.HAVEN_TEST_HOST === '1' ? { havenTest: here('./__haven-test/index.html') } : {}),
  }
}

// Store screenshots are only read from the app's own origin (landing page, Haven's setup
// wizard); inside the hosted bundle they would only grow every download.
function excludeFromHavenBundle(path: string) {
  return path.startsWith('listing/') || path.startsWith('__haven-test/')
}

export default defineConfig({
  // Relative asset URLs so the same build works from the app origin and from
  // Haven's `/__mindoodb_hosted_apps__/<bundleId>/` prefix.
  base: './',
  // wasm(): Automerge in the SDK's mock Haven (/__haven-test/) imports its .wasm as an ES module.
  plugins: [wasm(), react(), havenBundle({ exclude: excludeFromHavenBundle })],
  resolve: {
    alias: [
      ...localMindooDBAliases(),
      // GenOffice workspace packages: `exports` maps "." to src/index.ts and
      // "./x" to src/x(.ts), so the vendored sources resolve the same way.
      { find: /^@genoffice\/([a-z0-9-]+)$/, replacement: `${genofficePackages}/$1/src/index.ts` },
      { find: /^@genoffice\/([a-z0-9-]+)\/(.*)$/, replacement: `${genofficePackages}/$1/src/$2` },
      // Browser builds only; tests run in Node and get the real modules.
      ...(process.env.VITEST ? [] : [{ find: /^node:(crypto|fs|fs\/promises|os|path)$/, replacement: nodeStubs }]),
    ],
  },
  server: {
    host: '127.0.0.1',
    port: 4207,
  },
  build: {
    target: 'es2022',
    chunkSizeWarningLimit: 20_000,
    rollupOptions: { input: buildInputs() },
  },
})
