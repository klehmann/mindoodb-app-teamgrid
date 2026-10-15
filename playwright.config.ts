import { defineConfig, devices } from '@playwright/test'

/**
 * E2E tests against the mock Haven at `/__haven-test/`.
 *
 * - `pnpm test:e2e`: production build (`HAVEN_TEST_HOST=1`) served by `vite preview`,
 *   like the Cloudflare deploy. CI runs this.
 * - `pnpm test:e2e:local`: `vite` dev with `LOCAL_MINDOODB=1` (sibling checkouts).
 *
 * `pnpm test:fidelity` runs only the Excel comparison (e2e/fidelity).
 */
const local = process.env.E2E_LOCAL === '1'
const port = Number(process.env.E2E_PORT ?? (local ? 4384 : 4383))

export default defineConfig({
  testDir: './e2e',
  testMatch: '**/*.e2e.ts',
  fullyParallel: true,
  forbidOnly: !!process.env.CI,
  retries: 0,
  reporter: process.env.CI ? [['github'], ['html', { open: 'never' }]] : 'list',
  globalTeardown: './e2e/fidelity/teardown.ts',
  timeout: 120_000,
  expect: { timeout: 15_000 },
  use: {
    baseURL: `http://127.0.0.1:${port}`,
    viewport: { width: 1400, height: 900 },
    locale: 'en-US',
    trace: 'retain-on-failure',
    screenshot: 'only-on-failure',
  },
  projects: [{ name: 'chromium', use: { ...devices['Desktop Chrome'], viewport: { width: 1400, height: 900 } } }],
  webServer: {
    command: local
      ? `LOCAL_MINDOODB=1 vite --port ${port} --strictPort`
      : `HAVEN_TEST_HOST=1 vite build && vite preview --port ${port} --strictPort`,
    url: `http://127.0.0.1:${port}/__haven-test/`,
    reuseExistingServer: !process.env.CI,
    timeout: 300_000,
  },
})
