import { installDesktopApiShim } from './desktop-api-shim'

await installDesktopApiShim()
await import('../vendor/genoffice/apps/sheets/src/renderer/main')
