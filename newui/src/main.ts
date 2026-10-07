import { installDesktopApiShim } from './desktop-api-shim'

installDesktopApiShim()
await import('../vendor/genoffice/apps/sheets/src/renderer/main')
