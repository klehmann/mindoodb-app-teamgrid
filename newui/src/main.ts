import { installDesktopApiShim } from './desktop-api-shim'
import { installEditorInputFix } from './editor-input-fix'

installEditorInputFix()
await installDesktopApiShim()
await import('../vendor/genoffice/apps/sheets/src/renderer/main')
