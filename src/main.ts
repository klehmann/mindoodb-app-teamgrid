import { installDesktopApiShim } from './desktop-api-shim'
import { installEditorInputFix } from './editor-input-fix'
import './overrides.css'
import './responsive.css'

installEditorInputFix()
await installDesktopApiShim()
await import('../vendor/genoffice/apps/sheets/src/renderer/main')
