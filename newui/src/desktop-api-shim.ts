// Browser stand-in for GenOffice's Electron preload bridge (`window.desktopApi`).
// Workbook I/O runs against the WASM build of the xlsx sidecar and the
// vendored xlsx-gateway; everything not implemented yet resolves to
// `undefined` (or a no-op subscription) and is logged once, so the remaining
// surface stays visible in the console.
import type {
  DesktopApi,
  MenuAction,
  WorkbookFile,
} from '../vendor/genoffice/apps/sheets/src/shared/desktop-api'
import { loadXlsxEngine } from './xlsx/engine'
import { closeSession, openWorkbookBytes, readPivotDefinition, saveSession } from './xlsx/sessions'

const LANGUAGE = 'de'
const reported = new Set<string>()

function report(name: string): void {
  if (reported.has(name)) return
  reported.add(name)
  console.warn(`[newui] desktopApi.${name} is not implemented in the browser`)
}

// Development hook: a file queued here answers the next file dialog, so
// automated tests can open a workbook without a native picker, and the
// last saved bytes stay readable without going through a download.
const devHooks: {
  nextFile: File | null
  lastSaved: Uint8Array | null
  menu: (action: MenuAction) => void
} = {
  nextFile: null,
  lastSaved: null,
  menu: (action) => menuHandlers.forEach((handler) => handler(action)),
}

// The desktop app gets these from its native menu bar.
const menuHandlers = new Set<(action: MenuAction) => void>()
const MENU_SHORTCUTS: Record<string, MenuAction> = { o: 'open', s: 'save', S: 'save-as', p: 'print' }

function pickFile(accept: string): Promise<File | null> {
  if (devHooks.nextFile) {
    const file = devHooks.nextFile
    devHooks.nextFile = null
    return Promise.resolve(file)
  }
  return new Promise((resolve) => {
    const input = document.createElement('input')
    input.type = 'file'
    input.accept = accept
    input.addEventListener('change', () => resolve(input.files?.[0] ?? null), { once: true })
    input.addEventListener('cancel', () => resolve(null), { once: true })
    input.click()
  })
}

function download(bytes: Uint8Array, name: string): void {
  const blob = new Blob([bytes as BlobPart], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  })
  const link = document.createElement('a')
  link.href = URL.createObjectURL(blob)
  link.download = /\.xlsx$/i.test(name) ? name : `${name}.xlsx`
  link.click()
  setTimeout(() => URL.revokeObjectURL(link.href), 10_000)
}

const implemented: Partial<DesktopApi> = {
  getLanguage: async () => LANGUAGE,
  getTheme: async () => 'system',
  hasQueuedWorkbook: async () => false,
  consumeNewBlankWorkbook: async () => false,
  onMenuAction(handler) {
    menuHandlers.add(handler)
    return () => menuHandlers.delete(handler)
  },

  async selectWorkbook(): Promise<WorkbookFile | null> {
    const file = await pickFile('.xlsx,.xlsm')
    if (!file) return null
    const opened = await openWorkbookBytes(new Uint8Array(await file.arrayBuffer()), file.name, LANGUAGE)
    console.info(`[newui] opened ${file.name}: ${opened.sheets.length} sheet(s)`)
    return opened
  },
  readWorkbookRange: async (request) => (await loadXlsxEngine()).readRange(request) as never,
  readWorkbookFormulas: async (request) => (await loadXlsxEngine()).readFormulaCells(request) as never,
  readWorkbookMedia: async (request) => (await loadXlsxEngine()).readMedia(request) as never,
  readPivotDefinition: async (request) =>
    readPivotDefinition(request.sessionId, request.path, request.cachePath) as never,
  closeWorkbook: closeSession,

  async saveWorkbookEdits(request) {
    // Background saves (AutoSave) have nowhere to go until the MindooDB
    // backend exists; an explicit save downloads the patched xlsx.
    if (request.quiet) return { canceled: true }
    const { bytes, file, touchedEntries } = await saveSession(request, LANGUAGE)
    devHooks.lastSaved = bytes
    download(bytes, file.name)
    return { canceled: false, file, touchedEntries: [...touchedEntries] }
  },
}

export function installDesktopApiShim(): void {
  const api = new Proxy(implemented, {
    get(target, property) {
      if (typeof property !== 'string') return undefined
      const own = (target as Record<string, unknown>)[property]
      if (own) return own
      if (/^on[A-Z]/.test(property)) {
        return () => {
          report(property)
          return () => {}
        }
      }
      return async () => {
        report(property)
        return undefined
      }
    },
  })
  ;(window as unknown as { desktopApi: unknown; __newui: unknown }).desktopApi = api
  ;(window as unknown as { __newui: unknown }).__newui = devHooks
  window.addEventListener(
    'keydown',
    (event) => {
      if (!(event.metaKey || event.ctrlKey) || event.altKey) return
      const action = MENU_SHORTCUTS[event.shiftKey ? event.key.toUpperCase() : event.key.toLowerCase()]
      if (!action) return
      event.preventDefault()
      devHooks.menu(action)
    },
    true,
  )
}
