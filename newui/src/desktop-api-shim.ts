// Browser stand-in for GenOffice's Electron preload bridge (`window.desktopApi`).
// Workbook I/O runs against the WASM build of the xlsx sidecar and the
// vendored xlsx-gateway. Inside Haven, workbooks come from and go to MindooDB
// documents; standalone (plain `pnpm dev`) the editor opens and downloads
// xlsx files. Everything not implemented yet resolves to `undefined` (or a
// no-op subscription) and is logged once, so the remaining surface stays
// visible in the console.
import { LANGS, type Lang } from '@genoffice/i18n'

import type {
  DesktopApi,
  MenuAction,
  UiTheme,
  WorkbookFile,
} from '../vendor/genoffice/apps/sheets/src/shared/desktop-api'
import { connectHaven, listWorkbooks, type HavenConnection } from './haven/connection'
import { pickWorkbook } from './haven/picker'
import { changedSinceLoad } from './haven/store'
import { loadXlsxEngine } from './xlsx/engine'
import {
  closeSession,
  createEmptyDocument,
  importXlsxAsDocument,
  openStoredWorkbook,
  openWorkbookBytes,
  readPivotDefinition,
  saveSession,
  storedDocumentOf,
} from './xlsx/sessions'

/** How often an idle stored workbook checks for other people's changes. */
const REMOTE_POLL_MS = 5_000

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

let language: Lang = 'en'
/** Stored workbook to reopen on the next open request, without asking. */
let reopenDocumentId: string | null = null
let activeSessionId: string | null = null
let pendingEdits = 0

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

function toLang(locale: string | undefined): Lang {
  const code = locale?.slice(0, 2).toLowerCase()
  return (LANGS as readonly string[]).includes(code ?? '') ? (code as Lang) : 'en'
}

function toUiTheme(mode: string | undefined): UiTheme {
  return mode === 'light' || mode === 'dark' ? mode : 'system'
}

function opened(file: WorkbookFile): WorkbookFile {
  activeSessionId = file.sessionId
  return file
}

/** Haven: the workbook picker, looping until a workbook is open or the user cancels. */
async function selectStoredWorkbook(haven: HavenConnection): Promise<WorkbookFile | null> {
  if (reopenDocumentId) {
    const id = reopenDocumentId
    reopenDocumentId = null
    return opened(await openStoredWorkbook(haven, id, language))
  }
  const choice = await pickWorkbook(await listWorkbooks(haven), haven.canWrite)
  if (!choice) return null
  if (choice.kind === 'open') return opened(await openStoredWorkbook(haven, choice.id, language))
  if (choice.kind === 'new') {
    const id = await createEmptyDocument(haven, 'Neue Arbeitsmappe', 'Tabelle1')
    return opened(await openStoredWorkbook(haven, id, language))
  }
  const file = await pickFile('.xlsx,.xlsm')
  if (!file) return null
  const subject = file.name.replace(/\.xls[xm]$/i, '')
  const id = await importXlsxAsDocument(haven, new Uint8Array(await file.arrayBuffer()), subject, language)
  return opened(await openStoredWorkbook(haven, id, language))
}

/** Reopens the active stored workbook when someone else changed it and nothing is pending here. */
function watchRemoteChanges(haven: HavenConnection): void {
  let checking = false
  setInterval(async () => {
    if (checking || pendingEdits > 0 || document.visibilityState !== 'visible' || !activeSessionId) return
    const stored = storedDocumentOf(activeSessionId)
    if (!stored) return
    checking = true
    try {
      if ((await changedSinceLoad(haven, stored.loaded)) && pendingEdits === 0) {
        console.info('[newui] the workbook changed elsewhere; reloading it')
        reopenDocumentId = stored.loaded.id
        devHooks.menu('open')
      }
    } catch (error) {
      console.warn('[newui] checking for remote changes failed', error)
    } finally {
      checking = false
    }
  }, REMOTE_POLL_MS)
}

function createApi(haven: HavenConnection | null): Partial<DesktopApi> {
  return {
    getLanguage: async () => language as Awaited<ReturnType<DesktopApi['getLanguage']>>,
    getTheme: async () => toUiTheme(haven?.context.theme.mode),
    onThemeChanged: (handler) => haven?.session.onThemeChange((theme) => handler(toUiTheme(theme.mode))) ?? (() => {}),
    // Inside Haven the picker opens right away instead of an empty demo grid.
    hasQueuedWorkbook: async () => haven !== null,
    consumeNewBlankWorkbook: async () => false,
    getAutoSaveDefault: async () => ({ on: haven !== null, updatedAt: 0 }),
    onAutoSaveDefaultChanged: () => () => {},
    // AutoSave already writes every change to MindooDB; a separate crash
    // recovery copy has no place to go in the browser.
    writeWorkbookRecovery: async () => ({ ok: false }),
    notifyPendingEdits(count) {
      pendingEdits = count
    },
    onMenuAction(handler) {
      menuHandlers.add(handler)
      return () => menuHandlers.delete(handler)
    },

    async selectWorkbook(): Promise<WorkbookFile | null> {
      if (haven) return selectStoredWorkbook(haven)
      const file = await pickFile('.xlsx,.xlsm')
      if (!file) return null
      return opened(await openWorkbookBytes(new Uint8Array(await file.arrayBuffer()), file.name, language))
    },
    readWorkbookRange: async (request) => (await loadXlsxEngine()).readRange(request) as never,
    readWorkbookFormulas: async (request) => (await loadXlsxEngine()).readFormulaCells(request) as never,
    readWorkbookMedia: async (request) => (await loadXlsxEngine()).readMedia(request) as never,
    readPivotDefinition: async (request) =>
      readPivotDefinition(request.sessionId, request.path, request.cachePath) as never,
    closeWorkbook: closeSession,

    async saveWorkbookEdits(request) {
      const stored = storedDocumentOf(request.sessionId)
      // A plain xlsx has nowhere to go in the background; only an explicit
      // save downloads it.
      if (!stored && request.quiet) return { canceled: true }
      const outcome = await saveSession(request, language)
      if (outcome.stored) console.info(`[newui] saved to MindooDB (${outcome.stored.changed ? 'changed' : 'no changes'})`)
      opened(outcome.file)
      if (outcome.download) {
        devHooks.lastSaved = outcome.download
        download(outcome.download, outcome.file.name)
      }
      return { canceled: false, file: outcome.file, touchedEntries: [...outcome.touchedEntries] }
    },
  }
}

export async function installDesktopApiShim(): Promise<void> {
  let haven: HavenConnection | null = null
  try {
    haven = await connectHaven()
  } catch (error) {
    console.error('[newui] could not connect to Haven; continuing without it', error)
  }
  language = toLang(haven?.context.locale ?? navigator.language)
  const api = new Proxy(createApi(haven), {
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
  ;(window as unknown as { desktopApi: unknown }).desktopApi = api
  ;(window as unknown as { __newui: unknown }).__newui = devHooks
  if (haven) watchRemoteChanges(haven)
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
