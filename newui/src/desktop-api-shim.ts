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
import {
  connectHaven,
  listWorkbooks,
  readProperties,
  writeProperties,
  type HavenConnection,
} from './haven/connection'
import { installFileMenu, type FileAction } from './haven/file-menu'
import {
  askTitle,
  chooseWorkbook,
  closeWelcome,
  editProperties,
  showWelcome,
  welcomeStrings,
  type WelcomeChoice,
} from './haven/welcome'
import { changedSinceLoad } from './haven/store'
import { loadXlsxEngine } from './xlsx/engine'
import {
  closeSession,
  createEmptyDocument,
  createFromTemplate,
  importXlsxAsDocument,
  openStoredWorkbook,
  openWorkbookBytes,
  readPivotDefinition,
  saveSession,
  sessionFor,
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
/** Haven's UI language for our own screens (TeamGrid's languages, not GenOffice's). */
let welcomeLanguage = 'en'
/** Stored workbook the next open request opens without asking (reload, file menu). */
let pendingDocumentId: string | null = null
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

/** Creates or picks the workbook a welcome-screen or file-menu choice stands for. */
async function workbookForChoice(haven: HavenConnection, choice: WelcomeChoice): Promise<string | null> {
  if (choice.kind === 'open') return choice.id
  if (choice.kind === 'new') return createEmptyDocument(haven, choice.subject, 'Tabelle1')
  if (choice.kind === 'template') return createFromTemplate(haven, choice.id, choice.subject, language)
  const file = await pickFile('.xlsx,.xlsm')
  if (!file) return null
  const subject = file.name.replace(/\.xls[xm]$/i, '')
  return importXlsxAsDocument(haven, new Uint8Array(await file.arrayBuffer()), subject, language)
}

/**
 * Haven: with no workbook open, the welcome screen (new, open, from template,
 * import) until something opened; with one open, the list of workbooks.
 */
async function selectStoredWorkbook(haven: HavenConnection): Promise<WorkbookFile | null> {
  if (pendingDocumentId) {
    const id = pendingDocumentId
    pendingDocumentId = null
    return opened(await openStoredWorkbook(haven, id, language))
  }
  if (activeSessionId) {
    const strings = welcomeStrings(welcomeLanguage)
    const all = await listWorkbooks(haven)
    const picked = await chooseWorkbook(
      all.filter((workbook) => !workbook.istemplate),
      strings.openTitle,
      strings.empty,
      strings,
    )
    return picked ? opened(await openStoredWorkbook(haven, picked.id, language)) : null
  }
  for (;;) {
    const choice = await showWelcome({
      language: welcomeLanguage,
      canCreate: haven.canWrite,
      listWorkbooks: () => listWorkbooks(haven),
    })
    try {
      const id = await workbookForChoice(haven, choice)
      if (!id) continue
      const file = await openStoredWorkbook(haven, id, language)
      closeWelcome()
      return opened(file)
    } catch (error) {
      console.error('[newui] opening the workbook failed', error)
    }
  }
}

/** Opens `id` through the editor's own open flow (which saves pending edits first). */
function openInEditor(id: string): void {
  pendingDocumentId = id
  devHooks.menu('open')
}

async function waitForSave(timeoutMs = 15_000): Promise<void> {
  if (pendingEdits === 0) return
  devHooks.menu('save')
  for (const started = Date.now(); pendingEdits > 0 && Date.now() - started < timeoutMs; ) {
    await new Promise((resolve) => setTimeout(resolve, 200))
  }
}

async function onFileAction(haven: HavenConnection, action: FileAction): Promise<void> {
  const strings = welcomeStrings(welcomeLanguage)
  const stored = activeSessionId ? storedDocumentOf(activeSessionId) : undefined
  if (action === 'open') {
    devHooks.menu('open')
  } else if (action === 'new') {
    const subject = await askTitle(strings.newTitle, '', strings)
    if (subject !== null) openInEditor(await createEmptyDocument(haven, subject, 'Tabelle1'))
  } else if (action === 'template') {
    const templates = (await listWorkbooks(haven)).filter((workbook) => workbook.istemplate)
    const template = await chooseWorkbook(templates, strings.templateTitle, strings.emptyTemplates, strings)
    if (!template) return
    const subject = await askTitle(strings.newTitle, strings.copyOf.replace('{title}', template.subject), strings)
    if (subject !== null) openInEditor(await createFromTemplate(haven, template.id, subject, language))
  } else if (action === 'import') {
    const id = await workbookForChoice(haven, { kind: 'import' })
    if (id) openInEditor(id)
  } else if (action === 'export') {
    if (!activeSessionId) return
    await waitForSave()
    const session = sessionFor(activeSessionId!)
    // The title may have changed in the properties since the session opened.
    download(session.bytes, storedDocumentOf(activeSessionId!)?.loaded.subject ?? session.name)
  } else if (action === 'properties' && stored) {
    const properties = await editProperties(await readProperties(haven, stored.loaded.id), strings)
    if (!properties) return
    await writeProperties(haven, stored.loaded.id, properties)
    stored.loaded.subject = properties.subject
    // Our own write must not look like someone else's change.
    await changedSinceLoad(haven, stored.loaded)
  }
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
        openInEditor(stored.loaded.id)
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
  welcomeLanguage = (haven?.context.locale ?? navigator.language).slice(0, 2).toLowerCase()
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
  if (haven) {
    watchRemoteChanges(haven)
    installFileMenu(welcomeStrings(welcomeLanguage), haven.canWrite, (action) => {
      onFileAction(haven!, action).catch((error) => console.error(`[newui] ${action} failed`, error))
    })
  }
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
