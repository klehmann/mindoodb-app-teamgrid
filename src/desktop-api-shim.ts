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
import { installFillSeries } from './fill-series'
import { installFileMenu, type FileAction } from './haven/file-menu'
import { hideBanner, showBanner } from './haven/banner'
import { canBrowseHistory, listRevisions } from './haven/history'
import { historyStrings } from './haven/history-strings'
import {
  askTitle,
  chooseRevision,
  chooseWorkbook,
  closeWelcome,
  editProperties,
  showWelcome,
  welcomeStrings,
  type WelcomeChoice,
} from './haven/welcome'
import { changedSinceLoad, loadWorkbook } from './haven/store'
import { loadXlsxEngine } from './xlsx/engine'
import { viewSheetCommands, type ViewSheetCommands } from './haven/view-sheets'
import {
  closeSession,
  createEmptyDocument,
  createFromTemplate,
  importXlsxAsDocument,
  openHistoricalWorkbook,
  openStoredWorkbook,
  openWorkbookBytes,
  readPivotDefinition,
  saveSession,
  sessionFor,
  storedDocumentOf,
} from './xlsx/sessions'

/** How often an idle stored workbook checks for other people's changes. */
const REMOTE_POLL_MS = 5_000

/**
 * Desktop-only parts of the API with no meaning in Haven (native menus and
 * windows, crash recovery, the AI panel, the sidecar process). The editor
 * asks for them on every start; answering with nothing is correct, so they
 * are not reported.
 */
const DESKTOP_ONLY = new Set([
  'getDocumentTheme',
  'onDocumentThemeChanged',
  'getAiPanelPrefs',
  'onAiPanelPrefsChanged',
  'getAiSettings',
  'aiGskStatus',
  'onRecoveryPrompt',
  'consumeHeadlessExport',
  'onWorkbookRenamed',
  'onSidecarCrashed',
  'onCloseSaveRequest',
  'onMcpCommand',
  'signalMcpReady',
])

const reported = new Set<string>()

/** Development builds name a method the shim does not answer yet, once. */
function report(name: string): void {
  if (!import.meta.env.DEV || DESKTOP_ONLY.has(name) || reported.has(name)) return
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

/** A revision to open next (File → Browse revisions), and the workbook the open session shows an older state of. */
let pendingRevision: { id: string; timestamp: number } | null = null
let historicalOf: string | null = null
/** Haven's time travel: everything is read-only as of this time (ms). */
let timeTravelDate: number | null = null

function opened(file: WorkbookFile, revisionOf: string | null = null): WorkbookFile {
  activeSessionId = file.sessionId
  historicalOf = revisionOf
  if (revisionOf) {
    const strings = historyStrings(welcomeLanguage)
    showBanner(strings.historicalBanner, { label: strings.returnToCurrent, run: () => openInEditor(revisionOf) })
  } else if (!timeTravelDate) {
    hideBanner()
  }
  return file
}

function showTimeTravelBanner(): void {
  if (!timeTravelDate) return
  const date = new Intl.DateTimeFormat(welcomeLanguage, { dateStyle: 'medium', timeStyle: 'short' }).format(timeTravelDate)
  showBanner(historyStrings(welcomeLanguage).timeTravelBanner.replace('{date}', date))
}

/** Lists the open workbook's revisions and opens the picked one read-only (or the current state). */
async function browseRevisions(haven: HavenConnection): Promise<void> {
  const id = historicalOf ?? (activeSessionId ? storedDocumentOf(activeSessionId)?.loaded.id : undefined)
  if (!id) return
  const strings = historyStrings(welcomeLanguage)
  await waitForSave()
  let revisions
  try {
    revisions = await listRevisions(haven, await loadWorkbook(haven, id))
  } catch (error) {
    console.error('[newui] listing revisions failed', error)
    showBanner(strings.listFailed)
    return
  }
  const picked = await chooseRevision(revisions, strings, welcomeStrings(welcomeLanguage), welcomeLanguage)
  if (!picked) return
  if (picked.current) {
    openInEditor(id)
    return
  }
  pendingRevision = { id, timestamp: picked.timestamp }
  devHooks.menu('open')
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
  if (pendingRevision) {
    const { id, timestamp } = pendingRevision
    pendingRevision = null
    try {
      return opened(await openHistoricalWorkbook(haven, id, timestamp, language), id)
    } catch (error) {
      console.error('[newui] opening the revision failed', error)
      showBanner(historyStrings(welcomeLanguage).loadFailed)
      return null
    }
  }
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
  } else if (action === 'revisions') {
    await browseRevisions(haven)
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

/** Opens `id` again and, once the editor shows it, the sheet named `sheetName`. */
function reopenAt(id: string, sheetName: string): void {
  const before = activeSessionId
  openInEditor(id)
  const started = Date.now()
  const timer = setInterval(() => {
    const workbook = (window as unknown as {
      __univerAPI?: { getActiveWorkbook(): { getSheetByName(name: string): { activate(): unknown } | null } | null }
    }).__univerAPI?.getActiveWorkbook()
    const sheet = activeSessionId !== before ? workbook?.getSheetByName(sheetName) : null
    if (sheet) sheet.activate()
    if (sheet || Date.now() - started > 20_000) clearInterval(timer)
  }, 250)
}

/** Relabels the File menu after a language switch (set once it is installed). */
let setFileMenuStrings: ((strings: ReturnType<typeof welcomeStrings>) => void) | undefined

function createApi(haven: HavenConnection | null): Partial<DesktopApi> {
  return {
    getLanguage: async () => language as Awaited<ReturnType<DesktopApi['getLanguage']>>,
    getTheme: async () => toUiTheme(haven?.context.theme.mode),
    // Haven's language switches without a reload; so does the editor.
    onLanguageChanged: (handler) =>
      haven?.session.onLocaleChange((locale) => {
        language = toLang(locale)
        welcomeLanguage = locale.slice(0, 2).toLowerCase()
        setFileMenuStrings?.(welcomeStrings(welcomeLanguage))
        handler(language as Parameters<typeof handler>[0])
      }) ?? (() => {}),
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
  installFillSeries(() => welcomeLanguage)
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
    timeTravelDate = haven.context.timeTravelDate ?? null
    // Time travel is a frozen state: nothing changes underneath.
    if (!timeTravelDate) watchRemoteChanges(haven)
    const connected = haven
    ;(window as unknown as { teamGridViewSheets: ViewSheetCommands }).teamGridViewSheets = viewSheetCommands({
      haven: connected,
      language: () => welcomeLanguage,
      welcomeStrings: () => welcomeStrings(welcomeLanguage),
      activeWorkbookId: () => (activeSessionId ? storedDocumentOf(activeSessionId)?.loaded.id : undefined),
      waitForSave,
      reopen: reopenAt,
    })
    const historyAvailable = canBrowseHistory(haven) && !timeTravelDate
    setFileMenuStrings = installFileMenu(
      welcomeStrings(welcomeLanguage),
      haven.canWrite,
      (action) => {
        onFileAction(haven!, action).catch((error) => console.error(`[newui] ${action} failed`, error))
      },
      () => (historyAvailable ? historyStrings(welcomeLanguage).browseRevisions : undefined),
    )
    showTimeTravelBanner()
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
