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
  WorkbookCreateDocumentRequest,
  WorkbookExportPdfRequest,
  WorkbookFile,
} from '../vendor/genoffice/apps/sheets/src/shared/desktop-api'
import { editorSkill, registerTeamgridAgentTools } from './agent/registerAgentTools'
import {
  connectHaven,
  listWorkbooks,
  readProperties,
  writeProperties,
  type HavenConnection,
} from './haven/connection'
import { installFillSeries } from './fill-series'
import { installTextValues } from './formulas/text-values'
import { installExcelFunctions } from './formulas/excel-functions'
import { installImplicitIntersection } from './formulas/implicit-intersection'
import { activeWorkbook, installUniverAccess } from './univer-access'
import { hideBanner, showBanner } from './haven/banner'
import type { ToolbarCommand } from './haven/toolbar/AppToolbar'
import { mountToolbar, type ToolbarHandle } from './haven/toolbar/mount'
import { toolbarStrings } from './haven/toolbar-strings'
import { WorkbookWindows, type WindowView } from './haven/windows'
import { canBrowseHistory, listRevisions } from './haven/history'
import { historyStrings } from './haven/history-strings'
import {
  askChoice,
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
import { hasMacros } from './xlsx/macros'
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
  'onAiSettingsChanged',
  'openAiModelSettings',
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
let openRetry: ReturnType<typeof setInterval> | null = null
/** The workbook the last open request asked for, and the document it opened (a 1.x workbook's copy). */
let lastOpened: { requested: string; storedId: string } | null = null
let activeSessionId: string | null = null
let pendingEdits = 0
/** Answers the welcome screen in place of the user (an agent opening a workbook). */
let answerWelcome: ((choice: WelcomeChoice) => void) | null = null
/** Takes the next file GenOffice's create_document writes (an agent's CSV export). */
let documentSink: ((request: WorkbookCreateDocumentRequest) => void) | null = null

/** The workbooks opened in this launch (Window menu) and the app toolbar showing them. */
const windows = new WorkbookWindows()
let toolbar: ToolbarHandle | null = null
/** GenOffice's AutoSave switch (patches/sheets-window-toolbar.patch reports it). */
let autoSaveOn = true
/** An explicit save (toolbar, Cmd/Ctrl+S) is running; AutoSave's quiet saves do not count. */
let explicitSaving = false
/** The welcome screen is up (at start, after closing the last workbook, "All spreadsheets"). */
let welcomeShown = false
/** The next open request shows the welcome screen although a workbook is open. */
let welcomeNext = false
/** The view to restore once the editor shows the workbook it belongs to. */
let pendingView: { id: string; view: WindowView } | null = null

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
  // A workbook with a VBA project is only valid as .xlsm.
  const extension = hasMacros(bytes) ? '.xlsm' : '.xlsx'
  link.download = /\.xls[xm]$/i.test(name) ? name.replace(/\.xls[xm]$/i, extension) : `${name}${extension}`
  link.click()
  setTimeout(() => URL.revokeObjectURL(link.href), 10_000)
}

/** Prints a laid-out sheet (GenOffice's print job HTML) through the browser's print dialog. */
async function printHtml(request: WorkbookExportPdfRequest): Promise<void> {
  const size =
    typeof request.pageSize === 'string' ? request.pageSize : `${request.pageSize.width}in ${request.pageSize.height}in`
  const { top, right, bottom, left } = request.margins
  const page =
    `<style>@page{size:${size} ${request.landscape ? 'landscape' : 'portrait'};margin:${top}in ${right}in ${bottom}in ${left}in}` +
    `${request.scale !== 1 ? `body{zoom:${request.scale}}` : ''}</style>`
  const frame = document.createElement('iframe')
  frame.setAttribute('aria-hidden', 'true')
  frame.style.cssText = 'position:fixed;right:0;bottom:0;width:0;height:0;border:0;visibility:hidden'
  frame.srcdoc = /<head[^>]*>/i.test(request.html)
    ? request.html.replace(/<head[^>]*>/i, (head) => `${head}${page}`)
    : `${page}${request.html}`
  await new Promise<void>((resolve) => {
    frame.addEventListener('load', () => resolve(), { once: true })
    document.body.append(frame)
  })
  try {
    frame.contentWindow?.focus()
    // Blocks until the dialog closes; "Save as PDF" there is the PDF export.
    frame.contentWindow?.print()
  } finally {
    setTimeout(() => frame.remove(), 1_000)
  }
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

/** The file the editor was last given (open, reopen after a save). */
let activeFile: WorkbookFile | null = null
/** Masters of legacy array (CSE) formulas in the open file, `sheetId:row:column`; they keep their arrays. */
const arrayFormulaCells = new Set<string>()

function rememberArrayFormulas(sheetId: string, cells: ReadonlyArray<{ row: number; column: number; arrayRef?: string }> | undefined) {
  for (const cell of cells ?? []) if (cell.arrayRef) arrayFormulaCells.add(`${sheetId}:${cell.row}:${cell.column}`)
}

/** The session whose cells the editor read last: it shows that file once its first range loaded. */
let editorSessionId: string | null = null

/**
 * Whether the editor shows `activeFile` yet: it keeps showing the previous
 * workbook (or its blank start-up one) until the new file's first range loaded.
 * Two workbooks can have the same sheets (two new ones: "Tabelle1"), so the
 * session the editor reads from decides, not the sheet list alone.
 */
function editorShowsActiveFile(): boolean {
  if (!activeFile || editorSessionId !== activeFile.sessionId) return false
  const info = editorSkill()?.getActiveSheetInfo()
  if (!info || info.mode !== 'lazy' || info.sheets.length !== activeFile!.sheets.length) return false
  return activeFile!.sheets.every((sheet, index) => info.sheets[index]?.id === sheet.id && info.sheets[index]?.name === sheet.name)
}

function opened(file: WorkbookFile, revisionOf: string | null = null): WorkbookFile {
  if (activeSessionId !== file.sessionId) arrayFormulaCells.clear()
  activeSessionId = file.sessionId
  activeFile = file
  historicalOf = revisionOf
  if (revisionOf) {
    const strings = historyStrings(welcomeLanguage)
    showBanner(strings.historicalBanner, { label: strings.returnToCurrent, run: () => openInEditor(revisionOf) })
  } else if (!timeTravelDate) {
    hideBanner()
  }
  const stored = storedDocumentOf(file.sessionId)
  const id = revisionOf ?? stored?.loaded.id
  if (id) {
    const entry = windows.show(id, stored?.loaded.subject ?? file.name)
    if (entry.view && !pendingView) pendingView = { id, view: entry.view }
    if (pendingView?.id === id) restoreView(file.sessionId, pendingView.view)
    pendingView = null
  }
  syncToolbar()
  return file
}

const openedRequest = () => lastOpened

/** `opened` for a stored workbook an open request asked for by id. */
function openedFor(requested: string, file: WorkbookFile): WorkbookFile {
  const storedId = storedDocumentOf(file.sessionId)?.loaded.id
  if (storedId) lastOpened = { requested, storedId }
  return opened(file)
}

/** The sheet and selection the editor shows now, once it shows the active file. */
function currentView(): WindowView | undefined {
  if (!editorShowsActiveFile()) return undefined
  try {
    const sheet = activeWorkbook()?.getActiveSheet()
    if (!sheet) return undefined
    const range = sheet.getSelection()?.getActiveRange()?.getA1Notation()
    return { sheetName: sheet.getSheetName(), ...(range ? { range } : {}) }
  } catch {
    return undefined
  }
}

/** Shows `view`'s sheet and selection once the editor shows the session `sessionId`. */
function restoreView(sessionId: string, view: WindowView): void {
  const started = Date.now()
  const timer = setInterval(() => {
    if (activeSessionId !== sessionId || Date.now() - started > 20_000) return clearInterval(timer)
    if (!editorShowsActiveFile()) return
    clearInterval(timer)
    try {
      const sheet = activeWorkbook()?.getSheetByName(view.sheetName)
      if (!sheet) return
      sheet.activate()
      if (view.range) sheet.getRange(view.range).activate()
    } catch (error) {
      console.warn('[newui] restoring the view failed', error)
    }
  }, 150)
}

/** Remembers where the shown workbook was left, for when it opens again. */
function rememberView(): void {
  const id = windows.active?.id
  if (id) windows.remember(id, currentView())
}

function syncToolbar(): void {
  if (!toolbar) return
  const showing = activeSessionId !== null && !welcomeShown
  const stored = activeSessionId ? storedDocumentOf(activeSessionId) : undefined
  const active = welcomeShown ? undefined : windows.active
  toolbar.update({
    hasWorkbook: showing,
    title: stored?.loaded.subject ?? active?.title ?? '',
    windows: windows.list().map((window) => ({ id: window.id, title: window.title, active: window.id === active?.id })),
    pending: !autoSaveOn && pendingEdits > 0,
    saving: explicitSaving,
    readOnly: activeFile?.readOnly ?? false,
  })
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
    return openedFor(id, await openStoredWorkbook(haven, id, language))
  }
  const toWelcome = welcomeNext
  welcomeNext = false
  if (activeSessionId && !toWelcome) {
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
  windows.hide()
  welcomeShown = true
  syncToolbar()
  for (;;) {
    const choice = await new Promise<WelcomeChoice>((resolve) => {
      answerWelcome = resolve
      void showWelcome({
        language: welcomeLanguage,
        canCreate: haven.canWrite && !timeTravelDate,
        listWorkbooks: () => listWorkbooks(haven),
      }).then(resolve)
    })
    answerWelcome = null
    try {
      const id = await workbookForChoice(haven, choice)
      if (!id) continue
      const file = await openStoredWorkbook(haven, id, language)
      closeWelcome()
      welcomeShown = false
      return openedFor(id, file)
    } catch (error) {
      console.error('[newui] opening the workbook failed', error)
    }
  }
}

/** Shows a stored workbook: answers the welcome screen when it is up, else opens it in the editor. */
function showWorkbook(id: string): void {
  if (answerWelcome) answerWelcome({ kind: 'open', id })
  else openInEditor(id)
}

/**
 * Before the shown workbook makes room for another (or for itself, reloaded):
 * pending edits are saved, with AutoSave off after asking. False: the user
 * cancelled.
 */
async function leaveWorkbook(): Promise<boolean> {
  rememberView()
  if (pendingEdits === 0 || activeFile?.readOnly || welcomeShown) return true
  if (!autoSaveOn) {
    const strings = toolbarStrings(welcomeLanguage)
    const welcome = welcomeStrings(welcomeLanguage)
    const title = windows.active?.title || welcome.untitled
    const choice = await askChoice(
      strings.unsavedTitle,
      strings.unsavedBody.replace('{title}', title),
      [
        { key: 'discard', label: strings.dontSave },
        { key: 'save', label: strings.save },
      ],
      welcome,
    )
    if (choice === null) return false
    if (choice === 'discard') return true
  }
  await explicitSave()
  return pendingEdits === 0
}

/** Shows the welcome screen; the open workbooks stay in the Window menu. */
function showWelcomeScreen(): void {
  if (welcomeShown) return
  welcomeNext = true
  devHooks.menu('open')
}

/** Saves now (toolbar, Cmd/Ctrl+S) and waits for the save to finish. */
async function explicitSave(): Promise<void> {
  if (pendingEdits === 0 && !explicitSaving) return
  explicitSaving = true
  syncToolbar()
  try {
    await waitForSave()
  } finally {
    explicitSaving = false
    syncToolbar()
  }
}

/** Opens `id` through the editor's own open flow (callers save pending edits first). */
function openInEditor(id: string): void {
  pendingDocumentId = id
  devHooks.menu('open')
  // The editor ignores Open while it loads a workbook, or reloads one after a
  // save; ask again until the request was taken, and drop it if it never is.
  if (openRetry) clearInterval(openRetry)
  const started = Date.now()
  openRetry = setInterval(() => {
    if (pendingDocumentId === id && Date.now() - started < 30_000) {
      devHooks.menu('open')
      return
    }
    if (pendingDocumentId === id) pendingDocumentId = null
    clearInterval(openRetry!)
    openRetry = null
  }, 1_000)
}

/**
 * Opens a stored workbook for an agent and resolves once the editor shows it:
 * answers the welcome screen when it is up, else goes through the editor's
 * open flow like the File menu. A TeamGrid 1.x workbook opens as its copy, so
 * any newly opened stored workbook counts.
 */
async function openForAgent(id: string, timeoutMs = 30_000): Promise<void> {
  // Edits still waiting for AutoSave would be dropped with the workbook they belong to.
  await waitForSave()
  lastOpened = null
  if (answerWelcome) answerWelcome({ kind: 'open', id })
  else openInEditor(id)
  for (const started = Date.now(); Date.now() - started < timeoutMs; ) {
    await new Promise((resolve) => setTimeout(resolve, 100))
    // The document this request opened (a TeamGrid 1.x workbook opens as its copy), shown in
    // the editor; a save reopening the previous workbook meanwhile does not count.
    const shown = activeSessionId ? storedDocumentOf(activeSessionId)?.loaded.id : undefined
    const last = openedRequest() // set by the open flow meanwhile
    if (last?.requested === id && shown === last.storedId && editorShowsActiveFile()) return
  }
  throw new Error('The workbook did not open in time.')
}

/** One worksheet as CSV through GenOffice's create_document (displayed values, as its CSV export). */
async function csvForAgent(sheetId: string | undefined): Promise<{ name: string; text: string }> {
  const skill = editorSkill()
  if (!skill?.createDocument) throw new Error('The editor is not ready.')
  let captured: WorkbookCreateDocumentRequest | null = null
  documentSink = (request) => {
    captured = request
  }
  try {
    const outcome = await skill.createDocument({ type: 'csv', ...(sheetId ? { sheetId } : {}) })
    if (!outcome.ok) throw new Error(outcome.error)
  } finally {
    documentSink = null
  }
  const request = captured as WorkbookCreateDocumentRequest | null
  if (!request) throw new Error('The CSV export produced no file.')
  return { name: request.title, text: request.content }
}

/** Saves an agent's edit: the editor reports it as pending a moment after the edit applied. */
async function saveForAgent(): Promise<void> {
  for (const started = Date.now(); pendingEdits === 0 && Date.now() - started < 1_500; ) {
    await new Promise((resolve) => setTimeout(resolve, 25))
  }
  await waitForSave()
}

async function waitForSave(timeoutMs = 15_000): Promise<void> {
  if (pendingEdits === 0) return
  devHooks.menu('save')
  for (const started = Date.now(); pendingEdits > 0 && Date.now() - started < timeoutMs; ) {
    await new Promise((resolve) => setTimeout(resolve, 200))
  }
}

/**
 * File menu New / From template / Import: `ask` runs the dialogs and returns
 * what creates the workbook; it runs once the shown workbook is left.
 */
async function createAndShow(ask: () => Promise<(() => Promise<string>) | null>): Promise<void> {
  const create = await ask()
  if (create && (await leaveWorkbook())) showWorkbook(await create())
}

async function onToolbarCommand(haven: HavenConnection, command: ToolbarCommand): Promise<void> {
  const strings = welcomeStrings(welcomeLanguage)
  const stored = activeSessionId && !welcomeShown ? storedDocumentOf(activeSessionId) : undefined
  switch (command) {
    case 'open': {
      const all = await listWorkbooks(haven)
      const picked = await chooseWorkbook(
        all.filter((workbook) => !workbook.istemplate),
        strings.openTitle,
        strings.empty,
        strings,
      )
      if (picked && (await leaveWorkbook())) showWorkbook(picked.id)
      return
    }
    case 'new':
      return createAndShow(async () => {
        const subject = await askTitle(strings.newTitle, '', strings)
        return subject === null ? null : () => createEmptyDocument(haven, subject, 'Tabelle1')
      })
    case 'template':
      return createAndShow(async () => {
        const templates = (await listWorkbooks(haven)).filter((workbook) => workbook.istemplate)
        const template = await chooseWorkbook(templates, strings.templateTitle, strings.emptyTemplates, strings)
        if (!template) return null
        const subject = await askTitle(strings.newTitle, strings.copyOf.replace('{title}', template.subject), strings)
        return subject === null ? null : () => createFromTemplate(haven, template.id, subject, language)
      })
    case 'import':
      return createAndShow(async () => {
        const file = await pickFile('.xlsx,.xlsm')
        if (!file) return null
        const subject = file.name.replace(/\.xls[xm]$/i, '')
        return async () => importXlsxAsDocument(haven, new Uint8Array(await file.arrayBuffer()), subject, language)
      })
    case 'save':
      return explicitSave()
    case 'export-xlsx': {
      if (!activeSessionId || welcomeShown) return
      await waitForSave()
      const session = sessionFor(activeSessionId)
      // The title may have changed in the properties since the session opened.
      download(session.bytes, storedDocumentOf(activeSessionId)?.loaded.subject ?? session.name)
      return
    }
    case 'export-csv':
    case 'print':
      if (activeSessionId && !welcomeShown) devHooks.menu(command)
      return
    case 'revisions':
      return browseRevisions(haven)
    case 'refresh': {
      const id = windows.active?.id
      if (!id || welcomeShown) return
      // An older revision: back to the current state.
      if (!historicalOf && !(await leaveWorkbook())) return
      if (historicalOf) rememberView()
      pendingView = windows.active?.view ? { id, view: windows.active.view } : null
      openInEditor(id)
      return
    }
    case 'close': {
      const id = windows.active?.id
      if (!id || welcomeShown || !(await leaveWorkbook())) return
      const next = windows.close(id)
      if (next) openInEditor(next.id)
      else showWelcomeScreen()
      return
    }
    case 'list':
      if (await leaveWorkbook()) showWelcomeScreen()
      return
    case 'properties': {
      if (!stored) return
      const properties = await editProperties(await readProperties(haven, stored.loaded.id), strings)
      if (!properties) return
      await writeProperties(haven, stored.loaded.id, properties)
      stored.loaded.subject = properties.subject
      windows.rename(stored.loaded.id, properties.subject)
      syncToolbar()
      // Our own write must not look like someone else's change.
      await changedSinceLoad(haven, stored.loaded)
      return
    }
  }
}

/** Window menu: shows another open workbook (the shown one is saved first). */
async function switchWindow(id: string): Promise<void> {
  if (id === windows.active?.id && !welcomeShown) return
  if (await leaveWorkbook()) showWorkbook(id)
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
    const workbook = activeWorkbook()
    const sheet = activeSessionId !== before ? workbook?.getSheetByName(sheetName) : null
    if (sheet) sheet.activate()
    if (sheet || Date.now() - started > 20_000) clearInterval(timer)
  }, 250)
}

function createApi(haven: HavenConnection | null): Partial<DesktopApi> {
  return {
    getLanguage: async () => language as Awaited<ReturnType<DesktopApi['getLanguage']>>,
    getTheme: async () => toUiTheme(haven?.context.theme.mode),
    // Haven's language switches without a reload; so does the editor.
    onLanguageChanged: (handler) =>
      haven?.session.onLocaleChange((locale) => {
        language = toLang(locale)
        welcomeLanguage = locale.slice(0, 2).toLowerCase()
        toolbar?.update({ t: toolbarStrings(welcomeLanguage), w: welcomeStrings(welcomeLanguage) })
        handler(language as Parameters<typeof handler>[0])
      }) ?? (() => {}),
    onThemeChanged: (handler) => haven?.session.onThemeChange((theme) => handler(toUiTheme(theme.mode))) ?? (() => {}),
    // Inside Haven the picker opens right away instead of an empty demo grid.
    hasQueuedWorkbook: async () => haven !== null,
    consumeNewBlankWorkbook: async () => false,
    getAutoSaveDefault: async () => ({ on: haven !== null, updatedAt: 0 }),
    onAutoSaveDefaultChanged: () => () => {},
    async createDocument(request) {
      if (!documentSink) {
        report('createDocument')
        return { ok: false, error: 'Creating files is not available in the browser.' }
      }
      documentSink(request)
      return { ok: true }
    },
    // CSV of the active sheet as GenOffice serialized it (displayed values).
    async exportCsv(request) {
      const blob = new Blob([request.content], { type: 'text/csv;charset=utf-8' })
      const link = document.createElement('a')
      link.href = URL.createObjectURL(blob)
      link.download = /\.csv$/i.test(request.fileName) ? request.fileName : `${request.fileName}.csv`
      link.click()
      setTimeout(() => URL.revokeObjectURL(link.href), 10_000)
      return { canceled: false, path: link.download }
    },
    // No printer list and no PDF renderer in a browser: both go through the
    // print dialog, which saves PDFs as well.
    async printWorkbook(request) {
      await printHtml(request)
      return { ok: true }
    },
    async exportPdf(request) {
      await printHtml(request)
      return { canceled: false, path: request.fileName }
    },
    // AutoSave already writes every change to MindooDB; a separate crash
    // recovery copy has no place to go in the browser.
    writeWorkbookRecovery: async () => ({ ok: false }),
    notifyPendingEdits(count) {
      pendingEdits = count
      syncToolbar()
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
    async readWorkbookRange(request) {
      const result = (await (await loadXlsxEngine()).readRange(request)) as { cells?: Array<{ row: number; column: number; arrayRef?: string }> }
      rememberArrayFormulas(request.sheetId, result.cells)
      // A late read of the previous workbook must not count.
      if (request.sessionId === activeSessionId) editorSessionId = request.sessionId
      return result as never
    },
    async readWorkbookFormulas(request) {
      const result = (await (await loadXlsxEngine()).readFormulaCells(request)) as { cells?: Array<{ row: number; column: number; arrayRef?: string }> }
      rememberArrayFormulas(request.sheetId, result.cells)
      return result as never
    },
    findWorkbookCells: async (request) => (await loadXlsxEngine()).findCells(request) as never,
    readWorkbookRowOutline: async (request) => (await loadXlsxEngine()).readRowOutline(request) as never,
    // Threaded comments are signed with the Haven user's name.
    getUserDisplayName: async () => haven?.context.user.username || 'TeamGrid',
    onUserDisplayNameChanged: () => () => {},
    // No printer list in a browser: GenOffice falls back to the system print dialog.
    listPrinters: async () => [],
    openSourcePath: async () => {},
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
  installExcelFunctions()
  installTextValues()
  installImplicitIntersection((sheetId, row, column) => arrayFormulaCells.has(`${sheetId}:${row}:${column}`))
  installUniverAccess()
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
    window.addEventListener('teamgrid:autosave', (event) => {
      autoSaveOn = (event as CustomEvent<boolean>).detail
      syncToolbar()
    })
    toolbar = mountToolbar(
      {
        t: toolbarStrings(welcomeLanguage),
        w: welcomeStrings(welcomeLanguage),
        revisionsLabel: historyAvailable ? historyStrings(welcomeLanguage).browseRevisions : undefined,
        hasWorkbook: false,
        title: '',
        canCreate: connected.canWrite && !timeTravelDate,
        windows: [],
        pending: false,
        saving: false,
        iosMultitasking: connected.context.uiPreferences?.iosMultitaskingOptimized ?? false,
        readOnly: false,
      },
      {
        onCommand: (command) => {
          onToolbarCommand(connected, command).catch((error) => console.error(`[newui] ${command} failed`, error))
        },
        onSwitch: (id) => {
          switchWindow(id).catch((error) => console.error('[newui] switching workbooks failed', error))
        },
      },
    )
    connected.session.onUiPreferencesChange?.((preferences) =>
      toolbar?.update({ iosMultitasking: preferences.iosMultitaskingOptimized }),
    )
    showTimeTravelBanner()
    registerTeamgridAgentTools(connected.session, {
      editorReady: editorShowsActiveFile,
      databaseId: () => connected.databaseId,
      canCreate: () => connected.canWrite && !timeTravelDate,
      openWorkbook() {
        if (!activeSessionId) return null
        // An older revision opens as a plain file session, without a stored document.
        if (historicalOf) {
          return { workbookId: historicalOf, title: sessionFor(activeSessionId).name, readOnly: true, revisionOf: historicalOf }
        }
        const stored = storedDocumentOf(activeSessionId)
        if (!stored) return null
        return {
          workbookId: stored.loaded.id,
          title: stored.loaded.subject,
          readOnly: !connected.canWrite || timeTravelDate !== null,
        }
      },
      async listWorkbooks() {
        return (await listWorkbooks(connected)).map((entry) => ({
          workbookId: entry.id,
          title: entry.subject,
          template: entry.istemplate,
          ...(entry.updatedAt ? { updatedAt: entry.updatedAt } : {}),
          ...(entry.legacy ? { legacy: true } : {}),
        }))
      },
      open: (id) => openForAgent(id),
      async create({ title, templateId, xlsx }) {
        const id = xlsx
          ? await importXlsxAsDocument(connected, xlsx, title, language)
          : templateId
            ? await createFromTemplate(connected, templateId, title, language)
            : await createEmptyDocument(connected, title, 'Tabelle1')
        await openForAgent(id)
        return id
      },
      save: saveForAgent,
      async exportXlsx() {
        if (!activeSessionId) throw new Error('No workbook is open.')
        await waitForSave()
        const session = sessionFor(activeSessionId)
        return { name: storedDocumentOf(activeSessionId)?.loaded.subject ?? session.name, bytes: session.bytes }
      },
      exportCsv: csvForAgent,
    })
  }
  window.addEventListener(
    'keydown',
    (event) => {
      if (!(event.metaKey || event.ctrlKey) || event.altKey) return
      const action = MENU_SHORTCUTS[event.shiftKey ? event.key.toUpperCase() : event.key.toLowerCase()]
      if (!action) return
      event.preventDefault()
      // Inside Haven, Open and Save go through the toolbar's flow (save first, Window menu).
      if (haven && (action === 'open' || action === 'save')) {
        const connected = haven
        onToolbarCommand(connected, action).catch((error) => console.error(`[newui] ${action} failed`, error))
        return
      }
      devHooks.menu(action)
    },
    true,
  )
}
