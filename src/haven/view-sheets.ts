// The view-sheet commands of the Data tab: add a sheet filled from a
// MindooDB virtual view, refresh the active one, change its settings.
//
// The sheet is written to the stored workbook directly (model/view-sheet.ts),
// after the editor saved what was pending; the editor then reopens the
// workbook and shows the sheet. ExcelShell (patches/sheets-view-sheets.patch)
// renders the buttons from `window.teamGridViewSheets`.
import type { MindooDBAppResolvedViewDefinition } from 'mindoodb-app-sdk'

import { createId, liveSheets, type ViewBinding, type Workbook } from '../model/schema'
import { writesFor } from '../model/sync'
import {
  buildViewSheet,
  parseCategoryPath,
  readViewEntries,
  viewTable,
  viewTitle,
  type ViewSheetSettings,
} from '../model/view-sheet'
import type { HavenConnection } from './connection'
import { loadWorkbook, writeWorkbook } from './store'
import { viewSheetStrings, type ViewSheetStrings } from './view-sheet-strings'
import { button, dialog, element } from './welcome'
import type { WelcomeStrings } from './welcome-strings'

export interface ViewSheetHost {
  haven: HavenConnection
  language: () => string
  welcomeStrings: () => WelcomeStrings
  /** The open stored workbook's id, if one is open. */
  activeWorkbookId: () => string | undefined
  /** Saves what the editor still has pending. */
  waitForSave: () => Promise<void>
  /** Reopens the workbook in the editor and shows sheet `name`. */
  reopen: (id: string, sheetName: string) => void
}

/** What ExcelShell reads from `window.teamGridViewSheets`. */
export interface ViewSheetCommands {
  labels: () => Pick<ViewSheetStrings, 'group' | 'add' | 'addTip' | 'refresh' | 'refreshTip' | 'settings' | 'settingsTip'>
  add: () => void
  refresh: () => void
  settings: () => void
}

function activeSheetName(): string | undefined {
  const api = (window as unknown as {
    __univerAPI?: { getActiveWorkbook(): { getActiveSheet(): { getSheetName(): string } } | null }
  }).__univerAPI
  return api?.getActiveWorkbook()?.getActiveSheet().getSheetName()
}

function message(text: string, strings: ViewSheetStrings, welcome: WelcomeStrings): Promise<null> {
  return dialog<null>('TeamGrid', { ...welcome, cancel: strings.close }, () => ({ body: element('p', 'tg-message', text) }))
}

/** The settings dialog; resolves with the settings or null. */
function askSettings(
  title: string,
  initial: ViewSheetSettings,
  views: readonly MindooDBAppResolvedViewDefinition[],
  takenNames: ReadonlySet<string>,
  strings: ViewSheetStrings,
  welcome: WelcomeStrings,
): Promise<ViewSheetSettings | null> {
  return dialog<ViewSheetSettings>(title, welcome, (done) => {
    const body = element('div')
    const field = <T extends HTMLElement>(label: string, control: T, id: string) => {
      const node = element('label', undefined, label)
      node.htmlFor = control.id = id
      body.append(node, control)
      return control
    }
    const name = field(strings.nameLabel, element('input'), 'tg-view-name')
    name.value = initial.name
    const view = field(strings.viewLabel, element('select'), 'tg-view-id')
    for (const candidate of views) {
      const option = element('option', undefined, viewTitle(candidate))
      option.value = candidate.id
      view.append(option)
    }
    view.value = initial.viewId
    const check = (label: string, checked: boolean) => {
      const row = element('label', 'tg-check')
      const input = element('input')
      input.type = 'checkbox'
      input.checked = checked
      row.append(input, document.createTextNode(label))
      body.append(row)
      return input
    }
    const documents = check(strings.showDocuments, initial.showDocuments)
    const categories = check(strings.showCategories, initial.showCategories)
    const root = field(strings.rootLabel, element('input'), 'tg-view-root')
    root.placeholder = strings.rootPlaceholder
    root.value = initial.rootCategoryPath.join('\\')
    const error = element('div', 'tg-error')
    body.append(error, element('div', 'tg-hint', strings.hint))
    // A new sheet is named after its view until the user names it.
    let named = initial.name !== ''
    name.addEventListener('input', () => (named = name.value.trim() !== ''))
    const suggest = () => {
      if (!named) name.value = viewTitle(views.find((candidate) => candidate.id === view.value) ?? views[0]!)
    }
    view.addEventListener('change', suggest)
    suggest()
    const apply = button(strings.apply, 'tg-btn--primary')
    const submit = () => {
      const chosen = name.value.trim()
      const problem = !chosen
        ? strings.nameMissing
        : takenNames.has(chosen.toLowerCase())
          ? strings.nameTaken
          : !documents.checked && !categories.checked
            ? strings.nothingShown
            : ''
      error.textContent = problem
      if (problem) return
      done({
        name: chosen,
        viewId: view.value,
        showDocuments: documents.checked,
        showCategories: categories.checked,
        rootCategoryPath: parseCategoryPath(root.value),
      })
    }
    apply.addEventListener('click', submit)
    name.addEventListener('keydown', (event) => {
      if (event.key === 'Enter') submit()
    })
    return { body, actions: [apply] }
  })
}

/** Fills sheet `sheetId` (new when absent) from its view and writes it. */
async function fill(
  host: ViewSheetHost,
  workbookId: string,
  sheetId: string | undefined,
  settings: ViewSheetSettings,
  strings: ViewSheetStrings,
) {
  const view = host.haven.context.views.find((candidate) => candidate.id === settings.viewId)
  if (!view) throw new Error(strings.viewGone.replace('{view}', settings.viewId))
  let entries
  try {
    const navigator = await host.haven.session.openViewNavigator(settings.viewId, {
      includeCategories: settings.showCategories,
      includeDocuments: settings.showDocuments,
      hideEmptyCategories: true,
      ...(settings.rootCategoryPath.length ? { rootCategoryPath: settings.rootCategoryPath } : {}),
    })
    try {
      entries = await readViewEntries(navigator)
    } finally {
      await navigator.dispose()
    }
  } catch (error) {
    throw new Error(strings.failed.replace('{error}', error instanceof Error ? error.message : String(error)))
  }
  const loaded = await loadWorkbook(host.haven, workbookId)
  const workbook = loaded.stored.workbook
  const id = sheetId ?? createId('s')
  const binding: ViewBinding = {
    viewId: view.id,
    viewTitle: viewTitle(view),
    showDocuments: settings.showDocuments,
    showCategories: settings.showCategories,
    rootCategoryPath: settings.rootCategoryPath,
    lastRefreshedAt: new Date().toISOString(),
  }
  const sheet = buildViewSheet(workbook.worksheetsById[id], id, settings.name, binding, viewTable(view, entries, settings))
  const next: Workbook = {
    ...workbook,
    worksheetOrder: workbook.worksheetOrder.includes(id) ? workbook.worksheetOrder : [...workbook.worksheetOrder, id],
    worksheetsById: { ...workbook.worksheetsById, [id]: sheet },
  }
  await writeWorkbook(host.haven, loaded, writesFor(loaded.stored, next, [id]))
  host.reopen(workbookId, settings.name)
}

export function viewSheetCommands(host: ViewSheetHost): ViewSheetCommands {
  const strings = () => viewSheetStrings(host.language())
  /** Runs a command on the open workbook, reporting problems in a dialog. */
  const run = async (command: (workbookId: string, workbook: Workbook) => Promise<void>) => {
    const workbookId = host.activeWorkbookId()
    if (!workbookId) return
    try {
      if (!host.haven.canWrite) throw new Error(strings().readOnly)
      await host.waitForSave()
      await command(workbookId, (await loadWorkbook(host.haven, workbookId)).stored.workbook)
    } catch (error) {
      console.warn('[newui] view sheet command failed', error)
      await message(error instanceof Error ? error.message : String(error), strings(), host.welcomeStrings())
    }
  }
  const names = (workbook: Workbook, except?: string) =>
    new Set(liveSheets(workbook).filter((sheet) => sheet.id !== except).map((sheet) => sheet.name.toLowerCase()))
  /** The active sheet, when it is a view sheet; tells the user otherwise. */
  const activeViewSheet = (workbook: Workbook) => {
    const name = activeSheetName()
    const sheet = liveSheets(workbook).find((candidate) => candidate.name === name)
    if (!sheet?.viewBinding) throw new Error(strings().notViewSheet)
    return sheet as typeof sheet & { viewBinding: ViewBinding }
  }
  const settingsOf = (sheet: { name: string; viewBinding: ViewBinding }): ViewSheetSettings => ({
    name: sheet.name,
    viewId: sheet.viewBinding.viewId,
    showDocuments: sheet.viewBinding.showDocuments,
    showCategories: sheet.viewBinding.showCategories,
    rootCategoryPath: sheet.viewBinding.rootCategoryPath,
  })

  return {
    labels: () => strings(),
    add: () =>
      void run(async (workbookId, workbook) => {
        const views = host.haven.context.views
        if (views.length === 0) throw new Error(strings().noViews)
        const settings = await askSettings(
          strings().addTitle,
          { name: '', viewId: views[0]!.id, showDocuments: true, showCategories: true, rootCategoryPath: [] },
          views,
          names(workbook),
          strings(),
          host.welcomeStrings(),
        )
        if (settings) await fill(host, workbookId, undefined, settings, strings())
      }),
    refresh: () =>
      void run(async (workbookId, workbook) => {
        const sheet = activeViewSheet(workbook)
        if (!host.haven.context.views.some((view) => view.id === sheet.viewBinding.viewId)) {
          throw new Error(strings().viewGone.replace('{view}', sheet.viewBinding.viewTitle))
        }
        await fill(host, workbookId, sheet.id, settingsOf(sheet), strings())
      }),
    settings: () =>
      void run(async (workbookId, workbook) => {
        const sheet = activeViewSheet(workbook)
        const views = host.haven.context.views
        if (views.length === 0) throw new Error(strings().noViews)
        const settings = await askSettings(
          strings().settingsTitle,
          settingsOf(sheet),
          views,
          names(workbook, sheet.id),
          strings(),
          host.welcomeStrings(),
        )
        if (settings) await fill(host, workbookId, sheet.id, settings, strings())
      }),
  }
}
