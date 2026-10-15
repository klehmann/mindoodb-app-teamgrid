/**
 * TeamGrid's tools for AI agents, offered through Haven (WebMCP). Reading and
 * editing the open workbook goes through GenOffice Sheets' own AI skill
 * (`ai/tools.ts`): the same readers, the same workbook DSL and the same
 * plan/apply path its built-in assistant uses, so an agent edit lands in the
 * undo history and is saved to MindooDB like any edit in the grid. Around it,
 * TeamGrid adds what only it knows: the stored workbooks, opening, creating,
 * importing and exporting them.
 *
 * Pure factory: everything app-specific comes in through `deps`, so tests can
 * call `execute` directly. `registerTeamgridAgentTools` wires it up.
 */
import {
  createAgentGuideTool,
  MindooDBAppAgentToolError,
  type MindooDBAppAgentTool,
  type MindooDBAppAgentToolErrorCode,
} from 'mindoodb-app-sdk'

import { GUIDE_CATALOG } from '../../vendor/genoffice/apps/sheets/src/renderer/ai/guides'
import {
  buildWorkbookContext,
  executeWorkbookTool,
  WORKBOOK_TOOLS,
  type SheetsSkillDeps,
  type ToolExecution,
} from '../../vendor/genoffice/apps/sheets/src/renderer/ai/tools'

/** The stored workbook the editor shows. */
export interface OpenWorkbook {
  workbookId: string
  title: string
  /** Haven's time travel, an older revision, or no write access. */
  readOnly: boolean
  /** Set while an older revision is shown (File → Browse revisions). */
  revisionOf?: string
}

export interface WorkbookListEntry {
  workbookId: string
  title: string
  template: boolean
  updatedAt?: string
  /** A TeamGrid 1.x workbook: opening it opens a copy in the current format. */
  legacy?: boolean
}

export interface TeamgridAgentDeps {
  databaseId(): string
  /** The editor's AI-skill deps; null until the editor is mounted. */
  skill(): SheetsSkillDeps | null
  openWorkbook(): OpenWorkbook | null
  canCreate(): boolean
  listWorkbooks(): Promise<WorkbookListEntry[]>
  /** Opens a stored workbook in the editor (a 1.x workbook as its copy); resolves once the editor shows it. */
  open(workbookId: string): Promise<void>
  /** Creates a workbook (blank, from a template or from xlsx bytes), opens it and returns its id. */
  create(request: { title: string; templateId?: string; xlsx?: Uint8Array }): Promise<string>
  /** Writes pending edits to MindooDB; resolves once the editor shows the saved state with formulas recalculated. */
  save(): Promise<void>
  /** The open workbook as xlsx, after saving pending edits. */
  exportXlsx(): Promise<{ name: string; bytes: Uint8Array }>
  /** One worksheet's displayed values as CSV (GenOffice's CSV export). */
  exportCsv(sheetId: string | undefined): Promise<{ name: string; text: string }>
  takeFile(fileRef: string): Promise<File>
  provideFile(data: Blob | Uint8Array, name: string, mimeType: string): Promise<{ fileRef: string; size: number }>
}

const PREFIX = 'teamgrid'
const NEXT_OPEN = `call ${PREFIX}_workbooks_list, then ${PREFIX}_workbook_open`
const NEXT_EXPORT = 'pass the fileRef to haven_files_export to save the file into the exchange folder'
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

/** GenOffice's tools that TeamGrid offers unchanged (input and output as in its assistant). */
const SKILL_READ_TOOLS = [
  'read_range',
  'read_formats',
  'read_sheet_features',
  'find_cells',
  'aggregate_range',
  'trace_precedents',
  'trace_dependents',
] as const

function fail(code: MindooDBAppAgentToolErrorCode, message: string, next?: string): never {
  throw new MindooDBAppAgentToolError(code, message, next)
}

function optionalString(input: Record<string, unknown>, key: string): string | undefined {
  const value = input[key]
  return typeof value === 'string' && value.trim() ? value.trim() : undefined
}

/** GenOffice's descriptions name its own tools; agents see TeamGrid's. */
function retarget(text: string): string {
  return text
    .replaceAll('get_workbook_context', `${PREFIX}_get_context`)
    .replaceAll('propose_operations', `${PREFIX}_apply_ops`)
    .replaceAll('load_guide', `${PREFIX}_ops_guide`)
    .replace(/\(the user can roll back with the \[Undo\] button or ⌘Z\)/, '(the user can undo it in the grid)')
    .replace(/\(undo via the side panel \[Undo\] button or ⌘Z\)/, '(the user can undo it in the grid)')
}

function skillDef(name: string) {
  const def = WORKBOOK_TOOLS.find((tool) => tool.name === name)
  if (!def) throw new Error(`GenOffice has no ${name} tool`)
  return def
}

/** A failed skill call as an agent-tool error: state problems apart from bad input. */
function skillFailure(execution: ToolExecution): never {
  const text = retarget(execution.output)
  const state = /could not be (fully )?loaded|not available|No workbook|indexing|still computing/i.test(text)
  fail(state ? 'INVALID_STATE' : 'INVALID_INPUT', text)
}

/** Sheets as agents address them: id and name (ids stay valid while the workbook is open). */
function sheetList(skill: SheetsSkillDeps) {
  const info = skill.getActiveSheetInfo()
  return info.sheets.map((sheet) => ({
    sheetId: sheet.id,
    name: sheet.name,
    ...(sheet.rows !== undefined && sheet.columns !== undefined ? { rows: sheet.rows, columns: sheet.columns } : {}),
    ...(sheet.id === info.sheetId ? { active: true } : {}),
    ...(sheet.readOnlyOversized ? { readOnly: true } : {}),
  }))
}

function signatureIndex(): string {
  return [
    `Ops for ${PREFIX}_apply_ops (load a group for its fields):`,
    'basic: set_cell {sheetId, address, value} | set_formula {sheetId, address, formula} | clear_cell {sheetId, address} | rename_sheet {sheetId, name}',
    'writing: set_range, fill_range, copy_range, convert_to_values, clear_range, find_replace',
    'formatting: format_range',
    'layout: sort_range, merge_cells, unmerge_cells, set_row_height, set_col_width, set_rows_hidden, set_cols_hidden, set_freeze, set_page_setup',
    'structure: insert_rows, delete_rows, insert_cols, delete_cols, add_sheet, delete_sheet, duplicate_sheet, set_sheet_hidden, move_sheet, protect_sheet',
    'charts: add_chart, edit_chart, delete_visual, add_sparkline, add_shape, edit_shape, add_image',
    'pivot: add_pivot, refresh_pivot',
    'table: add_table, add_table_row, add_table_column, delete_table_row, delete_table_column, delete_table',
    'data: set_hyperlink, set_filter, clear_filter, set_filter_criteria, add_conditional_format, clear_conditional_formats, set_data_validation, set_note, add_defined_name, delete_defined_name',
    'Guide-only groups: financial-formatting, data-attribution, shape-image.',
  ].join('\n')
}

export function createTeamgridAgentTools(deps: TeamgridAgentDeps): MindooDBAppAgentTool[] {
  function editor(): SheetsSkillDeps {
    const skill = deps.skill()
    if (!skill || !deps.openWorkbook() || skill.getActiveSheetInfo().mode === 'none') {
      fail('INVALID_STATE', 'No workbook is open in TeamGrid.', NEXT_OPEN)
    }
    return skill
  }

  function writable(): SheetsSkillDeps {
    const skill = editor()
    const open = deps.openWorkbook()!
    if (open.readOnly) {
      fail(
        'NOT_ALLOWED',
        open.revisionOf
          ? 'An older revision is open; it is read-only.'
          : 'This workbook is read-only here (time travel or no write access).',
        open.revisionOf ? `call ${PREFIX}_workbook_open with workbookId "${open.revisionOf}" for the current state` : undefined,
      )
    }
    return skill
  }

  async function runSkill(name: string, input: Record<string, unknown>, skill: SheetsSkillDeps): Promise<string> {
    const execution = await executeWorkbookTool({ id: `agent-${name}`, name, input }, skill)
    if (execution.isError) skillFailure(execution)
    return retarget(execution.output)
  }

  function context() {
    const open = deps.openWorkbook()
    const skill = deps.skill()
    const info = open && skill ? skill.getActiveSheetInfo() : null
    if (!open || !skill || !info || info.mode === 'none') {
      return { databaseId: deps.databaseId(), canCreate: deps.canCreate(), workbook: null }
    }
    return {
      databaseId: deps.databaseId(),
      canCreate: deps.canCreate(),
      workbook: { ...open },
      activeSheetId: info.sheetId,
      ...(info.selection ? { selection: info.selection } : {}),
      sheets: sheetList(skill),
      overview: retarget(buildWorkbookContext(skill)),
    }
  }

  const getContext: MindooDBAppAgentTool = {
    name: 'get_context',
    description:
      'What is open in TeamGrid: the workbook (id, title, read-only), its sheets with sheetId, name and data extent, ' +
      'the active sheet, the user\'s selection, merges and charts. Start here; the sheetIds are what every other ' +
      `${PREFIX} tool and every op in ${PREFIX}_apply_ops takes. Without an open workbook it says so (${NEXT_OPEN}).`,
    inputSchema: { type: 'object', properties: {} },
    annotations: { readOnlyHint: true, untrustedContentHint: true },
    execute: async () => context(),
  }

  const workbooksList: MindooDBAppAgentTool = {
    name: 'workbooks_list',
    description:
      'Lists the workbooks stored in TeamGrid\'s database: workbookId, title, last change, whether it is a template. ' +
      `Pass workbookId to ${PREFIX}_workbook_open, or a template's id as templateId to ${PREFIX}_workbook_create.`,
    inputSchema: {
      type: 'object',
      properties: {
        templates: { type: 'boolean', description: 'true lists only templates, false only workbooks; both when omitted.' },
      },
    },
    annotations: { readOnlyHint: true, untrustedContentHint: true },
    async execute(input) {
      const all = await deps.listWorkbooks()
      const templates = input.templates
      return {
        workbooks: typeof templates === 'boolean' ? all.filter((entry) => entry.template === templates) : all,
      }
    },
  }

  const workbookOpen: MindooDBAppAgentTool = {
    name: 'workbook_open',
    description:
      `Opens a stored workbook in the TeamGrid editor (id from ${PREFIX}_workbooks_list); pending edits of the ` +
      `workbook shown before are saved first. A TeamGrid 1.x workbook (legacy: true) opens as a copy in the current ` +
      `format, with its own workbookId. Returns the same as ${PREFIX}_get_context.`,
    inputSchema: {
      type: 'object',
      required: ['workbookId'],
      properties: { workbookId: { type: 'string' } },
    },
    async execute(input) {
      const workbookId = optionalString(input, 'workbookId')
      if (!workbookId) fail('INVALID_INPUT', 'Give "workbookId".', `call ${PREFIX}_workbooks_list`)
      const open = deps.openWorkbook()
      if (open?.workbookId === workbookId && !open.revisionOf && deps.skill()) return context()
      if (!(await deps.listWorkbooks()).some((entry) => entry.workbookId === workbookId)) {
        fail('NOT_FOUND', `No workbook ${workbookId} in this database.`, `call ${PREFIX}_workbooks_list`)
      }
      await deps.open(workbookId)
      return context()
    },
  }

  const workbookCreate: MindooDBAppAgentTool = {
    name: 'workbook_create',
    description:
      'Creates a workbook in TeamGrid\'s database and opens it: blank, as a copy of a template (templateId from ' +
      `${PREFIX}_workbooks_list with templates: true), or imported from an .xlsx file (fileRef from haven_files_import ` +
      `with forApp "${PREFIX}"). Returns the same as ${PREFIX}_get_context. Fill a blank one with ${PREFIX}_apply_ops.`,
    inputSchema: {
      type: 'object',
      required: ['title'],
      properties: {
        title: { type: 'string', description: 'Title of the new workbook.' },
        templateId: { type: 'string', description: 'Copy this template (optional).' },
        fileRef: { type: 'string', description: 'Import this .xlsx/.xlsm file (optional).' },
      },
    },
    async execute(input) {
      const title = optionalString(input, 'title')
      if (!title) fail('INVALID_INPUT', 'Give a "title".')
      const templateId = optionalString(input, 'templateId')
      const fileRef = optionalString(input, 'fileRef')
      if (templateId && fileRef) fail('INVALID_INPUT', 'Give either "templateId" or "fileRef", not both.')
      if (!deps.canCreate()) fail('NOT_ALLOWED', 'This TeamGrid may not create workbooks here (time travel or no write access).')
      let xlsx: Uint8Array | undefined
      if (fileRef) {
        const file = await deps.takeFile(fileRef).catch(() =>
          fail('NOT_FOUND', `File ${fileRef} is unknown or expired.`, `import it again with haven_files_import (forApp "${PREFIX}")`),
        )
        if (!/\.xls[xm]$/i.test(file.name) && file.type !== XLSX_MIME) {
          fail('INVALID_INPUT', `${file.name} is not an .xlsx or .xlsm file.`)
        }
        xlsx = new Uint8Array(await file.arrayBuffer())
      }
      if (templateId) {
        const template = (await deps.listWorkbooks()).find((entry) => entry.workbookId === templateId)
        if (!template?.template) {
          fail('NOT_FOUND', `No template ${templateId}.`, `call ${PREFIX}_workbooks_list with templates: true`)
        }
      }
      await deps.create({ title, ...(templateId ? { templateId } : {}), ...(xlsx ? { xlsx } : {}) })
      return context()
    },
  }

  const workbookExport: MindooDBAppAgentTool = {
    name: 'workbook_export',
    description:
      'Exports the open workbook as .xlsx (everything: formulas, formats, charts), or one worksheet as .csv ' +
      `(displayed values only). Returns a fileRef; ${NEXT_EXPORT}.`,
    inputSchema: {
      type: 'object',
      properties: {
        format: { type: 'string', enum: ['xlsx', 'csv'], description: 'Default xlsx.' },
        sheetId: { type: 'string', description: 'csv only: the worksheet; default the active sheet.' },
      },
    },
    annotations: { readOnlyHint: true },
    async execute(input) {
      editor()
      const format = input.format ?? 'xlsx'
      if (format !== 'xlsx' && format !== 'csv') fail('INVALID_INPUT', '"format" is "xlsx" or "csv".')
      if (format === 'xlsx') {
        const { name, bytes } = await deps.exportXlsx()
        const fileName = /\.xls[xm]$/i.test(name) ? name : `${name}.xlsx`
        const { fileRef, size } = await deps.provideFile(bytes, fileName, XLSX_MIME)
        return { fileRef, name: fileName, size, next: NEXT_EXPORT }
      }
      const sheetId = optionalString(input, 'sheetId')
      if (sheetId && !editor().getActiveSheetInfo().sheets.some((sheet) => sheet.id === sheetId)) {
        fail('NOT_FOUND', `No sheet ${sheetId}.`, `call ${PREFIX}_get_context`)
      }
      const { name, text } = await deps.exportCsv(sheetId)
      const fileName = /\.csv$/i.test(name) ? name : `${name}.csv`
      const { fileRef, size } = await deps.provideFile(new Blob([text], { type: 'text/csv' }), fileName, 'text/csv')
      return { fileRef, name: fileName, size, next: NEXT_EXPORT }
    },
  }

  const readCells: MindooDBAppAgentTool = {
    name: 'read_cells',
    description:
      'Reads scattered cells (max 100; use read_range for blocks): per address the value (numbers as numbers, dates as ' +
      'serial numbers, formula cells their result), the displayed text where it differs, and the formula. Read cells ' +
      'before overwriting them.',
    inputSchema: {
      type: 'object',
      required: ['addresses'],
      properties: {
        addresses: { type: 'array', items: { type: 'string' }, description: 'Cell addresses like ["A1","C7"].' },
        sheetId: { type: 'string', description: `Sheet id from ${PREFIX}_get_context; default the active sheet.` },
      },
    },
    annotations: { readOnlyHint: true, untrustedContentHint: true },
    async execute(input) {
      const skill = editor()
      const raw = input.addresses
      if (!Array.isArray(raw) || raw.length === 0 || raw.some((a) => typeof a !== 'string')) {
        fail('INVALID_INPUT', '"addresses" must be a non-empty list of cell addresses.')
      }
      if (raw.length > 100) fail('INVALID_INPUT', 'At most 100 addresses per call; use read_range for blocks.')
      const addresses = (raw as string[]).map((a) => a.trim().toUpperCase().replace(/\$/g, ''))
      const bad = addresses.find((a) => !/^[A-Z]{1,3}[1-9][0-9]{0,6}$/.test(a))
      if (bad) fail('INVALID_INPUT', `Not a cell address: ${bad}`)
      const sheetId = optionalString(input, 'sheetId')
      if (sheetId && !skill.getActiveSheetInfo().sheets.some((sheet) => sheet.id === sheetId)) {
        fail('NOT_FOUND', `No sheet ${sheetId}.`, `call ${PREFIX}_get_context`)
      }
      const read = skill.readCells(addresses, sheetId) as Record<
        string,
        { value: unknown; formula?: string; rawValue?: unknown }
      >
      const cells: Record<string, { value: unknown; text?: unknown; formula?: string }> = {}
      for (const address of addresses) {
        const cell = read[address]
        // GenOffice's lazy reader gives the rendered text as `value` and the model value as `rawValue`.
        const value = cell?.rawValue !== undefined ? cell.rawValue : (cell?.value ?? null)
        cells[address] = {
          value,
          ...(cell && cell.value !== value && cell.value !== null ? { text: cell.value } : {}),
          ...(cell?.formula ? { formula: cell.formula } : {}),
        }
      }
      return { cells }
    },
  }

  const readTools: MindooDBAppAgentTool[] = SKILL_READ_TOOLS.map((name) => {
    const def = skillDef(name)
    return {
      name,
      description: retarget(def.description),
      inputSchema: JSON.parse(retarget(JSON.stringify(def.inputSchema))) as Record<string, unknown>,
      annotations: { readOnlyHint: true, untrustedContentHint: true },
      async execute(input: Record<string, unknown>) {
        return { result: await runSkill(name, input, editor()) }
      },
    }
  })

  const selectRange: MindooDBAppAgentTool = {
    name: 'select_range',
    description: retarget(skillDef('select_range').description).replace(
      /; to merely point at a location, cite it as \[C42\]\(sheetnav:\/\/C42\) in your reply and let them click\./,
      '.',
    ),
    inputSchema: JSON.parse(retarget(JSON.stringify(skillDef('select_range').inputSchema))) as Record<string, unknown>,
    async execute(input) {
      return { result: await runSkill('select_range', input, editor()) }
    },
  }

  const applyOps: MindooDBAppAgentTool = {
    name: 'apply_ops',
    description:
      retarget(skillDef('propose_operations').description) +
      ` The batch is saved to MindooDB right away. Load ${PREFIX}_ops_guide groups before ops you have not sent before.`,
    inputSchema: skillDef('propose_operations').inputSchema as Record<string, unknown>,
    async execute(input) {
      const skill = writable()
      const result = await runSkill('propose_operations', input, skill)
      await deps.save()
      return { result }
    },
  }

  const opsGuide = createAgentGuideTool({
    description:
      `The op reference for ${PREFIX}_apply_ops (GenOffice Sheets' workbook DSL). Without groups: one line per op ` +
      'group. With groups: field definitions, conventions and common mistakes. Load the groups you need before a ' +
      `batch you have not sent before. The guides call the edit tool propose_operations and the context tool ` +
      `get_workbook_context; in TeamGrid they are ${PREFIX}_apply_ops and ${PREFIX}_get_context. ` +
      `Groups: ${Object.keys(GUIDE_CATALOG).join(', ')}.`,
    groups: Object.entries(GUIDE_CATALOG).map(([name, entry]) => ({
      name,
      summary: entry.description,
      content: entry.content,
    })),
    signatures: signatureIndex,
  })

  return [
    getContext,
    workbooksList,
    workbookOpen,
    workbookCreate,
    workbookExport,
    readCells,
    ...readTools,
    selectRange,
    applyOps,
    opsGuide,
  ]
}
