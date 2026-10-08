// TeamGrid's tools for AI agents (WebMCP through Haven, prefix `teamgrid_`).
//
// The workbook tools are GenOffice's own, the ones its built-in AI uses
// (readers, operation guides and the apply path in ai/tools.ts, exposed by
// patches/sheets-agent-tools.patch), so an agent's edits land in the edit
// journal and undo history like the user's. They work on the open workbook
// (`scope: "document"`). Finding, creating and opening workbooks are TeamGrid's
// own tools (`scope: "app"`); embedded in another app only the document tools
// count.
import {
  MindooDBAppAgentToolError,
  type MindooDBAppAgentTool,
  type MindooDBAppAgentToolAnnotations,
} from 'mindoodb-app-sdk'

import { WORKBOOK_TOOLS } from '../../vendor/genoffice/apps/sheets/src/renderer/ai/tools'
import type { WorkbookSummary } from './connection'

/** What patches/sheets-agent-tools.patch puts on `window`. */
interface GenOfficeWorkbookTools {
  tools: ReadonlyArray<{ name: string; description: string; inputSchema: Record<string, unknown> }>
  hasWorkbook(): boolean
  execute(
    name: string,
    input: Record<string, unknown>,
  ):
    | { output: string; isError?: boolean; mutated: boolean; summary: string }
    | Promise<{ output: string; isError?: boolean; mutated: boolean; summary: string }>
}

export interface TeamGridAgentDeps {
  canWrite: boolean
  listWorkbooks(): Promise<WorkbookSummary[]>
  /** The stored workbook the editor shows, if any. */
  openWorkbook(): { id: string; title: string } | null
  /** Opens a stored workbook in the editor (also from the welcome screen). */
  open(id: string): void
  createEmpty(title: string): Promise<string>
  createFromTemplate(templateId: string, title: string): Promise<string>
  /** Saves pending edits now (AutoSave would anyway, a little later). */
  save(): Promise<void>
  /** Closes the workbook: back to the welcome screen. */
  close(): Promise<void>
  /** Undo / redo in the editor, like ⌘Z. */
  menu(action: 'undo' | 'redo'): void
  /** Whether the editor is between workbooks (a save reopening it, an open in progress). */
  reopening(): boolean
}

/** GenOffice's names where ours read better to an agent; create_document is not offered. */
const RENAMED: Record<string, string> = {
  get_workbook_context: 'workbook_context',
  propose_operations: 'apply_operations',
}
const SKIPPED = new Set(['create_document'])
const READ_ONLY = new Set([
  'get_workbook_context',
  'read_range',
  'aggregate_range',
  'load_guide',
  'read_formats',
  'read_sheet_features',
  'read_cells',
  'find_cells',
  'trace_precedents',
  'trace_dependents',
])

const OPEN_TIMEOUT_MS = 30_000
/** How long a call waits for the workbook while a save reopens it. */
const SWAP_WAIT_MS = 5_000
/** GenOffice reads the sheet's data extent from the opened file; a save brings it up to date. */
const STALE_EXTENT = /outside the worksheet data extent/i

/** `scope` reaches Haven from mindoodb-app-sdk 0.0.66 on; older SDKs drop it. */
type ScopedTool = MindooDBAppAgentTool & { scope?: 'app' | 'document' }

function genOfficeTools(): GenOfficeWorkbookTools | undefined {
  return (window as unknown as { teamGridWorkbookTools?: GenOfficeWorkbookTools }).teamGridWorkbookTools
}

function invalid(message: string, requiredAction?: string): never {
  throw new MindooDBAppAgentToolError('INVALID_INPUT', message, requiredAction)
}

function text(input: Record<string, unknown>, field: string): string | undefined {
  const value = input[field]
  if (value === undefined || value === null) return undefined
  if (typeof value !== 'string') invalid(`${field} must be a string.`)
  return value.trim() || undefined
}

/** Waits until the editor shows workbook `id` and its tools answer. */
async function waitForOpen(deps: TeamGridAgentDeps, id: string): Promise<boolean> {
  for (const started = Date.now(); Date.now() - started < OPEN_TIMEOUT_MS; ) {
    if (deps.openWorkbook()?.id === id && genOfficeTools()?.hasWorkbook()) return true
    await new Promise((resolve) => setTimeout(resolve, 200))
  }
  return false
}

export function createTeamGridAgentTools(deps: TeamGridAgentDeps): ScopedTool[] {
  /**
   * The editor's tools once a workbook is open. A save reopens the workbook from
   * the merged document, so a call during that swap waits a moment for it.
   */
  async function requireWorkbook(): Promise<GenOfficeWorkbookTools> {
    for (const started = Date.now(); Date.now() - started < SWAP_WAIT_MS; ) {
      const tools = genOfficeTools()
      if (tools?.hasWorkbook() && deps.openWorkbook()) return tools
      if (!deps.reopening()) break
      await new Promise((resolve) => setTimeout(resolve, 100))
    }
    const tools = genOfficeTools()
    if (!tools?.hasWorkbook() || !deps.openWorkbook()) {
      throw new MindooDBAppAgentToolError(
        'INVALID_STATE',
        'No workbook is open.',
        'open one with teamgrid_workbook_open (ids from teamgrid_workbooks_list) or create one with teamgrid_workbook_create',
      )
    }
    return tools
  }

  function requireWrite() {
    if (!deps.canWrite) throw new MindooDBAppAgentToolError('NOT_ALLOWED', 'You may only read this database.')
  }

  async function openAndReport(id: string) {
    deps.open(id)
    if (!(await waitForOpen(deps, id))) {
      throw new MindooDBAppAgentToolError('FAILED', 'The workbook did not open in time.')
    }
    return { opened: deps.openWorkbook(), next: 'teamgrid_workbook_context shows its sheets' }
  }

  // GenOffice's workbook tools, wrapped: their output is text for a model.
  const workbookTools: ScopedTool[] = genOfficeToolDefinitions().map((tool): ScopedTool => {
    const name = RENAMED[tool.name] ?? tool.name
    const readOnly = READ_ONLY.has(tool.name)
    const annotations: MindooDBAppAgentToolAnnotations = readOnly
      ? { readOnlyHint: true, untrustedContentHint: true }
      : { untrustedContentHint: true }
    return {
      name,
      description: tool.description
        .replace(/get_workbook_context/g, 'teamgrid_workbook_context')
        .replace(/propose_operations/g, 'teamgrid_apply_operations')
        .replace(/load_guide/g, 'teamgrid_load_guide')
        .replace(/read_range/g, 'teamgrid_read_range'),
      inputSchema: tool.inputSchema,
      annotations,
      scope: 'document' as const,
      async execute(input) {
        if (!readOnly) requireWrite()
        let outcome = await (await requireWorkbook()).execute(tool.name, input)
        // Cells written since the workbook opened lie outside the extent it
        // knows; after a save (which reopens it) they are inside.
        if (outcome.isError && STALE_EXTENT.test(outcome.output) && deps.canWrite) {
          await deps.save()
          outcome = await (await requireWorkbook()).execute(tool.name, input)
        }
        if (outcome.isError) throw new MindooDBAppAgentToolError('INVALID_INPUT', outcome.output)
        return {
          result: outcome.output,
          ...(outcome.mutated ? { changed: true, note: 'AutoSave writes it; teamgrid_undo takes it back.' } : {}),
        }
      },
    }
  })

  const tools: ScopedTool[] = [
    ...workbookTools,
    {
      name: 'undo',
      description: 'Take back the last change in the open workbook (like ⌘Z), the agent\'s or the user\'s.',
      inputSchema: { type: 'object', properties: {} },
      scope: 'document',
      async execute() {
        await requireWorkbook()
        requireWrite()
        deps.menu('undo')
        return { undone: true }
      },
    },
    {
      name: 'redo',
      description: 'Redo what teamgrid_undo took back.',
      inputSchema: { type: 'object', properties: {} },
      scope: 'document',
      async execute() {
        await requireWorkbook()
        requireWrite()
        deps.menu('redo')
        return { redone: true }
      },
    },
    {
      name: 'save',
      description: 'Save the open workbook now. AutoSave does it anyway a moment after each change.',
      inputSchema: { type: 'object', properties: {} },
      scope: 'document',
      async execute() {
        await requireWorkbook()
        requireWrite()
        await deps.save()
        return { saved: true }
      },
    },
    {
      name: 'workbooks_list',
      description: 'The spreadsheets in TeamGrid (title, id, whether it is a template), for teamgrid_workbook_open.',
      inputSchema: {
        type: 'object',
        properties: { query: { type: 'string', description: 'Part of the title.' }, templates: { type: 'boolean', description: 'Only templates.' } },
      },
      annotations: { readOnlyHint: true, untrustedContentHint: true },
      async execute(input) {
        const query = (text(input, 'query') ?? '').toLowerCase()
        const all = await deps.listWorkbooks()
        const workbooks = all
          .filter((workbook) => (input.templates === true ? workbook.istemplate : true))
          .filter((workbook) => !query || workbook.subject.toLowerCase().includes(query))
          .slice(0, 100)
          .map((workbook) => ({
            workbookId: workbook.id,
            title: workbook.subject,
            ...(workbook.istemplate ? { template: true } : {}),
            ...(workbook.updatedAt ? { updated: workbook.updatedAt } : {}),
          }))
        return { workbooks, open: deps.openWorkbook() }
      },
    },
    {
      name: 'workbook_open',
      description: 'Open a spreadsheet in the editor (saving the open one first). Then the workbook tools work on it.',
      inputSchema: { type: 'object', properties: { workbookId: { type: 'string' } }, required: ['workbookId'] },
      annotations: {},
      async execute(input) {
        const id = text(input, 'workbookId') ?? invalid('workbookId is required.', 'call teamgrid_workbooks_list')
        if (!(await deps.listWorkbooks()).some((workbook) => workbook.id === id)) {
          throw new MindooDBAppAgentToolError('NOT_FOUND', `No spreadsheet ${id}.`, 'call teamgrid_workbooks_list')
        }
        if (deps.openWorkbook()?.id === id) return { opened: deps.openWorkbook() }
        return await openAndReport(id)
      },
    },
    {
      name: 'workbook_create',
      description: 'Create a spreadsheet (empty, or from a template: templateId from teamgrid_workbooks_list with templates) and open it.',
      inputSchema: {
        type: 'object',
        properties: { title: { type: 'string' }, templateId: { type: 'string' } },
        required: ['title'],
      },
      annotations: {},
      async execute(input) {
        requireWrite()
        const title = text(input, 'title') ?? invalid('title is required.')
        const templateId = text(input, 'templateId')
        const id = templateId ? await deps.createFromTemplate(templateId, title) : await deps.createEmpty(title)
        return await openAndReport(id)
      },
    },
    {
      name: 'workbook_close',
      description:
        'Close the open spreadsheet. With AutoSave on, pending changes are saved first; with AutoSave off and unsaved ' +
        'changes it refuses: call teamgrid_save, or ask the user whether to keep them.',
      inputSchema: { type: 'object', properties: {} },
      annotations: {},
      async execute() {
        const open = deps.openWorkbook()
        if (!open) return { closed: false, note: 'Nothing was open.' }
        await deps.close()
        return { closed: open }
      },
    },
  ]
  return tools
}

/** GenOffice's tool definitions (descriptions and schemas), without the ones we do not offer. */
function genOfficeToolDefinitions() {
  return WORKBOOK_TOOLS.filter((tool) => !SKIPPED.has(tool.name))
}
