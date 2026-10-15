import { MindooDBAppAgentToolError, type MindooDBAppAgentTool } from 'mindoodb-app-sdk'
import { describe, expect, it, vi } from 'vitest'

import type { SheetsSkillDeps } from '../../vendor/genoffice/apps/sheets/src/renderer/ai/tools'
import { createTeamgridAgentTools, type OpenWorkbook, type TeamgridAgentDeps } from './teamgridAgentTools'

const SHEETS = [
  { id: 'sheet-1', name: 'Tabelle1', rows: 3, columns: 2 },
  { id: 'sheet-2', name: 'Data', rows: 0, columns: 0 },
]

/** A two-sheet workbook: A1 = 10, A2 = 20, A3 = =SUM(A1:A2) shown as "30". */
function fakeSkill(overrides: Partial<SheetsSkillDeps> = {}): SheetsSkillDeps {
  const cells: Record<string, { value: string | number; formula?: string; rawValue?: number }> = {
    A1: { value: '10', rawValue: 10 },
    A2: { value: '20', rawValue: 20 },
    A3: { value: '30', formula: '=SUM(A1:A2)', rawValue: 30 },
  }
  return {
    getActiveSheetInfo: () => ({
      mode: 'lazy',
      sheetId: 'sheet-1',
      sheetName: 'Tabelle1',
      knownAddresses: ['A1', 'A2', 'A3'],
      sheets: SHEETS,
      selection: 'A1',
    }),
    readCells: (addresses) => Object.fromEntries(addresses.flatMap((a) => (cells[a] ? [[a, cells[a]]] : []))) as never,
    readFormats: () => ({}),
    readSheetFeatures: () => 'No features.',
    findCells: () => ({ matches: [], truncated: false, incompleteSheets: [] }),
    selectRange: () => ({ ok: true, sheetName: 'Tabelle1' }),
    tracePrecedents: () => ({ refs: [] }),
    traceDependents: () => ({ dependents: [], truncated: false, incompleteSheets: [] }),
    proposeOperations: () => ({
      ok: true,
      plan: { cellChanges: [], formatChanges: [], sheetRenames: [], structuralChanges: [], warnings: [] } as never,
    }),
    ...overrides,
  }
}

function setup(options: { open?: OpenWorkbook | null; skill?: SheetsSkillDeps | null } = {}) {
  const open = options.open === undefined ? { workbookId: 'wb1', title: 'Budget', readOnly: false } : options.open
  const deps: TeamgridAgentDeps = {
    databaseId: () => 'teamgrid',
    skill: () => (options.skill === undefined ? fakeSkill() : options.skill),
    openWorkbook: () => open,
    canCreate: () => true,
    listWorkbooks: async () => [
      { workbookId: 'wb1', title: 'Budget', template: false },
      { workbookId: 'tpl', title: 'Invoice', template: true },
    ],
    open: vi.fn(async () => {}),
    create: vi.fn(async () => 'new-id'),
    save: vi.fn(async () => {}),
    exportXlsx: async () => ({ name: 'Budget', bytes: new Uint8Array([1, 2, 3]) }),
    exportCsv: async () => ({ name: 'Tabelle1', text: 'a,b\n' }),
    takeFile: async (ref) => new File([new Uint8Array([80, 75])], ref === 'bad' ? 'notes.txt' : 'data.xlsx'),
    provideFile: vi.fn(async (_data, name: string) => ({ fileRef: `ref:${name}`, size: 3 })),
  }
  const tools = createTeamgridAgentTools(deps)
  const tool = (name: string): MindooDBAppAgentTool => {
    const found = tools.find((t) => t.name === name)
    if (!found) throw new Error(`no tool ${name}`)
    return found
  }
  return { deps, tools, tool }
}

async function errorOf(promise: Promise<unknown>): Promise<MindooDBAppAgentToolError> {
  const error = await promise.then(
    () => null,
    (e: unknown) => e,
  )
  expect(error).toBeInstanceOf(MindooDBAppAgentToolError)
  return error as MindooDBAppAgentToolError
}

describe('TeamGrid agent tools', () => {
  it('offers the workbook tools and GenOffice readers under TeamGrid names', () => {
    const { tools } = setup()
    expect(tools.map((t) => t.name)).toEqual([
      'get_context',
      'workbooks_list',
      'workbook_open',
      'workbook_create',
      'workbook_export',
      'read_cells',
      'read_range',
      'read_formats',
      'read_sheet_features',
      'find_cells',
      'aggregate_range',
      'trace_precedents',
      'trace_dependents',
      'select_range',
      'apply_ops',
      'ops_guide',
    ])
    for (const t of tools.filter((t) => t.name !== 'ops_guide')) {
      const text = t.description + JSON.stringify(t.inputSchema)
      expect(text, t.name).not.toMatch(/get_workbook_context|propose_operations|load_guide/)
    }
  })

  it('reports the open workbook with its sheets', async () => {
    const context = (await setup().tool('get_context').execute({})) as Record<string, unknown>
    expect(context.workbook).toEqual({ workbookId: 'wb1', title: 'Budget', readOnly: false })
    expect(context.activeSheetId).toBe('sheet-1')
    expect(context.sheets).toEqual([
      { sheetId: 'sheet-1', name: 'Tabelle1', rows: 3, columns: 2, active: true },
      { sheetId: 'sheet-2', name: 'Data', rows: 0, columns: 0 },
    ])
    expect(context.overview).toContain('All sheets: Tabelle1 (id=sheet-1')
  })

  it('needs an open workbook for reads', async () => {
    const error = await errorOf(setup({ open: null }).tool('read_range').execute({ range: 'A1:B2' }))
    expect(error.code).toBe('INVALID_STATE')
  })

  it('reads cells with displayed text, stored value and formula', async () => {
    const result = await setup().tool('read_cells').execute({ addresses: ['a1', 'A3', 'B9'] })
    expect(result).toEqual({
      cells: {
        A1: { value: 10, text: '10' },
        A3: { value: 30, text: '30', formula: '=SUM(A1:A2)' },
        B9: { value: null },
      },
    })
    const bad = await errorOf(setup().tool('read_cells').execute({ addresses: ['A1:B2'] }))
    expect(bad.code).toBe('INVALID_INPUT')
  })

  it("runs GenOffice's readers", async () => {
    const result = (await setup().tool('read_range').execute({ range: 'A1:A3' })) as { result: string }
    expect(result.result).toContain('=SUM(A1:A2)')
    const outside = await errorOf(setup().tool('read_range').execute({ range: 'A1:Z99' }))
    expect(outside.code).toBe('INVALID_INPUT')
  })

  it('applies a batch through the skill and saves it', async () => {
    const proposeOperations = vi.fn(fakeSkill().proposeOperations)
    const { tool, deps } = setup({ skill: fakeSkill({ proposeOperations }) })
    const result = (await tool('apply_ops').execute({
      operations: [{ op: 'set_cell', sheetId: 'sheet-1', address: 'B1', value: 'Hi' }],
      summary: 'Greet',
    })) as { result: string }
    expect(proposeOperations).toHaveBeenCalledOnce()
    expect(result.result).toMatch(/Auto-applied/)
    expect(deps.save).toHaveBeenCalledOnce()
  })

  it('rejects malformed ops without touching the workbook', async () => {
    const proposeOperations = vi.fn(fakeSkill().proposeOperations)
    const { tool, deps } = setup({ skill: fakeSkill({ proposeOperations }) })
    const error = await errorOf(tool('apply_ops').execute({ operations: [{ op: 'explode' }], summary: 'x' }))
    expect(error.code).toBe('INVALID_INPUT')
    expect(proposeOperations).not.toHaveBeenCalled()
    expect(deps.save).not.toHaveBeenCalled()
  })

  it('refuses edits on a read-only revision', async () => {
    const { tool } = setup({ open: { workbookId: 'wb1', title: 'Budget', readOnly: true, revisionOf: 'wb1' } })
    const error = await errorOf(
      tool('apply_ops').execute({ operations: [{ op: 'clear_cell', sheetId: 'sheet-1', address: 'A1' }], summary: 'x' }),
    )
    expect(error.code).toBe('NOT_ALLOWED')
  })

  it('opens, creates and imports workbooks', async () => {
    const { tool, deps } = setup()
    await tool('workbook_open').execute({ workbookId: 'tpl' })
    expect(deps.open).toHaveBeenCalledWith('tpl')
    expect((await errorOf(tool('workbook_open').execute({ workbookId: 'nope' }))).code).toBe('NOT_FOUND')

    await tool('workbook_create').execute({ title: 'From template', templateId: 'tpl' })
    expect(deps.create).toHaveBeenLastCalledWith({ title: 'From template', templateId: 'tpl' })
    await tool('workbook_create').execute({ title: 'Imported', fileRef: 'f1' })
    expect(deps.create).toHaveBeenLastCalledWith({ title: 'Imported', xlsx: new Uint8Array([80, 75]) })
    expect((await errorOf(tool('workbook_create').execute({ title: 'x', fileRef: 'bad' }))).code).toBe('INVALID_INPUT')
    expect((await errorOf(tool('workbook_create').execute({ title: 'x', templateId: 'wb1' }))).code).toBe('NOT_FOUND')
  })

  it('exports xlsx and csv as files', async () => {
    const { tool } = setup()
    expect(await tool('workbook_export').execute({})).toMatchObject({ fileRef: 'ref:Budget.xlsx', name: 'Budget.xlsx' })
    expect(await tool('workbook_export').execute({ format: 'csv', sheetId: 'sheet-1' })).toMatchObject({
      fileRef: 'ref:Tabelle1.csv',
    })
    expect((await errorOf(tool('workbook_export').execute({ format: 'csv', sheetId: 'x' }))).code).toBe('NOT_FOUND')
  })

  it("serves GenOffice's op guides", async () => {
    const { tool } = setup()
    const index = JSON.stringify(await tool('ops_guide').execute({}))
    expect(index).toContain('format_range')
    const formatting = JSON.stringify(await tool('ops_guide').execute({ groups: ['formatting'] }))
    expect(formatting).toContain('format_range')
    expect(formatting.length).toBeGreaterThan(2000)
  })
})

describe('blocksOf', () => {
  it('loads scattered cells as one block when it is small enough, else row by row', async () => {
    const { blocksOf } = await import('./teamgridAgentTools')
    expect(blocksOf(['A1', 'C3'])).toEqual([{ startRow: 0, endRow: 2, startColumn: 0, endColumn: 2 }])
    expect(blocksOf(['A1', 'B5000'])).toEqual([
      { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
      { startRow: 4999, endRow: 4999, startColumn: 1, endColumn: 1 },
    ])
  })
})
