// TeamGrid's agent tools around a stand-in for the editor (window.teamGridWorkbookTools).
import { MindooDBAppAgentToolError } from 'mindoodb-app-sdk'
import { afterEach, describe, expect, it, vi } from 'vitest'

import { createTeamGridAgentTools, type TeamGridAgentDeps } from './agent-tools'

const win = globalThis as unknown as { window?: unknown; teamGridWorkbookTools?: unknown }
win.window ??= globalThis

function setup(overrides: Partial<TeamGridAgentDeps> = {}) {
  let open: { id: string; title: string } | null = null
  const deps: TeamGridAgentDeps = {
    canWrite: true,
    listWorkbooks: async () => [
      { id: 'wb1', subject: 'Budget', istemplate: false },
      { id: 'tpl', subject: 'Offer template', istemplate: true },
    ],
    openWorkbook: () => open,
    open: vi.fn((id: string) => {
      open = { id, title: id === 'wb1' ? 'Budget' : 'New' }
    }),
    createEmpty: vi.fn(async () => 'wb2'),
    createFromTemplate: vi.fn(async () => 'wb3'),
    save: vi.fn(async () => {}),
    close: vi.fn(async () => {
      open = null
    }),
    menu: vi.fn(),
    reopening: () => false,
    ...overrides,
  }
  const execute = vi.fn((name: string) =>
    name === 'propose_operations'
      ? { output: 'Applied 1 change', mutated: true, summary: 'ok' }
      : { output: `${name} result`, mutated: false, summary: 'ok' },
  )
  win.teamGridWorkbookTools = { tools: [], hasWorkbook: () => open !== null, execute }
  const tools = new Map(createTeamGridAgentTools(deps).map((tool) => [tool.name, tool]))
  const call = (name: string, input: Record<string, unknown> = {}) => tools.get(name)!.execute(input)
  return { deps, tools, call, execute }
}

afterEach(() => {
  delete win.teamGridWorkbookTools
})

describe('TeamGrid agent tools', () => {
  it("offers GenOffice's workbook tools for the open workbook and its own for finding and opening", () => {
    const { tools } = setup()
    const scope = (name: string) => (tools.get(name) as { scope?: string } | undefined)?.scope
    expect(scope('workbook_context')).toBe('document')
    expect(scope('read_range')).toBe('document')
    expect(scope('apply_operations')).toBe('document')
    expect(scope('load_guide')).toBe('document')
    expect(scope('workbooks_list')).toBeUndefined()
    expect(tools.has('create_document')).toBe(false)
    expect(tools.has('propose_operations')).toBe(false)
    expect(tools.get('apply_operations')!.description).toContain('teamgrid_workbook_context')
    expect(tools.get('read_range')!.annotations?.readOnlyHint).toBe(true)
  })

  it('asks for a workbook first, then reads and applies through the editor', async () => {
    const { call, execute } = setup()
    await expect(call('read_range', { range: 'A1:B2' })).rejects.toMatchObject({ code: 'INVALID_STATE' })
    expect(await call('workbook_open', { workbookId: 'wb1' })).toMatchObject({ opened: { id: 'wb1', title: 'Budget' } })
    expect(await call('read_range', { range: 'A1:B2' })).toEqual({ result: 'read_range result' })
    expect(await call('apply_operations', { operations: [], summary: 'x' })).toMatchObject({ result: 'Applied 1 change', changed: true })
    expect(execute).toHaveBeenCalledWith('propose_operations', { operations: [], summary: 'x' })
  })

  it('reports editor errors as invalid input and refuses writes without permission', async () => {
    const { call, deps } = setup({ canWrite: false })
    deps.open('wb1')
    win.teamGridWorkbookTools = {
      tools: [],
      hasWorkbook: () => true,
      execute: () => ({ output: 'Cannot parse range: ZZ', isError: true, mutated: false, summary: '' }),
    }
    const failure = call('read_range', { range: 'ZZ' })
    await expect(failure).rejects.toBeInstanceOf(MindooDBAppAgentToolError)
    await expect(failure).rejects.toMatchObject({ code: 'INVALID_INPUT', message: 'Cannot parse range: ZZ' })
    await expect(call('apply_operations', { operations: [] })).rejects.toMatchObject({ code: 'NOT_ALLOWED' })
    await expect(call('workbook_create', { title: 'X' })).rejects.toMatchObject({ code: 'NOT_ALLOWED' })
  })

  it('lists, creates from a template, undoes and closes', async () => {
    const { call, deps } = setup()
    expect(((await call('workbooks_list', { templates: true })) as { workbooks: unknown }).workbooks).toEqual([{ workbookId: 'tpl', title: 'Offer template', template: true }])
    expect(await call('workbook_create', { title: 'Offer Mindoo', templateId: 'tpl' })).toMatchObject({ opened: { id: 'wb3' } })
    expect(deps.createFromTemplate).toHaveBeenCalledWith('tpl', 'Offer Mindoo')
    await call('undo')
    expect(deps.menu).toHaveBeenCalledWith('undo')
    expect(await call('workbook_close')).toEqual({ closed: { id: 'wb3', title: 'New' } })
    await expect(call('workbook_open', { workbookId: 'nope' })).rejects.toMatchObject({ code: 'NOT_FOUND' })
  })
})
