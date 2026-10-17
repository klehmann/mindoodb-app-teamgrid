// The app toolbar above the ribbon: File and Window menus, several open
// workbooks, Save while AutoSave is off, back to the welcome screen.
import { test, type FrameLocator } from '@playwright/test'

import { TeamGridHost, expect } from './testHost'

interface Context {
  workbook: { workbookId: string; title: string } | null
  selection?: string
  activeSheetId?: string
}

function cells(host: TeamGridHost, values: Record<string, string | number>) {
  return host.call('apply_ops', {
    summary: 'Test data',
    operations: Object.entries(values).map(([address, value]) =>
      typeof value === 'string' && value.startsWith('=')
        ? { op: 'set_formula', sheetId: 'sheet-1', address, formula: value }
        : { op: 'set_cell', sheetId: 'sheet-1', address, value },
    ),
  })
}

/** Two stored workbooks with a few cells each; "Plan" is open. */
async function twoWorkbooks(host: TeamGridHost) {
  await host.call('workbook_create', { title: 'Budget' })
  await cells(host, { A1: 'Item', B1: 'Amount', B2: 1200, B3: 95, B4: '=SUM(B2:B3)' })
  await host.call('workbook_create', { title: 'Plan' })
  await cells(host, { A1: 'Task', B1: 'Days', B2: 5, B3: 12 })
}

const app = (host: TeamGridHost): FrameLocator => host.page.frameLocator('iframe')

async function shownTitle(host: TeamGridHost) {
  return (await host.call<Context>('get_context')).workbook?.title
}

async function pickWindow(host: TeamGridHost, title: string) {
  await app(host).getByTestId('teamgrid-menu-window').click()
  await app(host).getByRole('menuitemradio', { name: title }).click()
}

test.beforeEach(async ({ page }) => {
  const host = new TeamGridHost(page)
  await host.open()
  await twoWorkbooks(host)
})

test('the Window menu switches workbooks and keeps where each was left', async ({ page }) => {
  const host = new TeamGridHost(page)
  const frame = app(host)
  await expect(frame.getByTestId('teamgrid-doc-title')).toHaveText('Plan')
  await host.call('select_range', { sheetId: 'sheet-1', range: 'B3' })

  await frame.getByTestId('teamgrid-menu-window').click()
  await expect(frame.getByRole('menuitemradio')).toHaveText(['Budget', 'Plan'])
  await expect(frame.getByRole('menuitemradio', { name: 'Plan' })).toHaveAttribute('aria-checked', 'true')
  await frame.getByRole('menuitemradio', { name: 'Budget' }).click()
  await expect(frame.getByTestId('teamgrid-doc-title')).toHaveText('Budget')
  await expect.poll(() => shownTitle(host)).toBe('Budget')
  await expect.poll(async () => (await host.readCells('sheet-1', ['B4'])).B4?.value).toBe(1295)

  await pickWindow(host, 'Plan')
  await expect.poll(() => shownTitle(host)).toBe('Plan')
  await expect.poll(async () => (await host.call<Context>('get_context')).selection).toBe('B3')
})

test('closing a workbook shows the one before it, closing the last one the welcome screen', async ({ page }) => {
  const host = new TeamGridHost(page)
  const frame = app(host)
  await pickWindow(host, 'Budget')
  await expect.poll(() => shownTitle(host)).toBe('Budget')

  await frame.getByTestId('teamgrid-doc-close').click()
  await expect(frame.getByTestId('teamgrid-doc-title')).toHaveText('Plan')
  await expect.poll(() => shownTitle(host)).toBe('Plan')

  await frame.getByTestId('teamgrid-doc-close').click()
  await expect(frame.locator('#tg-welcome')).toBeVisible()
  await expect(frame.getByTestId('teamgrid-doc-title')).toHaveCount(0)
  await frame.getByTestId('teamgrid-menu-window').click()
  await expect(frame.getByRole('menuitemradio')).toHaveCount(0)
  // The toolbar stays above the welcome screen and its menus open over it.
  await expect(frame.getByTestId('teamgrid-menu-item-list')).toBeVisible()
})

test('"All spreadsheets" keeps the workbooks open in the Window menu', async ({ page }) => {
  const host = new TeamGridHost(page)
  const frame = app(host)
  await frame.getByTestId('teamgrid-menu-window').click()
  await frame.getByTestId('teamgrid-menu-item-list').click()
  await expect(frame.locator('#tg-welcome')).toBeVisible()

  await pickWindow(host, 'Budget')
  await expect(frame.locator('#tg-welcome')).toHaveCount(0)
  await expect.poll(() => shownTitle(host)).toBe('Budget')
})

test('with AutoSave off, Save shows and switching asks about unsaved changes', async ({ page }) => {
  const host = new TeamGridHost(page)
  const frame = app(host)
  await expect(frame.getByTestId('teamgrid-saved')).toBeVisible()
  await frame.locator('.autosave-toggle').click()
  await expect(frame.locator('.autosave-toggle')).not.toHaveClass(/\bon\b/)

  await host.call('select_range', { sheetId: 'sheet-1', range: 'A5' })
  await frame.locator('.univer-render-canvas, canvas').first().focus().catch(() => {})
  await page.keyboard.type('Unsaved')
  await page.keyboard.press('Enter')
  await expect(frame.getByTestId('teamgrid-save')).toBeVisible()

  // Cancel stays, "Don't save" drops the edit.
  await pickWindow(host, 'Budget')
  await expect(frame.getByRole('dialog')).toContainText('“Plan” has changes that are not saved yet.')
  await frame.getByRole('button', { name: 'Cancel' }).click()
  await expect(frame.getByTestId('teamgrid-doc-title')).toHaveText('Plan')

  await pickWindow(host, 'Budget')
  await frame.getByRole('button', { name: 'Don’t save' }).click()
  await expect.poll(() => shownTitle(host)).toBe('Budget')
  await pickWindow(host, 'Plan')
  await expect.poll(() => shownTitle(host)).toBe('Plan')
  expect((await host.readCells('sheet-1', ['A5'])).A5?.value ?? null).toBeNull()

  // Save in the toolbar writes the edit.
  await host.call('select_range', { sheetId: 'sheet-1', range: 'A6' })
  await page.keyboard.type('Kept')
  await page.keyboard.press('Enter')
  await frame.getByTestId('teamgrid-save').click()
  await expect(frame.getByTestId('teamgrid-saved')).toBeVisible()
  await frame.getByTestId('teamgrid-refresh').click()
  await expect.poll(async () => (await host.readCells('sheet-1', ['A6'])).A6?.value).toBe('Kept')
})

test('File menu exports the sheet as CSV', async ({ page }) => {
  const host = new TeamGridHost(page)
  const frame = app(host)
  await frame.getByTestId('teamgrid-menu-file').click()
  const download = page.waitForEvent('download')
  await frame.getByTestId('teamgrid-menu-item-export-csv').click()
  const file = await download
  expect(file.suggestedFilename()).toBe('Plan.csv')
})

test('time travel from the status shows earlier versions read-only, of the workbook or of one sheet', async ({ page }) => {
  const host = new TeamGridHost(page)
  const frame = app(host)
  // Saves more than a moment apart are versions of their own.
  await page.waitForTimeout(2_000)
  await cells(host, { A1: 'Task v2' })
  await page.waitForTimeout(2_000)
  await cells(host, { A1: 'Task v3' })

  await frame.getByTestId('teamgrid-saved').click()
  const panel = frame.getByTestId('teamgrid-history')
  await expect(panel).toBeVisible()
  await expect(frame.getByTestId('teamgrid-history-info')).toContainText('Now')
  const slider = frame.getByTestId('teamgrid-history-slider')
  const versions = Number(await slider.getAttribute('max'))
  expect(versions).toBeGreaterThanOrEqual(2)

  await panel.getByRole('button', { name: 'Earlier version' }).click()
  await expect(frame.getByTestId('teamgrid-read-only')).toContainText('Earlier version')
  await expect.poll(async () => (await host.readCells('sheet-1', ['A1'])).A1?.value).toBe('Task v2')
  expect((await host.call<Context & { workbook: { readOnly: boolean } }>('get_context')).workbook?.readOnly).toBe(true)

  await panel.getByRole('button', { name: 'Earlier version' }).click()
  await expect.poll(async () => (await host.readCells('sheet-1', ['A1'])).A1?.value).toBe('Task')

  // Back to now: editable again, the latest state.
  await frame.getByTestId('teamgrid-history-now').click()
  await expect(frame.getByTestId('teamgrid-saved')).toContainText('Current · Saved')
  await expect.poll(async () => (await host.readCells('sheet-1', ['A1'])).A1?.value).toBe('Task v3')

  // The sheet scope names the sheet in view and has the same saves here (one sheet).
  await frame.getByTestId('teamgrid-history-scope-sheet').click()
  await expect(frame.getByTestId('teamgrid-history-scope-sheet')).toHaveText('Sheet “Tabelle1”')
  await expect(frame.getByTestId('teamgrid-history-slider')).toHaveAttribute('max', String(versions))
  await frame.getByTestId('teamgrid-history-slider').fill('0')
  await expect(frame.getByTestId('teamgrid-read-only')).toContainText('Earlier version')
  await expect.poll(async () => (await host.readCells('sheet-1', ['A1'])).A1?.value).toBe('Task')

  // Closing the panel returns to the current version.
  await panel.getByRole('button', { name: 'Close' }).click()
  await expect(panel).toHaveCount(0)
  await expect(frame.getByTestId('teamgrid-saved')).toBeVisible()
  await expect.poll(async () => (await host.readCells('sheet-1', ['A1'])).A1?.value).toBe('Task v3')
})
