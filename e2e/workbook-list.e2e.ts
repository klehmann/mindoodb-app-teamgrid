// Workbook lists come from MindooDB's summary buffer (documents.query), not
// from walking every document of the database.
import { test } from '@playwright/test'

import { TeamGridHost, expect } from './testHost'

interface ListEntry {
  workbookId: string
  title: string
  legacy?: boolean
}

test('the workbook list is a summary query and offers 1.x workbooks until they are copied', async ({ page }) => {
  const host = new TeamGridHost(page)
  await host.open()
  await host.call('workbook_create', { title: 'Budget' })
  const methods = () =>
    page.evaluate(() => window.__havenTestHost!.log.requests.map((request) => `${request.method} ${JSON.stringify(request.params)}`))
  const before = (await methods()).length

  const { workbooks } = await host.call<{ workbooks: ListEntry[] }>('workbooks_list')
  expect(workbooks.map((entry) => [entry.title, entry.legacy ?? false])).toEqual([
    ['Budget', false],
    ['Umsatz 2025', true],
  ])
  const calls = (await methods()).slice(before)
  expect(calls.filter((call) => call.startsWith('documents.query')).length).toBe(2)
  expect(calls.filter((call) => call.startsWith('documents.list'))).toEqual([])

  // Opening the 1.x workbook makes a copy (found by query from then on); the list shows the copy.
  await host.call('workbook_open', { workbookId: 'teamgrid1_sample' })
  const after = await host.call<{ workbooks: ListEntry[] }>('workbooks_list')
  expect(after.workbooks.map((entry) => [entry.title, entry.legacy ?? false])).toEqual([
    ['Budget', false],
    ['Umsatz 2025', false],
  ])
})
