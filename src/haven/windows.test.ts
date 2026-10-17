import { describe, expect, it } from 'vitest'

import { WorkbookWindows } from './windows'

describe('WorkbookWindows', () => {
  it('lists workbooks in the order they were opened and marks the shown one', () => {
    const windows = new WorkbookWindows()
    windows.show('a', 'Budget')
    windows.show('b', 'Plan')
    windows.show('a', 'Budget 2027')
    expect(windows.list().map((window) => window.title)).toEqual(['Budget 2027', 'Plan'])
    expect(windows.active?.id).toBe('a')
  })

  it('keeps where a workbook was left when it comes forward again', () => {
    const windows = new WorkbookWindows()
    windows.show('a', 'Budget')
    windows.remember('a', { sheetName: 'Q2', range: 'B2:C4' })
    windows.show('b', 'Plan')
    expect(windows.show('a', 'Budget').view).toEqual({ sheetName: 'Q2', range: 'B2:C4' })
  })

  it('closing the shown workbook brings the most recently shown one forward', () => {
    const windows = new WorkbookWindows()
    windows.show('a', 'A')
    windows.show('b', 'B')
    windows.show('c', 'C')
    windows.show('a', 'A')
    expect(windows.close('a')?.id).toBe('c')
    expect(windows.active).toBeUndefined()
    expect(windows.close('c')?.id).toBe('b')
    expect(windows.close('b')).toBeUndefined()
    expect(windows.list()).toEqual([])
  })

  it('the welcome screen hides the shown workbook but keeps every window', () => {
    const windows = new WorkbookWindows()
    windows.show('a', 'A')
    windows.hide()
    expect(windows.active).toBeUndefined()
    expect(windows.list().map((window) => window.id)).toEqual(['a'])
  })
})
