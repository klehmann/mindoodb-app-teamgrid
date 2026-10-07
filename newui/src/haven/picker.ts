// Workbook picker shown when the editor asks for a file inside Haven: the
// stored workbooks of the database, plus "new" and "import xlsx". Plain DOM,
// so it needs nothing from the editor's React tree; colors come from the
// suite's theme tokens.
import type { WorkbookSummary } from './connection'

export type PickerChoice = { kind: 'open'; id: string } | { kind: 'new' } | { kind: 'import' } | null

const STYLE = `
.tg-picker-backdrop{position:fixed;inset:0;z-index:10000;display:flex;align-items:center;justify-content:center;background:color-mix(in srgb,var(--text) 30%,transparent)}
.tg-picker{width:min(520px,calc(100vw - 32px));max-height:min(640px,calc(100vh - 32px));display:flex;flex-direction:column;background:var(--surface);color:var(--text);border:1px solid var(--border);border-radius:8px;box-shadow:0 8px 32px color-mix(in srgb,var(--text) 25%,transparent);font:14px/1.4 system-ui,sans-serif}
.tg-picker h2{margin:0;padding:16px 20px 8px;font-size:16px;font-weight:600}
.tg-picker ul{list-style:none;margin:0;padding:4px 8px;overflow:auto;flex:1;min-height:80px}
.tg-picker li button{all:unset;box-sizing:border-box;width:100%;padding:8px 12px;border-radius:6px;cursor:pointer;display:flex;justify-content:space-between;gap:12px}
.tg-picker li button:hover,.tg-picker li button:focus-visible{background:var(--hover)}
.tg-picker .tg-date{color:var(--text-secondary);font-size:12px;white-space:nowrap}
.tg-picker .tg-empty{padding:16px 12px;color:var(--text-secondary)}
.tg-picker footer{display:flex;gap:8px;justify-content:flex-end;padding:12px 20px;border-top:1px solid var(--border)}
.tg-picker footer button{padding:6px 14px;border-radius:6px;border:1px solid var(--border-strong);background:var(--surface);color:var(--text);cursor:pointer;font:inherit}
.tg-picker footer button:hover{background:var(--hover)}
.tg-picker footer .tg-primary{background:var(--accent);border-color:var(--accent);color:var(--surface)}
`

let styleInstalled = false

export function pickWorkbook(workbooks: readonly WorkbookSummary[], canCreate: boolean): Promise<PickerChoice> {
  if (!styleInstalled) {
    const style = document.createElement('style')
    style.textContent = STYLE
    document.head.append(style)
    styleInstalled = true
  }
  return new Promise((resolve) => {
    const backdrop = document.createElement('div')
    backdrop.className = 'tg-picker-backdrop'
    const dialog = document.createElement('div')
    dialog.className = 'tg-picker'
    dialog.setAttribute('role', 'dialog')
    dialog.setAttribute('aria-modal', 'true')
    const title = document.createElement('h2')
    title.textContent = 'Arbeitsmappe öffnen'
    dialog.append(title)

    const finish = (choice: PickerChoice) => {
      backdrop.remove()
      document.removeEventListener('keydown', onKey, true)
      resolve(choice)
    }
    const onKey = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        event.stopPropagation()
        finish(null)
      }
    }

    const list = document.createElement('ul')
    if (workbooks.length === 0) {
      const empty = document.createElement('li')
      empty.className = 'tg-empty'
      empty.textContent = 'Noch keine Arbeitsmappen in dieser Datenbank.'
      list.append(empty)
    }
    for (const workbook of workbooks) {
      const item = document.createElement('li')
      const button = document.createElement('button')
      const name = document.createElement('span')
      name.textContent = workbook.subject
      button.append(name)
      if (workbook.updatedAt) {
        const date = document.createElement('span')
        date.className = 'tg-date'
        date.textContent = new Date(workbook.updatedAt).toLocaleString()
        button.append(date)
      }
      button.addEventListener('click', () => finish({ kind: 'open', id: workbook.id }))
      item.append(button)
      list.append(item)
    }
    dialog.append(list)

    const footer = document.createElement('footer')
    const button = (label: string, choice: PickerChoice, primary = false) => {
      const element = document.createElement('button')
      element.type = 'button'
      element.textContent = label
      if (primary) element.className = 'tg-primary'
      element.addEventListener('click', () => finish(choice))
      footer.append(element)
    }
    button('Abbrechen', null)
    if (canCreate) {
      button('xlsx importieren …', { kind: 'import' })
      button('Neue Arbeitsmappe', { kind: 'new' }, true)
    }
    dialog.append(footer)
    backdrop.append(dialog)
    backdrop.addEventListener('click', (event) => {
      if (event.target === backdrop) finish(null)
    })
    document.addEventListener('keydown', onKey, true)
    document.body.append(backdrop)
    ;(list.querySelector('button') ?? footer.querySelector('button'))?.focus()
  })
}
