// Welcome screen and its dialogs, shown inside Haven while no workbook is
// open: a short description and buttons to create a spreadsheet, open one,
// create one from a template or import an xlsx file (like TeamEdit and
// classic TeamGrid). Plain DOM, so it needs nothing from the editor's React
// tree; colors come from the suite's theme tokens.
import type { WorkbookSummary } from './connection'
import { welcomeStrings, type WelcomeStrings } from './welcome-strings'

export type WelcomeChoice =
  | { kind: 'new'; subject: string }
  | { kind: 'open'; id: string }
  | { kind: 'template'; id: string; subject: string }
  | { kind: 'import' }

const STYLE = `
.tg-welcome{position:fixed;inset:0;z-index:9000;display:flex;align-items:center;justify-content:center;padding:24px;background:var(--surface);color:var(--text);font:15px/1.5 system-ui,sans-serif}
.tg-welcome__inner{max-width:640px;text-align:center}
.tg-welcome h1{margin:0 0 16px;font-size:clamp(28px,4vw,40px);line-height:1.15;font-weight:700}
.tg-welcome p{margin:0 0 28px;color:var(--text-secondary);font-size:17px}
.tg-welcome__actions{display:flex;flex-wrap:wrap;gap:12px;justify-content:center;margin-bottom:12px}
.tg-btn{display:inline-flex;align-items:center;gap:10px;padding:10px 18px;border-radius:8px;border:1px solid var(--border-strong);background:var(--surface-subtle);color:var(--text);font:inherit;font-size:16px;cursor:pointer}
.tg-btn:hover{background:var(--hover)}
.tg-btn:focus-visible{outline:2px solid var(--accent);outline-offset:2px}
.tg-btn--primary{background:var(--accent);border-color:var(--accent);color:var(--sheets-on-accent)}
.tg-btn--primary:hover{background:var(--accent-dark,var(--accent))}
.tg-btn--link{border-color:transparent;background:none;color:var(--text-secondary);font-size:14px;padding:6px 10px}
.tg-btn svg{width:20px;height:20px;flex:none}
.tg-dialog-backdrop{position:fixed;inset:0;z-index:10000;display:flex;align-items:center;justify-content:center;background:color-mix(in srgb,var(--text) 30%,transparent)}
.tg-dialog{width:min(520px,calc(100vw - 32px));max-height:min(640px,calc(100vh - 32px));display:flex;flex-direction:column;background:var(--surface);color:var(--text);border:1px solid var(--border);border-radius:8px;box-shadow:0 8px 32px color-mix(in srgb,var(--text) 25%,transparent);font:14px/1.4 system-ui,sans-serif}
.tg-dialog h2{margin:0;padding:16px 20px 8px;font-size:16px;font-weight:600}
.tg-dialog ul{list-style:none;margin:0;padding:4px 8px;overflow:auto;flex:1;min-height:80px}
.tg-dialog li button{all:unset;box-sizing:border-box;width:100%;padding:8px 12px;border-radius:6px;cursor:pointer;display:flex;justify-content:space-between;gap:12px}
.tg-dialog li button:hover,.tg-dialog li button:focus-visible{background:var(--hover)}
.tg-dialog .tg-date{color:var(--text-secondary);font-size:12px;white-space:nowrap}
.tg-dialog .tg-empty{padding:16px 12px;color:var(--text-secondary)}
.tg-dialog label{display:block;padding:8px 20px 4px;color:var(--text-secondary)}
.tg-dialog input{box-sizing:border-box;width:calc(100% - 40px);margin:0 20px 16px;padding:8px 10px;border:1px solid var(--border-strong);border-radius:6px;background:var(--surface);color:var(--text);font:inherit}
.tg-dialog footer{display:flex;gap:8px;justify-content:flex-end;padding:12px 20px;border-top:1px solid var(--border)}
.tg-dialog footer .tg-btn{font-size:14px;padding:6px 14px}
`

const ICONS = {
  new: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5M12 11v6M9 14h6"/></svg>',
  open: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v1H3z"/><path d="M3 10h18l-2 9H5z"/></svg>',
  template: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><rect x="8" y="8" width="12" height="13" rx="2"/><path d="M16 8V5a2 2 0 0 0-2-2H6a2 2 0 0 0-2 2v11a2 2 0 0 0 2 2h2"/></svg>',
}

let styleInstalled = false

function installStyle() {
  if (styleInstalled) return
  const style = document.createElement('style')
  style.textContent = STYLE
  document.head.append(style)
  styleInstalled = true
}

function element<K extends keyof HTMLElementTagNameMap>(tag: K, className?: string, text?: string) {
  const node = document.createElement(tag)
  if (className) node.className = className
  if (text !== undefined) node.textContent = text
  return node
}

function button(label: string, className: string, icon?: string) {
  const node = element('button', `tg-btn ${className}`)
  node.type = 'button'
  if (icon) node.innerHTML = icon
  node.append(document.createTextNode(label))
  return node
}

/** A modal dialog; resolves with what `build` reports, or null on cancel/Escape. */
function dialog<T>(title: string, strings: WelcomeStrings, build: (done: (value: T | null) => void) => {
  body: HTMLElement
  actions?: HTMLButtonElement[]
}): Promise<T | null> {
  installStyle()
  return new Promise((resolve) => {
    const backdrop = element('div', 'tg-dialog-backdrop')
    const box = element('div', 'tg-dialog')
    box.setAttribute('role', 'dialog')
    box.setAttribute('aria-modal', 'true')
    const onKey = (event: KeyboardEvent) => {
      if (event.key !== 'Escape') return
      event.stopPropagation()
      done(null)
    }
    const done = (value: T | null) => {
      backdrop.remove()
      document.removeEventListener('keydown', onKey, true)
      resolve(value)
    }
    const { body, actions = [] } = build(done)
    const footer = element('footer')
    const cancel = button(strings.cancel, '')
    cancel.addEventListener('click', () => done(null))
    footer.append(cancel, ...actions)
    box.append(element('h2', undefined, title), body, footer)
    backdrop.append(box)
    backdrop.addEventListener('click', (event) => {
      if (event.target === backdrop) done(null)
    })
    document.addEventListener('keydown', onKey, true)
    document.body.append(backdrop)
    ;(box.querySelector('input, li button') as HTMLElement | null ?? cancel).focus()
  })
}

/** Lists workbooks to pick one; resolves with its id or null. */
export function chooseWorkbook(
  workbooks: readonly WorkbookSummary[],
  title: string,
  empty: string,
  strings: WelcomeStrings,
): Promise<WorkbookSummary | null> {
  return dialog<WorkbookSummary>(title, strings, (done) => {
    const list = element('ul')
    if (workbooks.length === 0) list.append(element('li', 'tg-empty', empty))
    for (const workbook of workbooks) {
      const item = element('li')
      const pick = element('button')
      pick.type = 'button'
      pick.append(element('span', undefined, workbook.subject))
      if (workbook.updatedAt) pick.append(element('span', 'tg-date', new Date(workbook.updatedAt).toLocaleString()))
      pick.addEventListener('click', () => done(workbook))
      item.append(pick)
      list.append(item)
    }
    return { body: list }
  })
}

/** Asks for a title; resolves with it (the default when left empty) or null. */
export function askTitle(title: string, initial: string, strings: WelcomeStrings): Promise<string | null> {
  return dialog<string>(title, strings, (done) => {
    const body = element('div')
    const label = element('label', undefined, strings.titleLabel)
    const input = element('input')
    input.placeholder = strings.titlePlaceholder
    input.value = initial
    label.htmlFor = input.id = 'tg-title-input'
    const submit = () => done(input.value.trim() || strings.untitled)
    input.addEventListener('keydown', (event) => {
      if (event.key === 'Enter') submit()
    })
    body.append(label, input)
    const create = button(strings.create, 'tg-btn--primary')
    create.addEventListener('click', submit)
    return { body, actions: [create] }
  })
}

export interface WelcomeOptions {
  language: string
  canCreate: boolean
  listWorkbooks: () => Promise<WorkbookSummary[]>
}

/**
 * Shows the welcome screen until the user picked something to open; it then
 * stays behind the editor's loading and is removed by `closeWelcome`.
 */
export function showWelcome(options: WelcomeOptions): Promise<WelcomeChoice> {
  installStyle()
  const strings = welcomeStrings(options.language)
  closeWelcome()
  const page = element('div', 'tg-welcome')
  page.id = 'tg-welcome'
  const inner = element('div', 'tg-welcome__inner')
  inner.append(element('h1', undefined, strings.title), element('p', undefined, strings.body))
  const actions = element('div', 'tg-welcome__actions')
  const more = element('div', 'tg-welcome__actions')
  inner.append(actions, more)
  page.append(inner)
  document.body.append(page)

  return new Promise((resolve) => {
    const newButton = button(`${strings.newDocument}…`, 'tg-btn--primary', ICONS.new)
    const openButton = button(strings.openDocument, '', ICONS.open)
    const templateButton = button(strings.newFromTemplate.replace(/\.\.\.$/, '…'), '', ICONS.template)
    const importButton = button(strings.importXlsx.replace(/\.\.\.$/, '…'), 'tg-btn--link')
    if (options.canCreate) actions.append(newButton)
    actions.append(openButton)
    if (options.canCreate) more.append(templateButton, importButton)

    newButton.addEventListener('click', async () => {
      const subject = await askTitle(strings.newTitle, '', strings)
      if (subject !== null) resolve({ kind: 'new', subject })
    })
    openButton.addEventListener('click', async () => {
      const all = await options.listWorkbooks()
      const picked = await chooseWorkbook(
        all.filter((workbook) => !workbook.istemplate),
        strings.openTitle,
        strings.empty,
        strings,
      )
      if (picked) resolve({ kind: 'open', id: picked.id })
    })
    templateButton.addEventListener('click', async () => {
      const all = await options.listWorkbooks()
      const template = await chooseWorkbook(
        all.filter((workbook) => workbook.istemplate),
        strings.templateTitle,
        strings.emptyTemplates,
        strings,
      )
      if (!template) return
      const subject = await askTitle(strings.newTitle, strings.copyOf.replace('{title}', template.subject), strings)
      if (subject !== null) resolve({ kind: 'template', id: template.id, subject })
    })
    importButton.addEventListener('click', () => resolve({ kind: 'import' }))
    ;(options.canCreate ? newButton : openButton).focus()
  })
}

export function closeWelcome(): void {
  document.getElementById('tg-welcome')?.remove()
}

export { welcomeStrings }
