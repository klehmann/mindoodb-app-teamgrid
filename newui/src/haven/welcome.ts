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

// Buttons and type follow TeamEdit / classic TeamGrid (PrimeVue Aura with
// the Mindoo preset): flat 2px corners, Mindoo blue for the primary action,
// Aura's surface colors for secondary ones, the gold focus ring.
const STYLE = `
:root{--tg-primary:#1f3a8a;--tg-primary-hover:#243c8f;--tg-primary-active:#182a63;--tg-on-primary:#ffffff;--tg-secondary-bg:#f4f6f8;--tg-secondary-hover:#e8eaee;--tg-secondary-active:#d5d9e0;--tg-secondary-text:#475569;--tg-focus-ring:rgba(212,160,23,.38);--tg-muted:#667795;--tg-radius:2px}
:root[data-theme="dark"]{--tg-secondary-bg:#27272a;--tg-secondary-hover:#3f3f46;--tg-secondary-active:#52525b;--tg-secondary-text:#d4d4d8;--tg-muted:#a7b4cf}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){--tg-secondary-bg:#27272a;--tg-secondary-hover:#3f3f46;--tg-secondary-active:#52525b;--tg-secondary-text:#d4d4d8;--tg-muted:#a7b4cf}}
.tg-welcome{position:fixed;inset:0;z-index:9000;display:flex;align-items:center;justify-content:center;padding:1.25rem 1.5rem;background:var(--surface);color:var(--text);font:16px/1.55 Figtree,Inter,ui-sans-serif,system-ui,-apple-system,"Segoe UI",sans-serif}
.tg-welcome__inner{max-width:36rem;display:grid;gap:.85rem;text-align:center}
.tg-welcome h1{margin:0;font-size:clamp(1.5rem,3.5vw,2.4rem);line-height:1.15;font-weight:700}
.tg-welcome p{margin:0;color:var(--tg-muted)}
.tg-welcome__actions{display:flex;flex-wrap:wrap;gap:.55rem;justify-content:center}
.tg-btn{display:inline-flex;align-items:center;justify-content:center;gap:.5rem;padding:.5rem .75rem;border-radius:var(--tg-radius);border:1px solid var(--tg-secondary-bg);background:var(--tg-secondary-bg);color:var(--tg-secondary-text);font:500 1rem/1.2 Figtree,Inter,ui-sans-serif,system-ui,-apple-system,"Segoe UI",sans-serif;cursor:pointer;transition:background-color .2s,border-color .2s,color .2s}
.tg-btn:hover{background:var(--tg-secondary-hover);border-color:var(--tg-secondary-hover)}
.tg-btn:active{background:var(--tg-secondary-active);border-color:var(--tg-secondary-active)}
.tg-btn:focus-visible{outline:2px solid var(--tg-focus-ring);outline-offset:2px}
.tg-btn--primary{background:var(--tg-primary);border-color:var(--tg-primary);color:var(--tg-on-primary)}
.tg-btn--primary:hover{background:var(--tg-primary-hover);border-color:var(--tg-primary-hover)}
.tg-btn--primary:active{background:var(--tg-primary-active);border-color:var(--tg-primary-active)}
.tg-btn svg{width:1rem;height:1rem;flex:none}
.tg-dialog-backdrop{position:fixed;inset:0;z-index:10000;display:flex;align-items:center;justify-content:center;background:color-mix(in srgb,var(--text) 30%,transparent)}
.tg-dialog{width:min(520px,calc(100vw - 32px));max-height:min(640px,calc(100vh - 32px));display:flex;flex-direction:column;background:var(--surface);color:var(--text);border:1px solid var(--border);border-radius:4px;box-shadow:0 8px 32px color-mix(in srgb,var(--text) 25%,transparent);font:14px/1.4 Figtree,Inter,ui-sans-serif,system-ui,-apple-system,"Segoe UI",sans-serif}
.tg-dialog h2{margin:0;padding:16px 20px 8px;font-size:16px;font-weight:600}
.tg-dialog ul{list-style:none;margin:0;padding:4px 8px;overflow:auto;flex:1;min-height:80px}
.tg-dialog li button{all:unset;box-sizing:border-box;width:100%;padding:8px 12px;border-radius:var(--tg-radius);cursor:pointer;display:flex;justify-content:space-between;gap:12px}
.tg-dialog li button:hover,.tg-dialog li button:focus-visible{background:var(--hover)}
.tg-dialog .tg-date{color:var(--text-secondary);font-size:12px;white-space:nowrap}
.tg-dialog .tg-empty{padding:16px 12px;color:var(--text-secondary)}
.tg-dialog label{display:block;padding:8px 20px 4px;color:var(--text-secondary)}
.tg-dialog input{box-sizing:border-box;width:calc(100% - 40px);margin:0 20px 16px;padding:8px 10px;border:1px solid var(--border-strong);border-radius:var(--tg-radius);background:var(--surface);color:var(--text);font:inherit}
.tg-dialog input:focus-visible,.tg-dialog textarea:focus-visible{outline:2px solid var(--tg-focus-ring);outline-offset:1px}
.tg-dialog textarea{box-sizing:border-box;width:calc(100% - 40px);margin:0 20px 4px;padding:8px 10px;border:1px solid var(--border-strong);border-radius:var(--tg-radius);background:var(--surface);color:var(--text);font:inherit;min-height:84px;resize:vertical}
.tg-dialog .tg-hint{margin:0 20px 12px;color:var(--text-secondary);font-size:12px}
.tg-dialog .tg-check{display:flex;align-items:center;gap:8px;padding:4px 20px 16px;color:var(--text)}
.tg-dialog .tg-check input{width:auto;margin:0}
.tg-dialog footer{display:flex;gap:8px;justify-content:flex-end;padding:12px 20px;border-top:1px solid var(--border)}
.tg-dialog footer .tg-btn{font-size:.875rem;padding:.4rem .7rem}
`

const ICONS = {
  import: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3v12M7 10l5 5 5-5"/><path d="M5 21h14"/></svg>',
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

export interface DocumentProperties {
  subject: string
  tags: string[]
  istemplate: boolean
}

/** The workbook's title, tags and template flag; resolves with the edited values or null. */
export function editProperties(initial: DocumentProperties, strings: WelcomeStrings): Promise<DocumentProperties | null> {
  return dialog<DocumentProperties>(strings.propertiesTitle, strings, (done) => {
    const body = element('div')
    const titleLabel = element('label', undefined, strings.titleLabel)
    const title = element('input')
    title.placeholder = strings.titlePlaceholder
    title.value = initial.subject
    titleLabel.htmlFor = title.id = 'tg-title-input'
    const tagsLabel = element('label', undefined, strings.tagsLabel)
    const tags = element('textarea')
    tags.placeholder = strings.tagsPlaceholder
    tags.value = initial.tags.join('\n')
    tagsLabel.htmlFor = tags.id = 'tg-tags-input'
    const hint = element('div', 'tg-hint', strings.tagsHint)
    const check = element('label', 'tg-check')
    const template = element('input')
    template.type = 'checkbox'
    template.id = 'tg-template-input'
    template.checked = initial.istemplate
    check.append(template, document.createTextNode(strings.useAsTemplate))
    body.append(titleLabel, title, tagsLabel, tags, hint, check)
    const save = button(strings.save, 'tg-btn--primary')
    save.addEventListener('click', () =>
      done({
        subject: title.value.trim() || strings.untitled,
        tags: [...new Set(tags.value.split('\n').map((tag) => tag.trim()).filter(Boolean))],
        istemplate: template.checked,
      }),
    )
    return { body, actions: [save] }
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
  inner.append(actions)
  page.append(inner)
  document.body.append(page)

  return new Promise((resolve) => {
    const newButton = button(`${strings.newDocument}…`, 'tg-btn--primary', ICONS.new)
    const openButton = button(strings.openDocument, '', ICONS.open)
    const templateButton = button(strings.newFromTemplate.replace(/\.\.\.$/, '…'), '', ICONS.template)
    const importButton = button(strings.importXlsx.replace(/\.\.\.$/, '…'), '', ICONS.import)
    if (options.canCreate) actions.append(newButton)
    actions.append(openButton)
    if (options.canCreate) actions.append(templateButton, importButton)

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
