// TeamGrid's file menu: a "File" button in the space GenOffice's ribbon tab
// row keeps free on the left, with what the desktop app has in its native
// menu bar plus Haven's own document actions.
import type { WelcomeStrings } from './welcome-strings'

export type FileAction = 'new' | 'open' | 'template' | 'import' | 'export' | 'properties' | 'close'

const STYLE = `
.tg-file{position:fixed;top:4px;left:8px;z-index:8000;display:inline-flex;align-items:center;gap:4px;height:28px;padding:0 10px;border:1px solid transparent;border-radius:6px;background:none;color:var(--text);font:600 13px/1 system-ui,sans-serif;cursor:pointer}
.tg-file:hover,.tg-file[aria-expanded="true"]{background:var(--hover);border-color:var(--border)}
.tg-file-menu{position:fixed;top:36px;left:8px;z-index:8001;min-width:220px;padding:4px;background:var(--surface);color:var(--text);border:1px solid var(--border);border-radius:8px;box-shadow:0 6px 24px color-mix(in srgb,var(--text) 20%,transparent);font:13px/1.4 system-ui,sans-serif}
.tg-file-menu button{all:unset;box-sizing:border-box;display:block;width:100%;padding:7px 12px;border-radius:5px;cursor:pointer}
.tg-file-menu button:hover,.tg-file-menu button:focus-visible{background:var(--hover)}
.tg-file-menu hr{border:0;border-top:1px solid var(--border);margin:4px 0}
`

/** Installs the File button and menu; the returned function switches its language. */
export function installFileMenu(
  initial: WelcomeStrings,
  canWrite: boolean,
  onAction: (action: FileAction) => void,
): (strings: WelcomeStrings) => void {
  let strings = initial
  const style = document.createElement('style')
  style.textContent = STYLE
  document.head.append(style)
  const trigger = document.createElement('button')
  trigger.type = 'button'
  trigger.className = 'tg-file'
  trigger.setAttribute('aria-haspopup', 'menu')
  trigger.setAttribute('aria-expanded', 'false')
  document.body.append(trigger)

  const dots = (label: string) => label.replace(/\.\.\.$/, '…')
  const entries = (): (readonly [FileAction, string] | null)[] => [
    ...(canWrite ? ([['new', `${strings.newDocument}…`]] as const) : []),
    ['open', `${strings.openDocument}…`],
    ...(canWrite ? ([['template', dots(strings.newFromTemplate)]] as const) : []),
    null,
    ...(canWrite ? ([['import', dots(strings.importXlsx)]] as const) : []),
    ['export', strings.exportXlsx],
    null,
    ['properties', `${strings.propertiesTitle}…`],
    null,
    ['close', strings.close],
  ]

  let menu: HTMLElement | null = null
  const close = () => {
    menu?.remove()
    menu = null
    trigger.setAttribute('aria-expanded', 'false')
    document.removeEventListener('pointerdown', onOutside, true)
    document.removeEventListener('keydown', onKey, true)
  }
  const onOutside = (event: Event) => {
    if (!menu?.contains(event.target as Node) && event.target !== trigger) close()
  }
  const onKey = (event: KeyboardEvent) => {
    if (event.key === 'Escape') close()
  }
  trigger.addEventListener('click', () => {
    if (menu) return close()
    menu = document.createElement('div')
    menu.className = 'tg-file-menu'
    menu.setAttribute('role', 'menu')
    for (const entry of entries()) {
      if (!entry) {
        menu.append(document.createElement('hr'))
        continue
      }
      const item = document.createElement('button')
      item.type = 'button'
      item.setAttribute('role', 'menuitem')
      item.textContent = entry[1]
      item.addEventListener('click', () => {
        close()
        onAction(entry[0])
      })
      menu.append(item)
    }
    document.body.append(menu)
    trigger.setAttribute('aria-expanded', 'true')
    document.addEventListener('pointerdown', onOutside, true)
    document.addEventListener('keydown', onKey, true)
    ;(menu.querySelector('button') as HTMLElement | null)?.focus()
  })

  const setStrings = (next: WelcomeStrings) => {
    strings = next
    trigger.textContent = `${strings.file} ▾`
    close()
  }
  setStrings(initial)
  return setStrings
}
