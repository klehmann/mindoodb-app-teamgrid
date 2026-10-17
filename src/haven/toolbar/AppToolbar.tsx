// The app's header above GenOffice's ribbon, as in TeamSlides, TeamEdit and
// Vega: app name, File and Window menus, Refresh, the open workbook's title
// (opens its properties, × closes it) and Save or the save status at the
// right end.
import { useEffect, useRef, useState, type ReactNode } from 'react'

import type { ToolbarStrings } from '../toolbar-strings'
import type { WelcomeStrings } from '../welcome-strings'
import './toolbar.css'

export type ToolbarCommand =
  | 'new'
  | 'open'
  | 'template'
  | 'import'
  | 'save'
  | 'properties'
  | 'export-xlsx'
  | 'export-csv'
  | 'print'
  | 'refresh'
  | 'history'
  | 'close'
  | 'list'

export interface ToolbarWindow {
  id: string
  title: string
  active: boolean
}

export interface ToolbarState {
  /** A workbook is shown in the editor (not the welcome screen). */
  hasWorkbook: boolean
  title: string
  canCreate: boolean
  windows: ToolbarWindow[]
  /** Edits that are not saved yet while AutoSave is off. */
  pending: boolean
  saving: boolean
  /** iPad multitasking: Haven's window controls sit at the top left. */
  iosMultitasking: boolean
  /** The open workbook cannot be changed (no write access, time travel, an older revision). */
  readOnly: boolean
  /**
   * The database keeps history: the status at the right end opens time travel
   * (as in TeamSlides and TeamEdit), with these texts.
   */
  history?: { currentLabel: string; tip: string; historicalTip: string } | undefined
  /** An earlier version is shown (time travel): the status says so instead of "Read only". */
  travelLabel?: string | undefined
}

export interface AppToolbarProps extends ToolbarState {
  t: ToolbarStrings
  w: WelcomeStrings
  onCommand: (command: ToolbarCommand) => void
  onSwitch: (id: string) => void
}

const IS_MAC = typeof navigator !== 'undefined' && /Mac|iPhone|iPad/.test(navigator.platform)
const shortcut = (key: string, shift = false) => (IS_MAC ? `${shift ? '⇧' : ''}⌘${key}` : `Ctrl+${shift ? 'Shift+' : ''}${key}`)
const dots = (label: string) => label.replace(/\.\.\.$/, '…')

// Lucide-style line icons (ISC), 24-unit grid
function Icon({ children, size = 16 }: { children: ReactNode; size?: number }) {
  return (
    <svg
      className="tg-tb-icon"
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth={1.8}
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      {children}
    </svg>
  )
}
const FILE = <path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z M14 2v4a2 2 0 0 0 2 2h4" />
const ICONS = {
  file: <Icon>{FILE}</Icon>,
  window: (
    <Icon>
      <rect x="2" y="4" width="20" height="16" rx="2" />
      <path d="M2 8h20M6 4v4M10 4v4" />
    </Icon>
  ),
  new: (
    <Icon>
      {FILE}
      <path d="M9 15h6M12 12v6" />
    </Icon>
  ),
  open: (
    <Icon>
      <path d="m6 14 1.5-2.9A2 2 0 0 1 9.24 10H20a2 2 0 0 1 1.94 2.5l-1.54 6a2 2 0 0 1-1.95 1.5H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h3.9a2 2 0 0 1 1.69.9l.81 1.2a2 2 0 0 0 1.67.9H18a2 2 0 0 1 2 2v2" />
    </Icon>
  ),
  template: (
    <Icon>
      <rect x="8" y="8" width="12" height="13" rx="2" />
      <path d="M16 8V5a2 2 0 0 0-2-2H6a2 2 0 0 0-2 2v11a2 2 0 0 0 2 2h2" />
    </Icon>
  ),
  import: (
    <Icon>
      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M17 8l-5-5-5 5M12 3v12" />
    </Icon>
  ),
  save: (
    <Icon>
      <path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2zM17 21v-8H7v8M7 3v5h8V3" />
    </Icon>
  ),
  download: (
    <Icon>
      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M7 10l5 5 5-5M12 15V3" />
    </Icon>
  ),
  csv: (
    <Icon>
      {FILE}
      <path d="M8 13h8M8 17h8M12 11v8" />
    </Icon>
  ),
  print: (
    <Icon>
      <path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2M6 9V3a1 1 0 0 1 1-1h10a1 1 0 0 1 1 1v6" />
      <rect x="6" y="14" width="12" height="8" rx="1" />
    </Icon>
  ),
  properties: (
    <Icon>
      <path d="M21 4h-7M10 4H3M21 12h-9M8 12H3M21 20h-5M12 20H3M14 2v4M8 10v4M16 18v4" />
    </Icon>
  ),
  sheet: (
    <Icon>
      <rect x="3" y="3" width="18" height="18" rx="2" />
      <path d="M3 9h18M3 15h18M9 3v18" />
    </Icon>
  ),
  list: (
    <Icon>
      <rect x="3" y="3" width="7" height="7" rx="1" />
      <rect x="14" y="3" width="7" height="7" rx="1" />
      <rect x="14" y="14" width="7" height="7" rx="1" />
      <rect x="3" y="14" width="7" height="7" rx="1" />
    </Icon>
  ),
  check: (
    <Icon>
      <path d="M20 6 9 17l-5-5" />
    </Icon>
  ),
  close: (
    <Icon size={14}>
      <path d="M18 6 6 18M6 6l12 12" />
    </Icon>
  ),
  caret: (
    <Icon size={14}>
      <path d="m6 9 6 6 6-6" />
    </Icon>
  ),
  saved: (
    <Icon size={14}>
      <path d="M20 6 9 17l-5-5" />
    </Icon>
  ),
  refresh: (
    <Icon>
      <path d="M3 12a9 9 0 0 1 9-9 9.75 9.75 0 0 1 6.74 2.74L21 8" />
      <path d="M21 3v5h-5M21 12a9 9 0 0 1-9 9 9.75 9.75 0 0 1-6.74-2.74L3 16" />
      <path d="M8 16H3v5" />
    </Icon>
  ),
  history: (
    <Icon>
      <path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8" />
      <path d="M3 3v5h5M12 7v5l4 2" />
    </Icon>
  ),
  lock: (
    <Icon size={14}>
      <rect x="3" y="11" width="18" height="11" rx="2" />
      <path d="M7 11V7a5 5 0 0 1 10 0v4" />
    </Icon>
  ),
}

type MenuEntry =
  | 'separator'
  | {
      key: string
      label: string
      icon?: ReactNode
      shortcut?: string
      disabled?: boolean
      checked?: boolean
      onSelect: () => void
    }

interface MenuSpec {
  key: string
  label: string
  icon: ReactNode
  entries: MenuEntry[]
}

/** A menubar: click opens a menu, hovering moves to the neighbour, Escape or outside closes. */
function MenuBar({ menus }: { menus: MenuSpec[] }) {
  const [open, setOpen] = useState<string | null>(null)
  const rootRef = useRef<HTMLDivElement>(null)
  useEffect(() => {
    if (!open) return
    const onDown = (event: PointerEvent) => {
      if (!rootRef.current?.contains(event.target as Node)) setOpen(null)
    }
    const onKey = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        setOpen(null)
        rootRef.current?.querySelector<HTMLButtonElement>(`[data-menu="${open}"]`)?.focus()
      }
    }
    const onBlur = () => setOpen(null)
    window.addEventListener('pointerdown', onDown, true)
    window.addEventListener('keydown', onKey, true)
    window.addEventListener('blur', onBlur)
    return () => {
      window.removeEventListener('pointerdown', onDown, true)
      window.removeEventListener('keydown', onKey, true)
      window.removeEventListener('blur', onBlur)
    }
  }, [open])

  const focusItem = (menu: HTMLElement | null, step: 1 | -1 | 0) => {
    const items = [...(menu?.querySelectorAll<HTMLButtonElement>('button:not(:disabled)') ?? [])]
    if (!items.length) return
    const at = items.indexOf(document.activeElement as HTMLButtonElement)
    const next = step === 0 ? 0 : (at + step + items.length) % items.length
    items[next]?.focus()
  }

  return (
    <div className="tg-tb-menubar" role="menubar" ref={rootRef}>
      {menus.map((menu, index) => (
        <div key={menu.key} className="tg-tb-menu-wrap">
          <button
            type="button"
            className={`tg-tb-menu-btn${open === menu.key ? ' is-open' : ''}`}
            role="menuitem"
            aria-haspopup="menu"
            aria-expanded={open === menu.key}
            data-menu={menu.key}
            data-testid={`teamgrid-menu-${menu.key}`}
            onClick={() => setOpen((current) => (current === menu.key ? null : menu.key))}
            onPointerEnter={() => setOpen((current) => (current && current !== menu.key ? menu.key : current))}
            onKeyDown={(event) => {
              if (event.key === 'ArrowDown') {
                event.preventDefault()
                setOpen(menu.key)
                requestAnimationFrame(() => focusItem(rootRef.current!.querySelector(`[data-menu-list="${menu.key}"]`), 0))
              } else if (event.key === 'ArrowRight' || event.key === 'ArrowLeft') {
                const sibling = menus[(index + (event.key === 'ArrowRight' ? 1 : menus.length - 1)) % menus.length]!
                rootRef.current?.querySelector<HTMLButtonElement>(`[data-menu="${sibling.key}"]`)?.focus()
                if (open) setOpen(sibling.key)
              }
            }}
          >
            {menu.icon}
            <span className="tg-tb-menu-label">{menu.label}</span>
            <span className="tg-tb-caret">{ICONS.caret}</span>
          </button>
          {open === menu.key && (
            <div
              className="tg-tb-dropdown"
              role="menu"
              aria-label={menu.label}
              data-menu-list={menu.key}
              onKeyDown={(event) => {
                if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
                  event.preventDefault()
                  focusItem(event.currentTarget, event.key === 'ArrowDown' ? 1 : -1)
                }
              }}
            >
              {menu.entries.map((entry, entryIndex) =>
                entry === 'separator' ? (
                  <div key={`sep-${entryIndex}`} className="tg-tb-sep" role="separator" />
                ) : (
                  <button
                    key={entry.key}
                    type="button"
                    role={entry.checked === undefined ? 'menuitem' : 'menuitemradio'}
                    aria-checked={entry.checked}
                    className="tg-tb-item"
                    data-testid={`teamgrid-menu-item-${entry.key}`}
                    disabled={entry.disabled}
                    onClick={() => {
                      setOpen(null)
                      entry.onSelect()
                    }}
                  >
                    <span className="tg-tb-item-icon">{entry.icon}</span>
                    <span className="tg-tb-item-label">{entry.label}</span>
                    {entry.shortcut && <span className="tg-tb-item-key">{entry.shortcut}</span>}
                  </button>
                ),
              )}
            </div>
          )}
        </div>
      ))}
    </div>
  )
}

export function AppToolbar(props: AppToolbarProps) {
  const { t, w, hasWorkbook, history, onCommand } = props
  const command = (key: ToolbarCommand) => () => onCommand(key)

  const fileEntries: MenuEntry[] = [
    { key: 'new', label: `${w.newDocument}…`, icon: ICONS.new, disabled: !props.canCreate, onSelect: command('new') },
    { key: 'open', label: `${w.openDocument}…`, icon: ICONS.open, shortcut: shortcut('O'), onSelect: command('open') },
    { key: 'template', label: dots(w.newFromTemplate), icon: ICONS.template, disabled: !props.canCreate, onSelect: command('template') },
    { key: 'import', label: dots(w.importXlsx), icon: ICONS.import, disabled: !props.canCreate, onSelect: command('import') },
    'separator',
    ...(props.readOnly
      ? []
      : ([
          { key: 'save', label: t.save, icon: ICONS.save, shortcut: shortcut('S'), disabled: !hasWorkbook, onSelect: command('save') },
        ] satisfies MenuEntry[])),
    { key: 'properties', label: `${w.propertiesTitle}…`, icon: ICONS.properties, disabled: !hasWorkbook, onSelect: command('properties') },
    'separator',
    { key: 'export-xlsx', label: w.exportXlsx, icon: ICONS.download, disabled: !hasWorkbook, onSelect: command('export-xlsx') },
    { key: 'export-csv', label: t.exportCsv, icon: ICONS.csv, disabled: !hasWorkbook, onSelect: command('export-csv') },
    'separator',
    { key: 'print', label: t.print, icon: ICONS.print, shortcut: shortcut('P'), disabled: !hasWorkbook, onSelect: command('print') },
  ]

  const windowEntries: MenuEntry[] = [
    ...props.windows.map(
      (window): MenuEntry => ({
        key: `window-${window.id}`,
        label: window.title || w.untitled,
        icon: window.active ? ICONS.check : ICONS.sheet,
        checked: window.active,
        onSelect: () => props.onSwitch(window.id),
      }),
    ),
    ...(props.windows.length ? (['separator'] as const) : []),
    { key: 'list', label: t.allWorkbooks, icon: ICONS.list, onSelect: command('list') },
    { key: 'close', label: t.closeWorkbook, icon: ICONS.close, disabled: !hasWorkbook, onSelect: command('close') },
  ]

  const menus: MenuSpec[] = [
    { key: 'file', label: w.file, icon: ICONS.file, entries: fileEntries },
    { key: 'window', label: t.window, icon: ICONS.window, entries: windowEntries },
  ]

  const saveLabel = props.saving ? t.saving : t.saveChanges

  return (
    <header className={`tg-tb${props.iosMultitasking ? ' tg-tb--ios-multitasking' : ''}`} data-testid="teamgrid-toolbar">
      <div className="tg-tb-leading">
        <span className="tg-tb-title">TeamGrid</span>
        <MenuBar menus={menus} />
        <button
          type="button"
          className="tg-tb-icon-btn"
          data-tip={t.refreshTip}
          aria-label={t.refresh}
          disabled={!hasWorkbook}
          data-testid="teamgrid-refresh"
          onClick={command('refresh')}
        >
          {ICONS.refresh}
        </button>
      </div>
      {hasWorkbook && (
        <div className="tg-tb-doc">
          <button
            type="button"
            className="tg-tb-doc-title"
            title={props.title}
            data-testid="teamgrid-doc-title"
            onClick={command('properties')}
          >
            {props.title || w.untitled}
          </button>
          <button
            type="button"
            className="tg-tb-doc-close"
            aria-label={t.closeWorkbook}
            title={t.closeWorkbook}
            data-testid="teamgrid-doc-close"
            onClick={command('close')}
          >
            {ICONS.close}
          </button>
        </div>
      )}
      <div className="tg-tb-meta">
        {hasWorkbook &&
          (props.readOnly ? (
            history ? (
              <button
                type="button"
                className="tg-tb-status tg-tb-status--readonly tg-tb-status--button"
                data-tip={props.travelLabel ? history.historicalTip : t.readOnlyHint}
                aria-label={props.travelLabel ?? t.readOnly}
                data-testid="teamgrid-read-only"
                onClick={command('history')}
              >
                {props.travelLabel ? ICONS.history : ICONS.lock}
                <span className="tg-tb-status-label">{props.travelLabel ?? t.readOnly}</span>
              </button>
            ) : (
              <span className="tg-tb-status tg-tb-status--readonly" role="status" title={t.readOnlyHint} data-testid="teamgrid-read-only">
                {ICONS.lock}
                <span className="tg-tb-status-label">{t.readOnly}</span>
              </span>
            )
          ) : props.pending || props.saving ? (
            <button
              type="button"
              className="tg-tb-save"
              title={saveLabel}
              aria-label={saveLabel}
              disabled={props.saving}
              data-testid="teamgrid-save"
              onClick={command('save')}
            >
              {ICONS.save}
              <span className="tg-tb-save-label">{saveLabel}</span>
            </button>
          ) : history ? (
            <button
              type="button"
              className="tg-tb-status tg-tb-status--button"
              data-tip={history.tip}
              aria-label={history.currentLabel}
              data-testid="teamgrid-saved"
              onClick={command('history')}
            >
              {ICONS.saved}
              <span className="tg-tb-status-label">{history.currentLabel}</span>
            </button>
          ) : (
            <span className="tg-tb-status" role="status" title={t.saved} data-testid="teamgrid-saved">
              {ICONS.saved}
              <span className="tg-tb-status-label">{t.saved}</span>
            </span>
          ))}
      </div>
    </header>
  )
}
