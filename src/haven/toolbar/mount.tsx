// Mounts the app toolbar above GenOffice's shell (its own React root, outside
// the editor's tree) and keeps it in step with the shim's state.
import { useSyncExternalStore } from 'react'
import { createRoot } from 'react-dom/client'

import type { ToolbarStrings } from '../toolbar-strings'
import type { WelcomeStrings } from '../welcome-strings'
import { AppToolbar, type ToolbarCommand, type ToolbarState } from './AppToolbar'

export interface ToolbarView extends ToolbarState {
  t: ToolbarStrings
  w: WelcomeStrings
}

export interface ToolbarHandle {
  update: (patch: Partial<ToolbarView>) => void
  get: () => ToolbarView
}

export function mountToolbar(
  initial: ToolbarView,
  handlers: { onCommand: (command: ToolbarCommand) => void; onSwitch: (id: string) => void },
): ToolbarHandle {
  let state = initial
  const listeners = new Set<() => void>()
  const subscribe = (listener: () => void) => {
    listeners.add(listener)
    return () => listeners.delete(listener)
  }

  const host = document.createElement('div')
  host.id = 'tg-toolbar'
  document.body.prepend(host)
  document.body.classList.add('tg-has-toolbar')
  // Screens of our own that cover the editor start below the toolbar.
  new ResizeObserver(() => {
    document.documentElement.style.setProperty('--tg-toolbar-height', `${host.offsetHeight}px`)
  }).observe(host)

  function Host() {
    const view = useSyncExternalStore(subscribe, () => state)
    return <AppToolbar {...view} onCommand={handlers.onCommand} onSwitch={handlers.onSwitch} />
  }
  createRoot(host).render(<Host />)

  return {
    update(patch) {
      state = { ...state, ...patch }
      listeners.forEach((listener) => listener())
    },
    get: () => state,
  }
}
