// Mounts the app toolbar above GenOffice's shell (its own React root, outside
// the editor's tree) and keeps it in step with the shim's state.
import { useSyncExternalStore } from 'react'
import { createRoot } from 'react-dom/client'

import type { HistoryStrings } from '../history-strings'
import type { ToolbarStrings } from '../toolbar-strings'
import type { WelcomeStrings } from '../welcome-strings'
import { AppToolbar, type ToolbarCommand, type ToolbarState } from './AppToolbar'
import { TimeTravelPanel, type TimeTravelApi } from './TimeTravelPanel'

export interface ToolbarView extends ToolbarState {
  t: ToolbarStrings
  w: WelcomeStrings
  h: HistoryStrings
  language: string
  /** Time travel is open, for the workbook with this id (its panel remounts per workbook). */
  historyFor: string | null
}

export interface ToolbarHandle {
  update: (patch: Partial<ToolbarView>) => void
  get: () => ToolbarView
}

export function mountToolbar(
  initial: ToolbarView,
  handlers: {
    onCommand: (command: ToolbarCommand) => void
    onSwitch: (id: string) => void
    timeTravel: TimeTravelApi
    onCloseHistory: () => void
  },
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
    return (
      <>
        <AppToolbar {...view} onCommand={handlers.onCommand} onSwitch={handlers.onSwitch} />
        {view.historyFor && (
          <TimeTravelPanel
            key={view.historyFor}
            t={view.h}
            language={view.language}
            api={handlers.timeTravel}
            onClose={handlers.onCloseHistory}
          />
        )}
      </>
    )
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
