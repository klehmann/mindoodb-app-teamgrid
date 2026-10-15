// The running editor's Univer API, for TeamGrid's own commands (which sheet
// is active, showing a sheet after a reopen). GenOffice exposes it as
// `window.__univerAPI` in development builds only; this takes it from the
// Univer hook (patches/sheets-univer-hook.patch) in every build.
import type { Univer } from '@univerjs/core'
import { FUniver } from '@univerjs/core/facade'

type UniverHook = (univer: Univer) => void

let api: FUniver | undefined

/** Chains onto `window.teamGridOnUniver`; install after the other hooks. */
export function installUniverAccess(): void {
  const host = window as unknown as { teamGridOnUniver?: UniverHook }
  const previous = host.teamGridOnUniver
  host.teamGridOnUniver = (univer) => {
    previous?.(univer)
    api = FUniver.newAPI(univer)
  }
}

/** The active workbook's sheets API, once the editor runs. */
export function activeWorkbook() {
  return api?.getActiveWorkbook() ?? null
}
