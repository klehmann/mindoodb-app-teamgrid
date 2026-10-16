/**
 * Offers TeamGrid's agent tools to Haven for the lifetime of the Haven session
 * and keeps Haven's view of the app (`setContext`) current: the open workbook,
 * the active sheet and the selection.
 */
import type { Univer } from '@univerjs/core'
import { FUniver } from '@univerjs/core/lib/facade'
import type { MindooDBAppSession } from 'mindoodb-app-sdk'

import type { SheetsSkillDeps } from '../../vendor/genoffice/apps/sheets/src/renderer/ai/tools'
import { createTeamgridAgentTools, type TeamgridAgentDeps } from './teamgridAgentTools'

/** How often the context Haven shows agents is compared with the editor's state. */
const CONTEXT_POLL_MS = 1_000
/** Longest wait for the editor and the formula engine after an agent's save. */
const SETTLE_TIMEOUT_MS = 10_000

/**
 * With no formula run starting within this time after a save, none is coming.
 * A save reopens the merged workbook, and the engine's run over it starts about
 * a second later (measured in the mock Haven).
 */
const QUIET_MS = 2_500

/** The same for a run triggered by cells a read streamed in. */
const CALCULATION_QUIET_MS = 1_000

const delay = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms))

interface FormulaEvents {
  calculationStart(callback: () => void): unknown
  calculationEnd(callback: () => void): unknown
}

/** The AI-skill deps GenOffice's App publishes on every render (patches/sheets-agent-tools.patch). */
export function editorSkill(): SheetsSkillDeps | null {
  const get = (window as unknown as { teamGridSkillDeps?: () => SheetsSkillDeps }).teamGridSkillDeps
  return get ? get() : null
}

export function registerTeamgridAgentTools(
  session: MindooDBAppSession,
  host: Omit<TeamgridAgentDeps, 'skill' | 'takeFile' | 'provideFile'> & {
    /** The editor shows the file it was last given (not the previous or a blank workbook). */
    editorReady(): boolean
  },
): () => void {
  const agent = session.agent
  if (!agent) {
    console.info('[TeamGrid] This Haven offers no agent tools (it predates them); TeamGrid\'s AI tools stay off.')
    return () => {}
  }
  // The formula engine's runs, from the running Univer (patches/sheets-univer-hook.patch).
  let calculating = false
  let lastCalculationStart = 0
  let lastCalculationEnd = 0
  const hook = window as unknown as { teamGridOnUniver?: (univer: Univer) => void }
  const previous = hook.teamGridOnUniver
  hook.teamGridOnUniver = (univer) => {
    previous?.(univer)
    const formula = (FUniver.newAPI(univer) as unknown as { getFormula?(): FormulaEvents }).getFormula?.()
    formula?.calculationStart(() => {
      calculating = true
      lastCalculationStart = Date.now()
    })
    formula?.calculationEnd(() => {
      calculating = false
      lastCalculationEnd = Date.now()
    })
  }

  /**
   * Univer recalculates asynchronously after opening a workbook, and a save
   * reopens the merged workbook: wait until the editor shows a workbook and the
   * engine finished a run that started after `since` (or none started), so the
   * agent's next read sees formula results, its own edits included.
   */
  async function settle(since: number, quietMs = QUIET_MS): Promise<void> {
    for (;;) {
      const elapsed = Date.now() - since
      if (elapsed > SETTLE_TIMEOUT_MS) return
      const ready = host.editorReady()
      const ran = lastCalculationStart >= since && lastCalculationEnd >= lastCalculationStart
      if (ready && !calculating && (ran || elapsed > quietMs)) return
      await delay(25)
    }
  }

  const tools = createTeamgridAgentTools({
    ...host,
    // Cells streamed in by a read start their run right away; no reopen to wait for.
    calculated: (since) => settle(since, CALCULATION_QUIET_MS),
    async save() {
      const since = Date.now()
      await host.save()
      await settle(since)
    },
    async open(workbookId) {
      const since = Date.now()
      await host.open(workbookId)
      await settle(since)
    },
    async create(request) {
      const since = Date.now()
      const id = await host.create(request)
      await settle(since)
      return id
    },
    skill: editorSkill,
    takeFile: (fileRef) => agent.takeFile(fileRef),
    provideFile: (data, name, mimeType) => agent.provideFile(data, { name, mimeType }),
  })

  let disposed = false
  agent
    .registerTools(tools)
    .then((registration) => {
      console.info(
        registration.enabled
          ? `[TeamGrid] Offered ${registration.exposedNames.length} agent tools to Haven.`
          : `[TeamGrid] Declared ${registration.exposedNames.length} agent tools; Haven does not offer them yet ` +
              '(switch on "Agent tools" in Haven\'s settings and "Offer tools to AI agents" for TeamGrid).',
        registration.exposedNames,
      )
    })
    .catch((error: unknown) => console.warn('[TeamGrid] Could not offer agent tools to Haven.', error))

  // Haven's context: small and semantic; the sheet list stays behind teamgrid_get_context
  let last = ''
  const publish = () => {
    if (disposed || document.visibilityState !== 'visible') return
    const open = host.openWorkbook()
    const info = open ? editorSkill()?.getActiveSheetInfo() : undefined
    const context =
      open && info && info.mode !== 'none'
        ? {
            workbook: { workbookId: open.workbookId, title: open.title, readOnly: open.readOnly },
            activeSheet: { sheetId: info.sheetId, name: info.sheetName },
            ...(info.selection ? { selection: info.selection } : {}),
          }
        : { workbook: null }
    const json = JSON.stringify(context)
    if (json === last) return
    last = json
    void agent.setContext(context).catch(() => undefined)
  }
  const timer = setInterval(publish, CONTEXT_POLL_MS)
  publish()

  return () => {
    disposed = true
    clearInterval(timer)
    void agent.registerTools([]).catch(() => undefined)
  }
}
