// Time travel (the toolbar's "Current · Saved" status): a slider over the
// history of the open workbook, or of the sheet in view, as in TeamSlides.
// Moving it shows that version in the editor, read-only; "Now" (the right end)
// is the workbook as it is.
import { useCallback, useEffect, useRef, useState, type CSSProperties } from 'react'

import type { HistoryStrings } from '../history-strings'
import './history.css'

export type HistoryScope = 'workbook' | 'sheet'

export interface HistoryPoint {
  /** The version is every document as of this time (ms). */
  at: number
  author?: string
}

export interface TimeTravelApi {
  /** The sheet in view (stored id and name), for the sheet scope. */
  activeSheet: () => { id: string; name: string } | undefined
  /** Earlier versions, oldest first; the current state is not among them. */
  timeline: (scope: HistoryScope, sheetId: string | undefined) => Promise<HistoryPoint[]>
  /** Shows the version at `at` (null: the current workbook) and resolves once the editor shows it. */
  show: (at: number | null, scope: HistoryScope, sheetId: string | undefined) => Promise<void>
}

export interface TimeTravelPanelProps {
  t: HistoryStrings
  language: string
  api: TimeTravelApi
  onClose: () => void
}

/** Slider moves closer together than this are one request. */
const SCRUB_MS = 150

export function TimeTravelPanel({ t, language, api, onClose }: TimeTravelPanelProps) {
  const [scope, setScope] = useState<HistoryScope>('workbook')
  const [sheet, setSheet] = useState(() => api.activeSheet())
  const [points, setPoints] = useState<HistoryPoint[] | null>(null)
  /** Slider position: 0 … points.length (the present). */
  const [position, setPosition] = useState(0)
  const [loading, setLoading] = useState(true)
  const [opening, setOpening] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const showingRef = useRef(false)

  const list = points ?? []
  const atNow = position >= list.length

  /** Reads the timeline (again); the editor and the slider go back to the present. */
  const load = useCallback(
    async (nextScope: HistoryScope) => {
      setLoading(true)
      setError(null)
      try {
        if (showingRef.current) {
          await api.show(null, nextScope, undefined)
          showingRef.current = false
        }
        const inView = api.activeSheet()
        setSheet(inView)
        const next = await api.timeline(nextScope, inView?.id)
        setPoints(next)
        setPosition(next.length)
      } catch (cause) {
        console.error('[newui] loading the versions failed', cause)
        setError(t.listFailed)
      } finally {
        setLoading(false)
      }
    },
    [api, t],
  )
  useEffect(() => {
    void load(scope)
  }, [load, scope])

  // At the present, the sheet scope follows the sheet in view.
  const positionRef = useRef(position)
  positionRef.current = position
  const pointsRef = useRef(list.length)
  pointsRef.current = list.length
  useEffect(() => {
    if (scope !== 'sheet') return
    const timer = setInterval(() => {
      if (showingRef.current || positionRef.current < pointsRef.current) return
      const inView = api.activeSheet()
      if (inView && inView.id !== sheet?.id) void load('sheet')
    }, 1_000)
    return () => clearInterval(timer)
  }, [api, load, scope, sheet])

  // Scrubbing: one request at a time, always for the newest position.
  const wanted = useRef<number | null>(null)
  const running = useRef(false)
  const timer = useRef<ReturnType<typeof setTimeout> | null>(null)
  const listRef = useRef(list)
  listRef.current = list
  const drain = useCallback(async () => {
    if (running.current) return
    running.current = true
    setOpening(true)
    try {
      while (wanted.current !== null) {
        const target = wanted.current
        wanted.current = null
        const point = listRef.current[target]
        if (!point && !showingRef.current) continue
        await api.show(point ? point.at : null, scope, sheet?.id)
        showingRef.current = point !== undefined
      }
      setError(null)
    } catch (cause) {
      console.error('[newui] opening the version failed', cause)
      setError(t.loadFailed)
    } finally {
      running.current = false
      setOpening(false)
    }
  }, [api, scope, sheet, t])
  const go = (next: number) => {
    const clamped = Math.max(0, Math.min(list.length, next))
    setPosition(clamped)
    wanted.current = clamped
    if (timer.current) clearTimeout(timer.current)
    timer.current = setTimeout(() => {
      timer.current = null
      void drain()
    }, SCRUB_MS)
  }

  const close = async () => {
    if (timer.current) clearTimeout(timer.current)
    wanted.current = null
    if (showingRef.current) {
      showingRef.current = false
      await api.show(null, scope, undefined).catch(() => {})
    }
    onClose()
  }
  const closeRef = useRef(close)
  closeRef.current = close
  useEffect(() => {
    const onKey = (event: KeyboardEvent) => {
      // Escape in the grid keeps its meaning (end an edit, clear the selection).
      if (event.key === 'Escape' && event.target instanceof Element && event.target.closest('.tg-history')) void closeRef.current()
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [])

  const point = atNow ? null : list[position]
  const format = new Intl.DateTimeFormat(language, { dateStyle: 'medium', timeStyle: 'medium' })
  const percent = list.length ? (position / list.length) * 100 : 100
  const versions = list.length === 1 ? t.versionsOne : t.versionsMany.replace('{n}', String(list.length))

  return (
    <section className="tg-history" role="dialog" aria-label={t.title} data-testid="teamgrid-history">
      <header className="tg-history__head">
        <svg className="tg-history__icon" width="16" height="16" viewBox="0 0 24 24" aria-hidden="true">
          <path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8" />
          <path d="M3 3v5h5M12 7v5l4 2" />
        </svg>
        <h2>{t.title}</h2>
        <button type="button" className="tg-history__close" aria-label={t.close} title={t.close} onClick={() => void close()}>
          ×
        </button>
      </header>

      <div className="tg-history__scope" role="radiogroup" aria-label={t.scopeLabel}>
        {(['workbook', 'sheet'] as const).map((value) => (
          <button
            key={value}
            type="button"
            role="radio"
            aria-checked={scope === value}
            className={scope === value ? 'is-active' : ''}
            data-testid={`teamgrid-history-scope-${value}`}
            disabled={loading || opening || (value === 'sheet' && !sheet)}
            onClick={() => setScope(value)}
          >
            {value === 'workbook' ? t.scopeWorkbook : t.scopeSheet.replace('{name}', sheet?.name ?? '…')}
          </button>
        ))}
      </div>

      {loading ? (
        <p className="tg-history__note" role="status">
          <span className="tg-history__spinner" aria-hidden="true" />
          {t.loading}
        </p>
      ) : list.length === 0 ? (
        <p className="tg-history__note">{t.empty}</p>
      ) : (
        <>
          <div className="tg-history__track">
            <button
              type="button"
              className="tg-history__step"
              aria-label={t.earlier}
              title={t.earlier}
              disabled={position <= 0}
              onClick={() => go(position - 1)}
            >
              ‹
            </button>
            <input
              type="range"
              min={0}
              max={list.length}
              step={1}
              value={position}
              aria-label={t.slider}
              aria-valuetext={point ? format.format(point.at) : t.now}
              data-testid="teamgrid-history-slider"
              style={{ '--tg-history-fill': `${percent}%` } as CSSProperties}
              onChange={(event) => go(Number(event.target.value))}
            />
            <button
              type="button"
              className="tg-history__step"
              aria-label={t.later}
              title={t.later}
              disabled={atNow}
              onClick={() => go(position + 1)}
            >
              ›
            </button>
          </div>
          <div className="tg-history__ends" aria-hidden="true">
            <span>{t.oldest}</span>
            <span>{t.now}</span>
          </div>
          <div className="tg-history__info" aria-live="polite" data-testid="teamgrid-history-info">
            <strong>{point ? format.format(point.at) : t.now}</strong>
            <span>{point ? (point.author ? t.by.replace('{author}', point.author) : '') : versions}</span>
            {opening && <span className="tg-history__spinner" role="status" aria-label={t.opening} />}
          </div>
          {!atNow && (
            <button type="button" className="tg-history__back" data-testid="teamgrid-history-now" onClick={() => go(list.length)}>
              {t.backToNow}
            </button>
          )}
        </>
      )}
      {error && (
        <p className="tg-history__error" role="alert">
          {error}
        </p>
      )}
    </section>
  )
}
