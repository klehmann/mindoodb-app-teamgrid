// Fill series of month and weekday names in TeamGrid's languages.
//
// Univer's auto-fill only knows English and Chinese names (LOOP_SERIES in
// @univerjs/sheets), so "Januar" is copied instead of continued. Its English
// lists also overlap ("May" is in both the short and the long month list),
// which turns January…May into "Jun". This rule runs before Univer's
// (priority 850 > 800): it knows long and short names in TeamGrid's eight
// languages, picks the list that holds every source cell (the UI language
// first, long names before short ones) and keeps the writing (JANUAR → FEBRUAR).
import { Direction, Tools, type ICellData, type Nullable, type Univer } from '@univerjs/core'
import { IAutoFillService } from '@univerjs/sheets'

/** TeamGrid's UI languages (welcome-strings.ts). */
const LANGUAGES = ['de', 'en', 'fr', 'it', 'es', 'nl', 'nb', 'pl'] as const
const RULE_TYPE = 'teamgridLocalSeries'

interface SeriesList {
  key: string
  language: string
  long: boolean
  names: string[]
  /** Lower-case names, for matching. */
  folded: string[]
}

const fold = (value: string) => value.trim().toLocaleLowerCase()

function intlNames(language: string, kind: 'month' | 'weekday', width: 'long' | 'short'): string[] {
  const format = new Intl.DateTimeFormat(language, { [kind]: width, timeZone: 'UTC' })
  // Months of 2021; weekdays from Monday 4 January 2021.
  const dates =
    kind === 'month'
      ? Array.from({ length: 12 }, (_, month) => new Date(Date.UTC(2021, month, 15)))
      : Array.from({ length: 7 }, (_, day) => new Date(Date.UTC(2021, 0, 4 + day)))
  // "Jan." and "Mo." as people type them: without the dot.
  return dates.map((date) => format.format(date).replace(/\.$/, ''))
}

/** Excel's German abbreviations, which differ from Intl's ("Mrz", "Okt"). */
const EXTRA_LISTS: Array<{ language: string; names: string[] }> = [
  { language: 'de', names: ['Jan', 'Feb', 'Mrz', 'Apr', 'Mai', 'Jun', 'Jul', 'Aug', 'Sep', 'Okt', 'Nov', 'Dez'] },
]

let cachedLists: SeriesList[] | null = null

/** Every list, without duplicates (a language's long and short names can coincide). */
export function seriesLists(): SeriesList[] {
  if (cachedLists) return cachedLists
  const lists: SeriesList[] = []
  const seen = new Set<string>()
  const add = (key: string, language: string, long: boolean, names: string[]) => {
    const folded = names.map(fold)
    const id = folded.join('|')
    if (seen.has(id) || new Set(folded).size !== folded.length) return
    seen.add(id)
    lists.push({ key, language, long, names, folded })
  }
  for (const language of LANGUAGES) {
    for (const kind of ['month', 'weekday'] as const) {
      add(`${language}-${kind}-long`, language, true, intlNames(language, kind, 'long'))
      add(`${language}-${kind}-short`, language, false, intlNames(language, kind, 'short'))
    }
  }
  for (const extra of EXTRA_LISTS) add(`${extra.language}-extra-${lists.length}`, extra.language, false, extra.names)
  cachedLists = lists
  return lists
}

/**
 * The list holding all `values`, or null: the UI language's first, then long
 * names before short ones, then the order above.
 */
export function chooseSeries(values: readonly string[], uiLanguage: string): SeriesList | null {
  const wanted = values.map(fold)
  if (wanted.some((value) => !value)) return null
  const candidates = seriesLists().filter((list) => wanted.every((value) => list.folded.includes(value)))
  if (!candidates.length) return null
  const rank = (list: SeriesList) => (list.language === uiLanguage ? 0 : 2) + (list.long ? 0 : 1)
  return [...candidates].sort((left, right) => rank(left) - rank(right))[0]!
}

/** JANUAR → upper case, januar → lower case, otherwise as the list writes it. */
function writing(values: readonly string[]): (name: string) => string {
  const letters = values.filter((value) => /\p{L}/u.test(value))
  if (letters.length && letters.every((value) => value === value.toLocaleUpperCase() && value !== value.toLocaleLowerCase())) {
    return (name) => name.toLocaleUpperCase()
  }
  if (letters.length && letters.every((value) => value === value.toLocaleLowerCase() && value !== value.toLocaleUpperCase())) {
    return (name) => name.toLocaleLowerCase()
  }
  return (name) => name
}

const text = (cell: Nullable<ICellData>) => (typeof cell?.v === 'string' ? cell.v : '')

/** A copy of `cell` as a plain value (no formula, no rich text). */
function plainCopy(cell: Nullable<ICellData>, value: string): ICellData {
  const copy = (cell ? Tools.deepClone(cell) : {}) as ICellData & { custom?: unknown }
  delete copy.f
  delete copy.si
  delete copy.p
  delete copy.custom
  copy.v = value
  return copy
}

/**
 * The cells after `data` (in fill order) continuing its series, `len` of them;
 * cells that repeat it when the steps between them are not equal.
 */
export function fillSeries(
  data: readonly Nullable<ICellData>[],
  len: number,
  reverse: boolean,
  uiLanguage: string,
): ICellData[] {
  const source = reverse ? [...data].reverse() : [...data]
  const values = source.map(text)
  const list = chooseSeries(values, uiLanguage)
  const n = list?.names.length ?? 0
  const indices = list ? values.map((value) => list.folded.indexOf(fold(value))) : []
  const steps = indices.slice(1).map((index, i) => (((index - indices[i]!) % n) + n) % n)
  const equal = steps.every((step) => step === steps[0])
  const result: ICellData[] = []
  for (let k = 1; k <= len; k++) {
    const template = source[(k - 1) % source.length]
    if (!list || !equal) {
      result.push(plainCopy(template, values[(k - 1) % values.length]!))
      continue
    }
    const step = steps.length ? steps[0]! : reverse ? -1 : 1
    const index = ((indices[indices.length - 1]! + step * k) % n + n) % n
    result.push(plainCopy(template, writing(values)(list.names[index]!)))
  }
  return reverse ? result.reverse() : result
}

/** Sets `window.teamGridOnUniver` (patches/sheets-univer-hook.patch), which registers the rule. */
export function installFillSeries(uiLanguage: () => string): void {
  ;(window as unknown as { teamGridOnUniver?: (univer: Univer) => void }).teamGridOnUniver = (univer) => {
    const rule = {
      type: RULE_TYPE,
      priority: 850,
      match: (cell: Nullable<ICellData>) => chooseSeries([text(cell)], uiLanguage()) !== null,
      isContinue: (prev: { type?: string; cellData?: Nullable<ICellData> }, cur: Nullable<ICellData>) =>
        prev.type === RULE_TYPE && chooseSeries([text(prev.cellData), text(cur)], uiLanguage()) !== null,
      applyFunctions: {
        SERIES: (dataWithIndex: { data: Nullable<ICellData>[] }, len: number, direction: Direction) =>
          fillSeries(dataWithIndex.data, len, direction === Direction.UP || direction === Direction.LEFT, uiLanguage()),
      },
    }
    // The service comes with the sheets plugin, which may still be starting.
    const started = Date.now()
    const register = () => {
      try {
        univer.__getInjector().get(IAutoFillService).registerRule(rule as never)
      } catch (error) {
        if (Date.now() - started < 20_000) setTimeout(register, 100)
        else console.warn('[newui] could not add the localized fill series', error)
      }
    }
    register()
  }
}
