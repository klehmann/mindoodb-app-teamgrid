/**
 * Excel's criteria for COUNTIF, as its results in IronCalc's countif.xlsx show them
 * (and as Excel's documentation describes):
 *
 * - `=` (or no operator) with a number matches numbers equal to it and text that reads
 *   as that number; `<>` with a number excludes only the equal numbers.
 * - `<`, `<=`, `>`, `>=` compare numbers with numbers and text with text (case-insensitive);
 *   other kinds never match.
 * - Text criteria match text case-insensitively with the wildcards `*`, `?` and `~`.
 * - `TRUE`/`FALSE` match booleans; an error (or its text, `#N/A`) matches that error.
 * - `""` and `=` match empty cells and empty text; `<>` matches every cell that is not
 *   empty (empty text counts as filled); `<>x` also counts empty cells.
 * - An empty criteria cell is the number 0.
 */

export type CriteriaCell =
  | { kind: 'empty' }
  | { kind: 'number'; value: number }
  | { kind: 'text'; value: string }
  | { kind: 'boolean'; value: boolean }
  | { kind: 'error'; value: string }

type Operator = '=' | '<>' | '<' | '<=' | '>' | '>='

const ERRORS = new Set(['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#SPILL!', '#CALC!', '#GETTING_DATA'])

/** A number as Excel reads it in a criterion or a text cell: `12`, `-1.5e3`, `50%`. */
export function readNumber(text: string): number | undefined {
  const trimmed = text.trim()
  if (trimmed === '') return undefined
  const percent = trimmed.endsWith('%')
  const body = percent ? trimmed.slice(0, -1) : trimmed
  if (!/^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i.test(body)) return undefined
  const value = Number(body)
  return percent ? value / 100 : value
}

function wildcard(pattern: string): RegExp {
  let source = ''
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i]!
    if (ch === '~' && i + 1 < pattern.length && '*?~'.includes(pattern[i + 1]!)) source += `\\${pattern[++i]}`
    else if (ch === '*') source += '[\\s\\S]*'
    else if (ch === '?') source += '[\\s\\S]'
    else source += ch.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&')
  }
  return new RegExp(`^${source}$`, 'i')
}

const compare = (a: number | string, b: number | string, op: Operator) =>
  op === '<' ? a < b : op === '<=' ? a <= b : op === '>' ? a > b : a >= b

/**
 * A matcher for one criterion, or undefined for criteria this implementation leaves to
 * Univer (dates and times written as text).
 */
export function criteriaMatcher(criterion: CriteriaCell): ((cell: CriteriaCell) => boolean) | undefined {
  if (criterion.kind === 'empty') return criteriaMatcher({ kind: 'number', value: 0 })
  if (criterion.kind === 'number') return (cell) => cell.kind === 'number' && cell.value === criterion.value || (cell.kind === 'text' && readNumber(cell.value) === criterion.value)
  if (criterion.kind === 'boolean') return (cell) => cell.kind === 'boolean' && cell.value === criterion.value
  if (criterion.kind === 'error') return (cell) => cell.kind === 'error' && cell.value === criterion.value

  const match = /^(<>|<=|>=|=|<|>)?([\s\S]*)$/.exec(criterion.value)!
  const op = (match[1] ?? '=') as Operator
  const operand = match[2]!
  // Dates and times ("1/1/2020", "10:30") need Excel's date parsing.
  if (/\d[/:]\d|\d-\w{3}-\d/.test(operand) && readNumber(operand) === undefined) return undefined

  const isEmpty = (cell: CriteriaCell) => cell.kind === 'empty' || (cell.kind === 'text' && cell.value === '')
  if (operand === '') {
    if (op === '=') return isEmpty
    if (op === '<>') return (cell) => cell.kind !== 'empty'
    return () => false
  }
  const upper = operand.toUpperCase()
  const number = readNumber(operand)
  const equal: (cell: CriteriaCell) => boolean =
    number !== undefined
      ? (cell) => (cell.kind === 'number' && cell.value === number) || (op === '=' && cell.kind === 'text' && readNumber(cell.value) === number)
      : upper === 'TRUE' || upper === 'FALSE'
        ? (cell) => cell.kind === 'boolean' && cell.value === (upper === 'TRUE')
        : ERRORS.has(upper)
          ? (cell) => cell.kind === 'error' && cell.value === upper
          : ((pattern) => (cell: CriteriaCell) => cell.kind === 'text' && pattern.test(cell.value))(wildcard(operand))
  if (op === '=') return equal
  if (op === '<>') return (cell) => !equal(cell)
  if (number !== undefined) return (cell) => cell.kind === 'number' && compare(cell.value, number, op)
  if (upper === 'TRUE' || upper === 'FALSE') {
    const flag = upper === 'TRUE' ? 1 : 0
    return (cell) => cell.kind === 'boolean' && compare(Number(cell.value), flag, op)
  }
  const lower = operand.toLowerCase()
  return (cell) => cell.kind === 'text' && cell.value !== '' && compare(cell.value.toLowerCase(), lower, op)
}
