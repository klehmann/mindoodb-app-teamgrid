/** How a formula result read from TeamGrid is compared with the one Excel stored. */
import type { CellResult } from './workbookCells'

/** Excel keeps 15 significant digits; engines differ in the last bits of long computations. */
const RELATIVE = 1e-9
const ABSOLUTE = 1e-12

export function sameResult(got: unknown, expected: CellResult): boolean {
  if (expected === null || expected === '') return got === null || got === undefined || got === ''
  if (typeof expected === 'number') {
    if (typeof got === 'boolean' || got === null || got === undefined || got === '') return false
    const value = typeof got === 'number' ? got : Number(got)
    if (!Number.isFinite(value)) return false
    return Math.abs(value - expected) <= Math.max(ABSOLUTE, Math.abs(expected) * RELATIVE)
  }
  if (typeof expected === 'boolean') {
    return got === expected || (typeof got === 'string' && got.toUpperCase() === String(expected).toUpperCase())
  }
  return typeof got === 'string' ? got === expected : String(got ?? '') === expected
}

/** Why a cell differs, for the report: the engine lacks the function, errs, or computes another result. */
export function mismatchKind(got: unknown, expected: CellResult, expectedError: boolean): 'missing' | 'error' | 'empty' | 'value' {
  if (got === '#NAME?' && expected !== '#NAME?') return 'missing'
  if (got === null || got === undefined || got === '') return 'empty'
  if (typeof got === 'string' && /^#[A-Z0-9/!?]+[!?A]?$/.test(got) && !expectedError) return 'error'
  return 'value'
}
