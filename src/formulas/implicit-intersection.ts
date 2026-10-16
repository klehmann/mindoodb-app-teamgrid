/**
 * Implicit intersection for formulas that are no dynamic arrays, as Excel computes them.
 *
 * A formula Excel stored without the dynamic-array mark returns one value: where an
 * argument or the result is a range or an array, Excel takes the cell in the formula's
 * own row (a column range) or column (a row range), `#VALUE!` when there is none, and
 * the first element of an array that has no position (`{1,2}+1` → 2). Excel 365 shows
 * such formulas with `@`. Univer spills every array result, so these cells showed
 * `#SPILL!` (or overwrote neighbours) where Excel shows a value.
 *
 * GenOffice writes a formula as a dynamic array exactly when it calls a spilling
 * function (`spillsDynamicArray`: FILTER, SORT, SEQUENCE, UNIQUE …), and TeamGrid stores
 * and exports formulas that way. The engine follows the same rule: a formula without
 * such a call, that is not a legacy array (CSE) formula, gets the intersection of its
 * array result. Arrays an argument builds inside the formula (`SUM(A1:A3*2)`) are not
 * intersected, which differs from Excel's per-argument rule; the fidelity tests measure
 * what is left.
 */
import type { Univer } from '@univerjs/core'
import { spillsDynamicArray } from '@genoffice/xlsx-gateway/gateway/future-functions'
import {
  type ArrayValueObject,
  type BaseValueObject,
  ErrorType,
  ErrorValueObject,
  FormulaDataModel,
  IFormulaRuntimeService,
} from '@univerjs/engine-formula'

type Variant = BaseValueObject & {
  isReferenceObject(): boolean
}

interface Reference {
  getRangePosition(): { startRow: number; endRow: number; startColumn: number; endColumn: number }
  toArrayValueObject(): ArrayValueObject
  isColumn?(): boolean
  isRow?(): boolean
}

interface Area {
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
}

const columnNumber = (letters: string) => [...letters.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0)

/** The rectangular references written in a formula (A1:B9, A:A, 3:5), 0-based; text in quotes skipped. */
export function formulaAreas(formula: string): Area[] {
  const code = formula.replace(/"(?:[^"]|"")*"/g, '""')
  const areas: Area[] = []
  const pattern = /(?<![A-Za-z0-9_.$])\$?([A-Z]{1,3})\$?(\d+)?:\$?([A-Z]{1,3})\$?(\d+)?(?![A-Za-z0-9_(])|(?<![A-Za-z0-9_.$:])\$?(\d+):\$?(\d+)(?![A-Za-z0-9_.(])/gi
  for (const match of code.matchAll(pattern)) {
    if (match[5] !== undefined) {
      areas.push({ startRow: Number(match[5]) - 1, endRow: Number(match[6]) - 1, startColumn: 0, endColumn: 16_383 })
    } else if ((match[2] === undefined) === (match[4] === undefined)) {
      const whole = match[2] === undefined
      areas.push({
        startRow: whole ? 0 : Number(match[2]) - 1,
        endRow: whole ? 1_048_575 : Number(match[4]) - 1,
        startColumn: columnNumber(match[1]!) - 1,
        endColumn: columnNumber(match[3]!) - 1,
      })
    }
  }
  return areas
}

const valueError = () => ErrorValueObject.create(ErrorType.VALUE)

/** The intersection of an area with the formula's cell: the element's offset, or null. */
function intersection(area: Area, row: number, column: number): { row: number; column: number } | null {
  if (area.startColumn === area.endColumn && row >= area.startRow && row <= area.endRow) return { row: row - area.startRow, column: 0 }
  if (area.startRow === area.endRow && column >= area.startColumn && column <= area.endColumn) return { row: 0, column: column - area.startColumn }
  return null
}

/** The one value a non-dynamic formula at (row, column) returns for an array or range result. */
export function implicitIntersection(variant: Variant, formula: string, row: number, column: number): BaseValueObject {
  if (variant.isReferenceObject()) {
    const reference = variant as unknown as Reference
    const position = reference.getRangePosition()
    const at = intersection(position, row, column)
    if (!at) return valueError()
    return reference.toArrayValueObject().get(at.row, at.column) ?? valueError()
  }
  const array = variant as unknown as ArrayValueObject
  const rows = array.getRowCount()
  const columns = array.getColumnCount()
  // An array shaped like a range of the formula came from that range: intersect with it.
  // Univer cuts whole columns and rows to the sheet's size, so those match by their other side.
  for (const area of formulaAreas(formula)) {
    const areaRows = area.endRow - area.startRow + 1
    const areaColumns = area.endColumn - area.startColumn + 1
    const sameRows = areaRows === rows || (area.startRow === 0 && area.endRow === 1_048_575 && rows > 1)
    const sameColumns = areaColumns === columns || (area.startColumn === 0 && area.endColumn === 16_383 && columns > 1)
    if (!sameRows || !sameColumns || (rows > 1 && columns > 1)) continue
    const at = intersection(area, row, column)
    if (!at) return valueError()
    return at.row < rows && at.column < columns ? (array.get(at.row, at.column) ?? valueError()) : valueError()
  }
  return array.get(0, 0) ?? valueError()
}

/**
 * Applies the intersection on every Univer the editor creates
 * (patches/sheets-univer-hook.patch). `isArrayFormula` names the cells of legacy array
 * (CSE) formulas, which keep their array.
 */
export function installImplicitIntersection(isArrayFormula: (sheetId: string, row: number, column: number) => boolean): void {
  const hook = window as unknown as { teamGridOnUniver?: (univer: Univer) => void }
  const previous = hook.teamGridOnUniver
  hook.teamGridOnUniver = (univer) => {
    previous?.(univer)
    const injector = univer.__getInjector()
    const runtime = injector.get(IFormulaRuntimeService)
    const formulas = injector.get(FormulaDataModel)
    const setRuntimeData = runtime.setRuntimeData.bind(runtime)
    runtime.setRuntimeData = (variant) => {
      const value = variant as Variant
      const multiple =
        (value.isReferenceObject() || value.isArray()) &&
        (() => {
          const { startRow, endRow, startColumn, endColumn } = (value as unknown as Reference).getRangePosition?.() ?? {
            startRow: 0,
            endRow: (value as unknown as ArrayValueObject).getRowCount() - 1,
            startColumn: 0,
            endColumn: (value as unknown as ArrayValueObject).getColumnCount() - 1,
          }
          return startRow !== endRow || startColumn !== endColumn
        })()
      if (multiple) {
        const { currentRow: row, currentColumn: column, currentSubUnitId: sheetId, currentUnitId: unitId } = runtime
        const formula = formulas.getFormulaStringByCell(row, column, sheetId, unitId)
        if (formula && !spillsDynamicArray(formula) && !isArrayFormula(sheetId, row, column)) {
          return setRuntimeData(implicitIntersection(value, formula, row, column) as typeof variant)
        }
      }
      return setRuntimeData(variant)
    }
  }
}
