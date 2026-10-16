/**
 * Excel date serials (1900 system): 1 = 1900-01-01, and 60 = 1900-02-29, the day Excel
 * keeps for Lotus compatibility although 1900 was no leap year.
 */

/** 9999-12-31, the last date Excel accepts. */
export const MAX_DATE_SERIAL = 2_958_465

const DAY_MS = 86_400_000
/** Serial 61 = 1900-03-01; from there serials count real days. */
const EPOCH_MS = Date.UTC(1899, 11, 30)

export function serialToDate(serial: number): { year: number; month: number; day: number } {
  const whole = Math.floor(serial)
  if (whole === 60) return { year: 1900, month: 2, day: 29 }
  if (whole === 0) return { year: 1900, month: 1, day: 0 }
  const date = new Date(EPOCH_MS + (whole < 60 ? whole + 1 : whole) * DAY_MS)
  return { year: date.getUTCFullYear(), month: date.getUTCMonth() + 1, day: date.getUTCDate() }
}

export function dateToSerial(year: number, month: number, day: number): number {
  if (year === 1900 && month === 2 && day === 29) return 60
  const serial = Math.round((Date.UTC(year, month - 1, day) - EPOCH_MS) / DAY_MS)
  return serial <= 60 ? serial - 1 : serial
}

/** Days in a month as Excel counts them (February 1900 has 29). */
export function daysInMonth(year: number, month: number): number {
  if (year === 1900 && month === 2) return 29
  return new Date(Date.UTC(year, month, 0)).getUTCDate()
}

/** EDATE: the same day `months` later, at most the month's last day; null when outside Excel's dates. */
export function addMonths(serial: number, months: number): number | null {
  if (serial < 0 || serial > MAX_DATE_SERIAL) return null
  const { year, month, day } = serialToDate(serial)
  const index = year * 12 + (month - 1) + months
  const targetYear = Math.floor(index / 12)
  const targetMonth = (index % 12) + 1
  if (targetYear < 1900 || targetYear > 9999) return null
  const result = dateToSerial(targetYear, targetMonth, Math.min(Math.max(day, 1), daysInMonth(targetYear, targetMonth)))
  return result < 1 || result > MAX_DATE_SERIAL ? null : result
}
