/**
 * ACCRINT as Excel computes it. Univer's results for several frequency/basis
 * combinations were off, and depended on what the engine had calculated before (the
 * same workbook gave other results when opened after another one). The algorithm
 * follows IronCalc (base/src/functions/financial/accrint.rs, Apache-2.0/MIT), itself a
 * port of the ExcelFinancialFunctions library; its results match Excel's in IronCalc's
 * test workbooks.
 */
import { serialToDate } from './dates'

interface Day {
  year: number
  month: number
  day: number
}

const lastDay = (year: number, month: number) => new Date(Date.UTC(year, month, 0)).getUTCDate()
const isLastDay = (d: Day) => d.day === lastDay(d.year, d.month)
const isLastFeb = (d: Day) => d.month === 2 && isLastDay(d)
const time = (d: Day) => Date.UTC(d.year, d.month - 1, d.day)
const compare = (a: Day, b: Day) => time(a) - time(b)
const days = (after: Day, before: Day) => Math.round((time(after) - time(before)) / 86_400_000)

/** .NET AddMonths: the day clamped to the target month's last day. */
function addMonths(d: Day, n: number): Day {
  const total = d.year * 12 + (d.month - 1) + n
  const year = Math.floor(total / 12)
  const month = (((total % 12) + 12) % 12) + 1
  return { year, month, day: Math.min(d.day, lastDay(year, month)) }
}

function changeMonth(d: Day, n: number, toLastDay: boolean): Day {
  const moved = addMonths(d, n)
  return toLastDay ? { ...moved, day: lastDay(moved.year, moved.month) } : moved
}

const diff360 = (sd: number, sm: number, sy: number, ed: number, em: number, ey: number) => (ey - sy) * 360 + (em - sm) * 30 + (ed - sd)

/** US (NASD) 30/360; `modifyBoth` is Excel's ModifyBothDates variant. */
function diff360Us(start: Day, end: Day, modifyBoth: boolean): number {
  let sd = start.day
  let ed = end.day
  if (isLastFeb(end) && (isLastFeb(start) || modifyBoth)) ed = 30
  if (ed === 31 && (sd >= 30 || modifyBoth)) ed = 30
  if (sd === 31) sd = 30
  if (isLastFeb(start)) sd = 30
  return diff360(sd, start.month, start.year, ed, end.month, end.year)
}

function diff360Eu(start: Day, end: Day): number {
  return diff360(start.day === 31 ? 30 : start.day, start.month, start.year, end.day === 31 ? 30 : end.day, end.month, end.year)
}

function diff365(start: Day, end: Day): number {
  const s = { ...start, day: start.month === 2 && start.day > 28 ? 28 : start.day }
  const e = { ...end, day: end.month === 2 && end.day > 28 ? 28 : end.day }
  return (end.year - start.year) * 365 + days(e, s)
}

/** Coupon boundaries stepped from `start` by `months` until `end` is passed: [front, trailing]. */
function findPcdNcd(start: Day, end: Day, months: number, toLastDay: boolean): [Day, Day] {
  let front = start
  let trailing = end
  for (;;) {
    if (months > 0 ? compare(front, end) >= 0 : compare(front, end) <= 0) return [front, trailing]
    trailing = front
    front = changeMonth(front, months, toLastDay)
  }
}

function daysBetween(start: Day, end: Day, numerator: boolean, basis: number): number {
  if (basis === 0) return diff360Us(start, end, false)
  if (basis === 4) return diff360Eu(start, end)
  if (basis === 2) return numerator ? days(end, start) : diff360Us(start, end, false)
  if (basis === 3) return numerator ? days(end, start) : diff365(start, end)
  return days(end, start)
}

function coupDays(settlement: Day, maturity: Day, frequency: number, basis: number): number {
  if (basis === 0 || basis === 2 || basis === 4) return 360 / frequency
  if (basis === 3) return 365 / frequency
  const [pcd, ncd] = findPcdNcd(maturity, settlement, -(12 / frequency), isLastDay(maturity))
  return days(ncd, pcd)
}

export interface AccrintInput {
  issue: number
  firstInterest: number
  settlement: number
  rate: number
  par: number
  frequency: number
  basis: number
  calcMethod: boolean
}

/** The accrued interest, or the Excel error for invalid arguments. */
export function accrint({ issue: issueSerial, firstInterest: firstSerial, settlement: settlementSerial, rate, par, frequency, basis, calcMethod }: AccrintInput): number | '#NUM!' {
  const issueWhole = Math.floor(issueSerial)
  const settlementWhole = Math.floor(settlementSerial)
  if (rate <= 0 || par <= 0 || ![1, 2, 4].includes(frequency) || basis < 0 || basis > 4 || issueWhole >= settlementWhole) return '#NUM!'
  if (issueWhole < 1 || Math.floor(firstSerial) < 1) return '#NUM!'
  const issue = serialToDate(issueWhole)
  const firstInterest = serialToDate(Math.floor(firstSerial))
  const settlement = serialToDate(settlementWhole)
  const months = 12 / frequency
  const endMonth = isLastDay(firstInterest)
  const calc = calcMethod ? 1 : 0

  const regularPcd = changeMonth(firstInterest, -months, endMonth)
  const pcd = compare(settlement, firstInterest) > 0 && calcMethod ? findPcdNcd(firstInterest, settlement, months, endMonth)[1] : regularPcd
  const firstDate = compare(issue, pcd) > 0 ? issue : pcd
  let accrued = daysBetween(firstDate, settlement, true, basis) / coupDays(regularPcd, firstInterest, frequency, basis)

  // Coupon periods backward from pcd to the issue.
  let front = pcd
  while (compare(front, issue) > 0) {
    const periodEnd = front
    front = changeMonth(front, -months, endMonth)
    const periodStart = front
    if (compare(issue, periodStart) <= 0) {
      accrued += calc
    } else {
      const numerator = basis === 0 ? diff360Us(issue, periodEnd, false) : daysBetween(issue, periodEnd, true, basis)
      const denominator = basis === 0 ? diff360Us(periodStart, periodEnd, true) : basis === 3 ? 365 / frequency : daysBetween(periodStart, periodEnd, false, basis)
      accrued += numerator / denominator
    }
  }
  return (par * rate * accrued) / frequency
}
