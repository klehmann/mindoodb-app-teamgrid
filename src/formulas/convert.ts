/**
 * CONVERT with Excel's units: Univer's table misses units and prefix rules (which units
 * take a metric or binary prefix, km2 = 10^6 m2) and has other factors for some. The
 * table and rules follow IronCalc (base/src/functions/engineering/convert.rs,
 * Apache-2.0/MIT), whose results match Excel's in its test workbook; units are looked up
 * deterministically (the unit itself first, then the longest prefix).
 */
import { type ArrayValueObject, BaseFunction, type BaseValueObject, ErrorType, ErrorValueObject, NumberValueObject } from '@univerjs/engine-formula'

const PREFIXES: Record<string, number> = {
  Y: 1e24, Z: 1e21, E: 1e18, P: 1e15, T: 1e12, G: 1e9, M: 1e6, k: 1e3, h: 100, da: 10, e: 10,
  d: 0.1, c: 0.01, m: 0.001, u: 1e-6, n: 1e-9, p: 1e-12, f: 1e-15, a: 1e-18, z: 1e-21, y: 1e-24,
  Yi: 2 ** 80, Zi: 2 ** 70, Ei: 2 ** 60, Pi: 2 ** 50, Ti: 2 ** 40, Gi: 2 ** 30, Mi: 2 ** 20, ki: 2 ** 10,
}

const KINDS: Record<string, Record<string, number>> = {
  weight: {
    g: 1, sg: 14593.9029372064, lbm: 453.59237, u: 1.660538782e-24, ozm: 28.349523125, grain: 0.06479891,
    cwt: 45359.237, shweight: 45359.237, uk_cwt: 50802.34544, lcwt: 50802.34544, stone: 6350.29318,
    ton: 907184.74, brton: 1016046.9088, LTON: 1016046.9088, uk_ton: 1016046.9088,
  },
  distance: {
    m: 1, mi: 1609.344, Nmi: 1852, in: 0.0254, ft: 0.3048, yd: 0.9144, ang: 1e-10, ell: 1.143,
    ly: 9.4607304725808e15, parsec: 3.08567758128155e16, pc: 3.08567758128155e16, Picapt: 0.000352777777777778,
    Pica: 0.000352777777777778, pica: 0.00423333333333333, survey_mi: 1609.34721869444,
  },
  time: { yr: 31557600, day: 86400, d: 86400, hr: 3600, mn: 60, min: 60, sec: 1, s: 1 },
  pressure: { Pa: 1, p: 1, atm: 101325, at: 101325, mmHg: 133.322, psi: 6894.75729316836, Torr: 133.322368421053 },
  force: { N: 1, dyn: 0.00001, dy: 0.00001, lbf: 4.4482216152605, pond: 0.00980665 },
  energy: {
    J: 1, e: 0.0000001, c: 4.184, cal: 4.1868, eV: 1.602176487e-19, ev: 1.602176487e-19, HPh: 2684519.53769617,
    hh: 2684519.53769617, Wh: 3600, wh: 3600, flb: 1.3558179483314, BTU: 1055.05585262, btu: 1055.05585262,
  },
  power: { HP: 745.69987158227, h: 745.69987158227, PS: 735.49875, W: 1, w: 1 },
  magnetism: { T: 1, ga: 0.0001 },
  volume: {
    tsp: 0.00000492892159375, tspm: 0.000005, tbs: 0.00001478676478125, oz: 0.0000295735295625,
    cup: 0.0002365882365, pt: 0.000473176473, us_pt: 0.000473176473, uk_pt: 0.00056826125, qt: 0.000946352946,
    uk_qt: 0.0011365225, gal: 0.003785411784, uk_gal: 0.00454609, l: 0.001, L: 0.001, lt: 0.001, ang3: 1e-30,
    'ang^3': 1e-30, barrel: 0.158987294928, bushel: 0.03523907016688, ft3: 0.028316846592, 'ft^3': 0.028316846592,
    in3: 0.000016387064, 'in^3': 0.000016387064, ly3: 8.46786664623715e47, 'ly^3': 8.46786664623715e47, m3: 1,
    'm^3': 1, mi3: 4168181825.44058, 'mi^3': 4168181825.44058, yd3: 0.764554857984, 'yd^3': 0.764554857984,
    Nmi3: 6352182208, 'Nmi^3': 6352182208, Picapt3: 4.39039566186557e-11, 'Picapt^3': 4.39039566186557e-11,
    Pica3: 4.39039566186557e-11, 'Pica^3': 4.39039566186557e-11, GRT: 2.8316846592, regton: 2.8316846592,
    MTON: 1.13267386368,
  },
  area: {
    uk_acre: 4046.8564224, us_acre: 4046.87260987425, ang2: 1e-20, 'ang^2': 1e-20, ar: 100, ft2: 0.09290304,
    'ft^2': 0.09290304, ha: 10000, in2: 0.00064516, 'in^2': 0.00064516, ly2: 8.95054210748189e31,
    'ly^2': 8.95054210748189e31, m2: 1, 'm^2': 1, Morgen: 2500, mi2: 2589988.110336, 'mi^2': 2589988.110336,
    Nmi2: 3429904, 'Nmi^2': 3429904, Picapt2: 0.000000124452160493827, Pica2: 0.000000124452160493827,
    'Pica^2': 0.000000124452160493827, 'Picapt^2': 0.000000124452160493827, yd2: 0.83612736, 'yd^2': 0.83612736,
  },
  information: { bit: 1, byte: 8 },
  speed: {
    admkn: 0.514773333333333, kn: 0.514444444444444, 'm/h': 0.000277777777777778, 'm/hr': 0.000277777777777778,
    'm/s': 1, 'm/sec': 1, mph: 0.44704,
  },
  temperature: { C: 1, cel: 1, F: 1, fah: 1, K: 1, kel: 1, Rank: 1, Reau: 1 },
}

/** Units that take a prefix (there is no kilo-Celsius). */
const PREFIXABLE = new Set([
  'Pa', 'p', 'atm', 'at', 'mmHg', 'g', 'u', 'm', 'ang', 'ly', 'parsec', 'pc', 'ang2', 'ang^2', 'ar', 'm2', 'm^2',
  'N', 'dyn', 'dy', 'pond', 'J', 'e', 'c', 'cal', 'eV', 'ev', 'Wh', 'wh', 'W', 'w', 'T', 'ga', 'uk_pt', 'l', 'L',
  'lt', 'ang3', 'ang^3', 'm3', 'm^3', 'bit', 'byte', 'm/h', 'm/hr', 'm/s', 'm/sec', 'mph', 'K', 'kel',
])
const CUBIC = new Set(['ang3', 'ang^3', 'm3', 'm^3'])

interface Unit {
  kind: string
  name: string
  factor: number
}

const UNITS = new Map<string, Unit>()
for (const [kind, units] of Object.entries(KINDS)) {
  for (const [name, factor] of Object.entries(units)) UNITS.set(name, { kind, name, factor })
}
const PREFIX_NAMES = Object.keys(PREFIXES).sort((a, b) => b.length - a.length)

export function lookupUnit(text: string): Unit | undefined {
  const exact = UNITS.get(text)
  if (exact) return exact
  for (const prefix of PREFIX_NAMES) {
    if (!text.startsWith(prefix)) continue
    const unit = UNITS.get(text.slice(prefix.length))
    if (!unit || !PREFIXABLE.has(unit.name)) continue
    const modifier = PREFIXES[prefix]!
    // 1 km2 is 10^6 m2, 1 km3 is 10^9 m3 (but 1 kar is 1000 ar)
    const scale = unit.kind === 'area' && unit.name !== 'ar' ? modifier ** 2 : unit.kind === 'volume' && CUBIC.has(unit.name) ? modifier ** 3 : modifier
    return { ...unit, factor: unit.factor * scale }
  }
  return undefined
}

function toKelvin(value: number, unit: string): number {
  if (unit === 'C' || unit === 'cel') return value + 273.15
  if (unit === 'Rank') return (5 * value) / 9
  if (unit === 'Reau') return (5 * value) / 4 + 273.15
  if (unit === 'F' || unit === 'fah') return (5 / 9) * (value + 459.67)
  return value
}

function fromKelvin(kelvin: number, unit: string): number {
  if (unit === 'C' || unit === 'cel') return kelvin - 273.15
  if (unit === 'Rank') return (9 * kelvin) / 5
  if (unit === 'Reau') return (4 * (kelvin - 273.15)) / 5
  if (unit === 'F' || unit === 'fah') return (9 * kelvin) / 5 - 459.67
  return kelvin
}

/** Number in `from` units as `to` units; undefined when a unit is unknown or they do not fit (#N/A). */
export function convertUnits(value: number, from: string, to: string): number | undefined {
  const source = lookupUnit(from)
  const target = lookupUnit(to)
  if (!source || !target || source.kind !== target.kind) return undefined
  if (source.kind === 'temperature') return fromKelvin(toKelvin(value * source.factor, source.name), target.name) / target.factor
  return (value * source.factor) / target.factor
}

function single(value: BaseValueObject): BaseValueObject | undefined {
  if (!value.isArray()) return value
  const array = value as ArrayValueObject
  return array.getRowCount() === 1 && array.getColumnCount() === 1 ? (array.get(0, 0) ?? undefined) : undefined
}

export function createConvert(base: (name: string) => new (name: string) => BaseFunction): BaseFunction {
  const Base = base('CONVERT')
  return new (class extends Base {
    override calculate(number: BaseValueObject, fromUnit: BaseValueObject, toUnit: BaseValueObject) {
      const n = single(number)
      const from = single(fromUnit)
      const to = single(toUnit)
      // Arrays (spilled conversions) stay with Univer.
      if (!n || !from || !to) return super.calculate(number, fromUnit, toUnit)
      for (const value of [n, from, to]) if (value.isError()) return value
      if (n.isBoolean()) return ErrorValueObject.create(ErrorType.VALUE)
      const numeric = n.isNull() ? 0 : Number(n.getValue())
      if (n.isString() && (String(n.getValue()).trim() === '' || !Number.isFinite(numeric))) return ErrorValueObject.create(ErrorType.VALUE)
      const result = convertUnits(numeric, String(from.getValue() ?? ''), String(to.getValue() ?? ''))
      return result === undefined ? ErrorValueObject.create(ErrorType.NA) : NumberValueObject.create(result)
    }
  })('CONVERT')
}
