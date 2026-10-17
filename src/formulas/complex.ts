/**
 * Excel's complex-number functions (COMPLEX, IM*). Univer's parser rejects numbers in
 * exponent notation, which its own results contain ("1.2e-45-3i"), so every IM*
 * function over such a value failed with #NUM!; it also writes "1-1i" where Excel writes
 * "1-i". Parsing, Excel's way of writing the parts and the formulas follow IronCalc
 * (base/src/functions/engineering/complex.rs, Apache-2.0/MIT), whose results match
 * Excel's in its test workbook. IMSUM and IMPRODUCT take any number of arguments, as in
 * Excel.
 */
import {
  type ArrayValueObject,
  BaseFunction,
  type BaseValueObject,
  ErrorType,
  ErrorValueObject,
  NumberValueObject,
  StringValueObject,
} from '@univerjs/engine-formula'

type Suffix = 'i' | 'j'
export interface Complex {
  x: number
  y: number
  suffix: Suffix
}

const NUMBER = String.raw`(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?`
const FORMS: Array<[RegExp, (m: RegExpExecArray, sign: number) => Complex]> = [
  // 7.5
  [new RegExp(`^(${NUMBER})$`), (m, sign) => ({ x: sign * Number(m[1]), y: 0, suffix: 'i' })],
  // 2i
  [new RegExp(`^(${NUMBER})([ij])$`), (m, sign) => ({ x: 0, y: sign * Number(m[1]), suffix: m[2] as Suffix })],
  // 7-i
  [new RegExp(`^(${NUMBER})([+-])([ij])$`), (m, sign) => ({ x: sign * Number(m[1]), y: m[2] === '-' ? -1 : 1, suffix: m[3] as Suffix })],
  // 7-5i
  [
    new RegExp(`^(${NUMBER})([+-])(${NUMBER})([ij])$`),
    (m, sign) => ({ x: sign * Number(m[1]), y: (m[2] === '-' ? -1 : 1) * Number(m[3]), suffix: m[4] as Suffix }),
  ],
]

/** A complex number as Excel writes it ("1+2i", "-i", "1.2E-45-3j", "5"); null when it is none. */
export function parseComplex(text: string): Complex | null {
  const sign = text.startsWith('-') ? -1 : 1
  const rest = sign < 0 ? text.slice(1) : text
  if (rest === 'i' || rest === 'j') return { x: 0, y: sign, suffix: rest }
  for (const [pattern, build] of FORMS) {
    const match = pattern.exec(rest)
    if (match) return build(match, sign)
  }
  return null
}

/** A number in plain decimal notation (1e-18 → "0.000000000000000001"), shortest digits. */
function plain(value: number): string {
  const text = String(value)
  const match = /^(-?)(\d)(?:\.(\d+))?e([+-]\d+)$/.exec(text)
  if (!match) return text
  const [, sign, lead, fraction = '', exponentText] = match
  const digits = lead! + fraction
  const exponent = Number(exponentText)
  if (exponent < 0) return `${sign}0.${'0'.repeat(-exponent - 1)}${digits}`
  return sign + digits + '0'.repeat(Math.max(0, exponent - fraction.length))
}

/** One part: 15 significant digits; Excel writes the very small ones (up to 9e-20) in exponent notation. */
function part(value: number): string {
  if (value !== 0 && Math.abs(value) <= 9e-20) return value.toExponential().replace('e+', 'E').replace('e', 'E')
  return plain(value)
}

export function formatComplex({ x: rawX, y: rawY, suffix }: Complex): string {
  const x = Number(rawX.toPrecision(15))
  const y = Number(rawY.toPrecision(15))
  const yText = y === 1 ? '' : y === -1 ? '-' : part(y)
  if (x === 0 && y === 0) return '0'
  if (y === 0) return part(x)
  if (x === 0) return `${yText}${suffix}`
  return `${part(x)}${y > 0 ? '+' : ''}${yText}${suffix}`
}

class ArgumentError {
  constructor(readonly result: BaseValueObject) {}
}
const error = (type: ErrorType) => new ArgumentError(ErrorValueObject.create(type))

/** One value from an argument: a single cell's value, a 1×1 array's element. */
function single(value: BaseValueObject | undefined): BaseValueObject | undefined {
  let current = value
  if (current?.isReferenceObject()) current = (current as unknown as { toArrayValueObject(): ArrayValueObject }).toArrayValueObject()
  if (current?.isArray()) {
    const array = current as ArrayValueObject
    if (array.getRowCount() !== 1 || array.getColumnCount() !== 1) throw new ArgumentError(ErrorValueObject.create(ErrorType.VALUE))
    return array.get(0, 0) ?? undefined
  }
  return current
}

/** The text of an inumber argument (numbers as Excel writes them). */
function textOf(value: BaseValueObject | undefined): string {
  const one = single(value)
  if (!one || one.isNull()) return ''
  if (one.isError()) throw new ArgumentError(one)
  if (one.isBoolean()) throw error(ErrorType.VALUE)
  if (one.isNumber()) return part(Number(one.getValue()))
  return String(one.getValue())
}

/** An inumber argument; empty text is 0 unless `strict` (IMCOS & co. in Excel answer #NUM!). */
function complexOf(value: BaseValueObject | undefined, strict = false): Complex {
  const text = textOf(value)
  if (text === '' && !strict) return { x: 0, y: 0, suffix: 'i' }
  const parsed = parseComplex(text)
  if (!parsed) throw error(ErrorType.NUM)
  return parsed
}

/** A number argument (COMPLEX's parts, IMPOWER's power): numbers and numeric text. */
function numberOf(value: BaseValueObject | undefined): number {
  const one = single(value)
  if (!one || one.isNull()) return 0
  if (one.isError()) throw new ArgumentError(one)
  if (one.isNumber()) return Number(one.getValue())
  if (one.isString()) {
    const text = String(one.getValue()).trim()
    if (text !== '' && Number.isFinite(Number(text))) return Number(text)
  }
  throw error(ErrorType.VALUE)
}

const result = (value: Complex): BaseValueObject => {
  if (!Number.isFinite(value.x) || !Number.isFinite(value.y)) return ErrorValueObject.create(ErrorType.NUM)
  return StringValueObject.create(formatComplex(value))
}
const real = (value: number): BaseValueObject =>
  Number.isFinite(value) ? NumberValueObject.create(value) : ErrorValueObject.create(ErrorType.NUM)

function sameSuffix(values: Complex[]): Suffix {
  // Numbers without an imaginary part carry the default "i" and mix with "j".
  const suffixes = new Set(values.filter((value) => value.y !== 0).map((value) => value.suffix))
  if (suffixes.size > 1) throw error(ErrorType.VALUE)
  return suffixes.values().next().value ?? values[0]?.suffix ?? 'i'
}

type Unary = (z: Complex) => BaseValueObject

const UNARY: Record<string, { strict?: boolean; compute: Unary }> = {
  IMABS: { compute: ({ x, y }) => real(Math.hypot(x, y)) },
  IMAGINARY: { compute: ({ y }) => real(y) },
  IMREAL: { compute: ({ x }) => real(x) },
  IMARGUMENT: { compute: ({ x, y }) => (x === 0 && y === 0 ? ErrorValueObject.create(ErrorType.DIV_BY_ZERO) : real(Math.atan2(y, x))) },
  IMCONJUGATE: { compute: (z) => result({ ...z, y: -z.y }) },
  IMCOS: { strict: true, compute: ({ x, y, suffix }) => result({ x: Math.cos(x) * Math.cosh(y), y: -Math.sin(x) * Math.sinh(y), suffix }) },
  IMCOSH: { strict: true, compute: ({ x, y, suffix }) => result({ x: Math.cosh(x) * Math.cos(y), y: Math.sinh(x) * Math.sin(y), suffix }) },
  IMCOT: {
    strict: true,
    compute: ({ x, y, suffix }) => {
      if (x === 0 && y !== 0) return result({ x: 0, y: -1 / Math.tanh(y), suffix })
      if (y === 0) return result({ x: 1 / Math.tan(x), y: 0, suffix })
      const cot = 1 / Math.tan(x)
      const coth = 1 / Math.tanh(y)
      const t = cot * cot + coth * coth
      return result({ x: (cot * coth * coth - cot) / t, y: (-cot * cot * coth - coth) / t, suffix })
    },
  },
  IMCSC: {
    strict: true,
    compute: ({ x, y, suffix }) => {
      const t = Math.sin(x) ** 2 * Math.cosh(y) ** 2 + Math.cos(x) ** 2 * Math.sinh(y) ** 2
      return result({ x: (Math.sin(x) * Math.cosh(y)) / t, y: (-Math.cos(x) * Math.sinh(y)) / t, suffix })
    },
  },
  IMCSCH: {
    strict: true,
    compute: ({ x, y, suffix }) => {
      const t = Math.sinh(x) ** 2 * Math.cos(y) ** 2 + Math.cosh(x) ** 2 * Math.sin(y) ** 2
      return result({ x: (Math.sinh(x) * Math.cos(y)) / t, y: (-Math.cosh(x) * Math.sin(y)) / t, suffix })
    },
  },
  IMEXP: { compute: ({ x, y, suffix }) => result({ x: Math.exp(x) * Math.cos(y), y: Math.exp(x) * Math.sin(y), suffix }) },
  IMLN: { compute: ({ x, y, suffix }) => result({ x: Math.log(Math.hypot(x, y)), y: Math.atan2(y, x), suffix }) },
  IMLOG10: { compute: ({ x, y, suffix }) => result({ x: Math.log10(Math.hypot(x, y)), y: Math.atan2(y, x) * Math.LOG10E, suffix }) },
  IMLOG2: { compute: ({ x, y, suffix }) => result({ x: Math.log2(Math.hypot(x, y)), y: Math.atan2(y, x) * Math.LOG2E, suffix }) },
  IMSEC: {
    compute: ({ x, y, suffix }) => {
      const t = Math.cos(x) ** 2 * Math.cosh(y) ** 2 + Math.sin(x) ** 2 * Math.sinh(y) ** 2
      return result({ x: (Math.cos(x) * Math.cosh(y)) / t, y: (Math.sin(x) * Math.sinh(y)) / t, suffix })
    },
  },
  IMSECH: {
    compute: ({ x, y, suffix }) => {
      const t = Math.cosh(x) ** 2 * Math.cos(y) ** 2 + Math.sinh(x) ** 2 * Math.sin(y) ** 2
      return result({ x: (Math.cosh(x) * Math.cos(y)) / t, y: (-Math.sinh(x) * Math.sin(y)) / t, suffix })
    },
  },
  IMSIN: { compute: ({ x, y, suffix }) => result({ x: Math.sin(x) * Math.cosh(y), y: Math.cos(x) * Math.sinh(y), suffix }) },
  IMSINH: { compute: ({ x, y, suffix }) => result({ x: Math.sinh(x) * Math.cos(y), y: Math.cosh(x) * Math.sin(y), suffix }) },
  IMSQRT: {
    compute: ({ x, y, suffix }) => {
      const r = Math.sqrt(Math.hypot(x, y))
      const a = Math.atan2(y, x) / 2
      return result({ x: r * Math.cos(a), y: r * Math.sin(a), suffix })
    },
  },
  IMTAN: {
    compute: ({ x, y, suffix }) => {
      const tan = Math.tan(x)
      const tanh = Math.tanh(y)
      const t = 1 + tan * tan * tanh * tanh
      return result({ x: (tan - tan * tanh * tanh) / t, y: (tanh + tan * tan * tanh) / t, suffix })
    },
  },
}

const OTHER: Record<string, (args: BaseValueObject[]) => BaseValueObject> = {
  COMPLEX: ([realNum, iNum, suffix]) => {
    const x = numberOf(realNum)
    const y = numberOf(iNum)
    const text = suffix === undefined ? 'i' : textOf(suffix)
    if (text !== '' && text !== 'i' && text !== 'j') return ErrorValueObject.create(ErrorType.VALUE)
    return result({ x, y, suffix: text === 'j' ? 'j' : 'i' })
  },
  IMDIV: ([a, b]) => {
    const z1 = complexOf(a)
    const z2 = complexOf(b)
    const suffix = sameSuffix([z1, z2])
    const t = z2.x * z2.x + z2.y * z2.y
    if (t === 0) return ErrorValueObject.create(ErrorType.NUM)
    return result({ x: (z1.x * z2.x + z1.y * z2.y) / t, y: (-z1.x * z2.y + z1.y * z2.x) / t, suffix })
  },
  IMSUB: ([a, b]) => {
    const z1 = complexOf(a)
    const z2 = complexOf(b)
    return result({ x: z1.x - z2.x, y: z1.y - z2.y, suffix: sameSuffix([z1, z2]) })
  },
  IMPOWER: ([z, power]) => {
    const { x, y, suffix } = complexOf(z)
    const one = single(power)
    if (one?.isBoolean()) return ErrorValueObject.create(ErrorType.VALUE)
    const n = numberOf(power)
    const r = Math.hypot(x, y)
    const a = Math.atan2(y, x)
    return result({ x: r ** n * Math.cos(a * n), y: r ** n * Math.sin(a * n), suffix })
  },
  IMSUM: (args) => {
    const values = inumbers(args)
    return result({ x: values.reduce((sum, z) => sum + z.x, 0), y: values.reduce((sum, z) => sum + z.y, 0), suffix: sameSuffix(values) })
  },
  IMPRODUCT: (args) => {
    const values = inumbers(args)
    const product = values.reduce((p, z) => ({ x: p.x * z.x - p.y * z.y, y: p.x * z.y + p.y * z.x }), { x: 1, y: 0 })
    return result({ ...product, suffix: sameSuffix(values) })
  },
}

/** IMSUM's and IMPRODUCT's arguments: an empty cell given alone is 0 (as in Excel), empty cells of a range do not count. */
function inumbers(args: BaseValueObject[]): Complex[] {
  const values: Complex[] = []
  for (const arg of args) {
    const array = arg.isReferenceObject() ? (arg as unknown as { toArrayValueObject(): ArrayValueObject }).toArrayValueObject() : arg
    if (!array.isArray() || ((array as ArrayValueObject).getRowCount() === 1 && (array as ArrayValueObject).getColumnCount() === 1)) {
      values.push(complexOf(array))
      continue
    }
    ;(array as ArrayValueObject).iterator((cell) => {
      if (cell && !cell.isNull()) values.push(complexOf(cell))
    })
  }
  return values
}

/** Executors for these functions, subclassing Univer's (`base` gives its class by name). */
export function createComplexExecutors(base: (name: string) => new (name: string) => BaseFunction): BaseFunction[] {
  const wrap = (name: string, compute: (args: BaseValueObject[]) => BaseValueObject): BaseFunction => {
    const Base = base(name)
    const executor = new (class extends Base {
      override calculate(...args: BaseValueObject[]) {
        try {
          return compute(args)
        } catch (failure) {
          if (failure instanceof ArgumentError) return failure.result
          throw failure
        }
      }
    })(name)
    return executor
  }
  return [
    ...Object.entries(UNARY).map(([name, { strict, compute }]) => wrap(name, ([z]) => compute(complexOf(z, strict)))),
    ...Object.entries(OTHER).map(([name, compute]) => wrap(name, compute)),
  ]
}
