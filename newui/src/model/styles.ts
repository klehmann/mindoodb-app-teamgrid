// Content-addressed style registry: a style's id is a hash of its canonical
// JSON, so equal styles written by different people share one entry.
import type { WorkbookStyleEdit } from '@genoffice/xlsx-gateway/shared/edit-schemas'

import type { BorderEdge, CellStyle, StoredStyle, StyleId } from './schema'

/** JSON with object keys sorted, so equal content compares equal whatever the key order. */
export function canonicalJson(value: unknown): string {
  if (value === null || typeof value !== 'object') return JSON.stringify(value)
  if (Array.isArray(value)) return `[${value.map(canonicalJson).join(',')}]`
  const entries = Object.entries(value as Record<string, unknown>)
    .filter(([, entry]) => entry !== undefined)
    .sort(([left], [right]) => (left < right ? -1 : left > right ? 1 : 0))
  return `{${entries.map(([key, entry]) => `${JSON.stringify(key)}:${canonicalJson(entry)}`).join(',')}}`
}

/** cyrb53: a fast 53-bit string hash; collisions are irrelevant at style-table sizes. */
function hash53(text: string): string {
  let h1 = 0xdeadbeef
  let h2 = 0x41c6ce57
  for (let index = 0; index < text.length; index += 1) {
    const code = text.charCodeAt(index)
    h1 = Math.imul(h1 ^ code, 2654435761)
    h2 = Math.imul(h2 ^ code, 1597334677)
  }
  h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909)
  h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909)
  return (4294967296 * (2097151 & h2) + (h1 >>> 0)).toString(16).padStart(14, '0')
}

export function styleIdOf(style: StoredStyle): StyleId {
  return `st_${hash53(canonicalJson(style))}`
}

/** Font of the exported workbook's default format (the gateway's minimal stylesheet). */
const DEFAULT_FONT_FAMILY = 'Calibri'
const DEFAULT_FONT_SIZE = 11

/** Stored ids for a file's style table by xf index; undefined = default format. */
export function styleTable(styles: readonly CellStyle[]): { ids: (StyleId | undefined)[]; byId: Record<StyleId, StoredStyle> } {
  const byId: Record<StyleId, StoredStyle> = {}
  const ids = styles.map((style) => {
    const stored = toStyleEdit(style)
    if (!stored) return undefined
    const id = styleIdOf(stored)
    byId[id] = stored
    return id
  })
  return { ids, byId }
}

function hexColor(color: string | undefined): string | undefined {
  if (!color) return undefined
  const hex = color.replace(/^#/, '')
  if (/^[0-9A-Fa-f]{8}$/.test(hex)) return `#${hex.slice(2).toUpperCase()}`
  if (/^[0-9A-Fa-f]{6}$/.test(hex)) return `#${hex.toUpperCase()}`
  return undefined
}

function styleColor(color: string | undefined, theme: number | undefined, tint: number | undefined) {
  if (theme !== undefined && theme <= 11) return tint === undefined ? { theme } : { theme, tint }
  return hexColor(color)
}

const BORDER_STYLES = new Set([
  'thin', 'medium', 'thick', 'dashed', 'dotted', 'double', 'hair', 'dashDot', 'dashDotDot',
  'mediumDashed', 'mediumDashDot', 'mediumDashDotDot', 'slantDashDot',
])

function borderEdit(edge: BorderEdge | undefined) {
  if (!edge || !BORDER_STYLES.has(edge.style)) return undefined
  const color = hexColor(edge.color)
  return color ? { style: edge.style, color } : { style: edge.style }
}

const HORIZONTAL = new Set(['left', 'center', 'right', 'justify', 'distributed'])
const VERTICAL: Record<string, 'top' | 'center' | 'bottom'> = { top: 'top', center: 'center', middle: 'center', bottom: 'bottom' }

/** A stored style as a delta on the blank workbook's default style. */
export function toStyleEdit(style: CellStyle): WorkbookStyleEdit | undefined {
  const edit: Record<string, unknown> = {}
  if (style.bold) edit.bold = true
  if (style.italic) edit.italic = true
  if (style.underline) edit.underline = true
  if (style.strikethrough) edit.strikethrough = true
  if (style.wrapText) edit.wrapText = true
  // The default font is left out: a style without a font reads back with it,
  // and both must map to the same stored style.
  if (style.fontFamily && style.fontFamily !== DEFAULT_FONT_FAMILY) edit.fontFamily = style.fontFamily
  if (style.fontSize && style.fontSize !== DEFAULT_FONT_SIZE) edit.fontSize = style.fontSize
  const fontColor = styleColor(style.fontColor, style.fontColorTheme, style.fontColorTint)
  if (fontColor) edit.fontColor = fontColor
  const fillColor = styleColor(style.fillColor, style.fillColorTheme, style.fillColorTint)
  if (fillColor) edit.fillColor = fillColor
  if (style.horizontalAlignment && HORIZONTAL.has(style.horizontalAlignment)) {
    edit.horizontalAlignment = style.horizontalAlignment
  }
  const vertical = style.verticalAlignment ? VERTICAL[style.verticalAlignment] : undefined
  if (vertical) edit.verticalAlignment = vertical
  if (style.textRotation !== undefined && (style.textRotation <= 180 || style.textRotation === 255)) {
    edit.textRotation = style.textRotation
  }
  if (style.indent) edit.indent = Math.min(style.indent, 250)
  if (style.numberFormat && style.numberFormat !== 'General') edit.numberFormat = style.numberFormat
  for (const side of ['Top', 'Bottom', 'Left', 'Right'] as const) {
    const border = borderEdit(style[`border${side}`])
    if (border) edit[`border${side}`] = border
  }
  return Object.keys(edit).length > 0 ? (edit as WorkbookStyleEdit) : undefined
}
