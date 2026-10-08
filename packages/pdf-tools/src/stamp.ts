import type { Font, Matrix } from 'mupdf'
import {
  deref,
  type Mupdf,
  type PDFDocument,
  type PDFObject,
  type ToolInput,
  type ToolOutput,
  savePdf,
  stem,
  withPdf,
} from './core'
import { parsePageSet } from './page-ranges'

export type Rgb = [number, number, number]

/** "#RRGGBB" -> [r, g, b] in 0-1; black for anything else. */
export function parseHexColor(hex: string): Rgb {
  const m = /^#?([0-9a-f]{6})$/i.exec(hex.trim())
  if (!m) return [0, 0, 0]
  const n = parseInt(m[1], 16)
  return [((n >> 16) & 255) / 255, ((n >> 8) & 255) / 255, (n & 255) / 255]
}

/**
 * Bytes for a simple Latin font: ASCII and Latin-1 map to themselves under
 * WinAnsiEncoding; anything else becomes "?" until Arabic text gets a shaped
 * font path (Phase 6).
 */
function encodeLatin(text: string): number[] {
  const out: number[] = []
  for (const ch of text) {
    const c = ch.codePointAt(0)!
    out.push((c >= 32 && c < 127) || (c >= 160 && c <= 255) ? c : 63)
  }
  return out
}

/** A PDF literal string for the bytes, with the three special characters escaped. */
function pdfString(bytes: readonly number[]): string {
  let s = '('
  for (const b of bytes) {
    if (b === 0x28 || b === 0x29 || b === 0x5c) s += '\\' + String.fromCharCode(b)
    else if (b < 32 || b > 126) s += '\\' + b.toString(8).padStart(3, '0')
    else s += String.fromCharCode(b)
  }
  return s + ')'
}

function textWidth(font: Font, bytes: readonly number[], size: number): number {
  let w = 0
  for (const b of bytes) w += font.advanceGlyph(font.encodeCharacter(b))
  return w * size
}

const num = (v: number) => (Math.abs(v) < 1e-6 ? '0' : v.toFixed(4).replace(/\.?0+$/, ''))

/** Put `value` under a fresh key in `dict`, returning the key used. */
function addUnique(doc: PDFDocument, res: PDFObject, kind: string, base: string, value: PDFObject) {
  let sub = res.get(kind)
  if (sub.isNull()) {
    sub = doc.newDictionary()
    res.put(kind, sub)
  }
  sub = deref(sub)
  let key = base
  for (let i = 1; !sub.get(key).isNull(); i++) key = `${base}${i}`
  sub.put(key, value)
  return key
}

interface OverlayContext {
  /** visible page size in points (rotation applied) */
  width: number
  height: number
  font: string
  gs: string
}

/**
 * Draw on top of a page. `build` writes content-stream operators in visual
 * coordinates: origin at the bottom-left of the page as displayed, y up,
 * whatever the page's /Rotate or crop box. The existing content is wrapped in
 * q/Q so its leftover graphics state cannot leak into the overlay.
 */
function overlayPage(
  m: Mupdf,
  doc: PDFDocument,
  index: number,
  fontRef: PDFObject,
  opacity: number,
  build: (ctx: OverlayContext) => string,
) {
  const page = doc.loadPage(index)
  try {
    const [x0, y0, x1, y1] = page.getBounds()
    const toPdf = m.Matrix.invert(page.getTransform())
    const visualToPdf: Matrix = m.Matrix.concat([1, 0, 0, -1, x0, y1], toPdf)
    const obj = page.getObject()
    let res = obj.getInheritable('Resources')
    if (res.isNull()) res = doc.newDictionary()
    // an inherited dictionary moves onto the page so the new keys cannot miss it
    obj.put('Resources', res)
    res = deref(res)
    const font = addUnique(doc, res, 'Font', 'HFTool', fontRef)
    const gs = addUnique(
      doc,
      res,
      'ExtGState',
      'HFToolGs',
      doc.addObject({ Type: doc.newName('ExtGState'), CA: opacity, ca: opacity }),
    )
    const ops = build({ width: x1 - x0, height: y1 - y0, font, gs })
    const before = doc.addStream('q\n', {})
    const after = doc.addStream(
      `Q\nq ${visualToPdf.map(num).join(' ')} cm /${gs} gs\n${ops}\nQ\n`,
      {},
    )
    const contents = obj.get('Contents')
    const arr = doc.newArray()
    arr.push(before)
    const resolved = deref(contents)
    if (resolved.isArray()) resolved.forEach((v) => arr.push(v))
    else if (!contents.isNull()) arr.push(contents)
    arr.push(after)
    obj.put('Contents', arr)
  } finally {
    page.destroy()
  }
}

function fontFor(m: Mupdf, doc: PDFDocument, bold: boolean) {
  const font = new m.Font(bold ? 'Helvetica-Bold' : 'Helvetica')
  return { font, ref: doc.addSimpleFont(font, 'Latin') }
}

export interface WatermarkOptions {
  text: string
  fontSize: number
  /** "#RRGGBB" */
  color: string
  /** 0-1 */
  opacity: number
  /** degrees counter-clockwise, or along the page diagonal */
  rotation: number | 'diagonal'
  position: 'center' | 'top' | 'bottom'
  bold?: boolean
  /** page-range expression; empty means every page */
  pages?: string
}

export function addWatermark(m: Mupdf, input: ToolInput, opts: WatermarkOptions): ToolOutput {
  return withPdf(m, input, (doc) => {
    const { font, ref } = fontFor(m, doc, opts.bold ?? true)
    const bytes = encodeLatin(opts.text)
    const size = Math.max(1, opts.fontSize)
    const w = textWidth(font, bytes, size)
    const [r, g, b] = parseHexColor(opts.color)
    const opacity = Math.min(1, Math.max(0, opts.opacity))
    for (const p of parsePageSet(opts.pages ?? '', doc.countPages())) {
      overlayPage(m, doc, p, ref, opacity, ({ width, height, font: f }) => {
        const angle =
          opts.rotation === 'diagonal' ? Math.atan2(height, width) : (opts.rotation * Math.PI) / 180
        const cx = width / 2
        const cy =
          opts.position === 'top'
            ? height - size * 1.5
            : opts.position === 'bottom'
              ? size * 1.5
              : height / 2
        const cos = Math.cos(angle)
        const sin = Math.sin(angle)
        return [
          `${num(r)} ${num(g)} ${num(b)} rg`,
          `BT /${f} ${num(size)} Tf`,
          `${num(cos)} ${num(sin)} ${num(-sin)} ${num(cos)} ${num(cx)} ${num(cy)} Tm`,
          `${num(-w / 2)} ${num(-size * 0.35)} Td ${pdfString(bytes)} Tj ET`,
        ].join('\n')
      })
    }
    font.destroy()
    return { name: `${stem(input.name)}-watermarked.pdf`, bytes: savePdf(doc) }
  })
}

export type NumberPosition =
  'bottom-center' | 'bottom-right' | 'bottom-left' | 'top-center' | 'top-right' | 'top-left'

export interface PageNumberOptions {
  /** "{n}", "Page {n} of {total}", "- {n} -" */
  format: string
  position: NumberPosition
  fontSize: number
  /** distance from the page edge in points */
  margin: number
  /** number printed on the first numbered page */
  firstNumber: number
  /** "#RRGGBB" */
  color?: string
  /** pages that get a number; empty means every page */
  pages?: string
}

export function addPageNumbers(m: Mupdf, input: ToolInput, opts: PageNumberOptions): ToolOutput {
  return withPdf(m, input, (doc) => {
    const { font, ref } = fontFor(m, doc, false)
    const selected = [...parsePageSet(opts.pages ?? '', doc.countPages())].sort((a, b) => a - b)
    const size = Math.max(1, opts.fontSize)
    const [r, g, b] = parseHexColor(opts.color ?? '#000000')
    const first = selected[0] ?? 0
    const last = selected[selected.length - 1] ?? 0
    const total = opts.firstNumber + (last - first)
    for (const p of selected) {
      const label = opts.format
        .replace(/\{n\}/g, String(opts.firstNumber + (p - first)))
        .replace(/\{total\}/g, String(total))
      const bytes = encodeLatin(label)
      const w = textWidth(font, bytes, size)
      overlayPage(m, doc, p, ref, 1, ({ width, height, font: f }) => {
        const [vert, horiz] = opts.position.split('-') as ['top' | 'bottom', string]
        const x =
          horiz === 'left'
            ? opts.margin
            : horiz === 'right'
              ? width - opts.margin - w
              : (width - w) / 2
        const y = vert === 'top' ? height - opts.margin - size * 0.75 : opts.margin
        return [
          `${num(r)} ${num(g)} ${num(b)} rg`,
          `BT /${f} ${num(size)} Tf ${num(x)} ${num(y)} Td ${pdfString(bytes)} Tj ET`,
        ].join('\n')
      })
    }
    font.destroy()
    return { name: `${stem(input.name)}-numbered.pdf`, bytes: savePdf(doc) }
  })
}
