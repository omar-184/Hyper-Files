import {
  type Mupdf,
  type PDFDocument,
  type PDFPage,
  PdfToolError,
  type ToolInput,
  type ToolOutput,
  savePdf,
  stem,
  withPdf,
} from './core'
import { parsePageSet } from './page-ranges'
import { encodeLatin, fontFor, num, overlayPage, pdfString, textWidth } from './stamp'

/** Normalized box: 0-1 of the rendered page, origin bottom-left, y up. */
export interface OcrBox {
  x0: number
  y0: number
  x1: number
  y1: number
}

/** One recognized line, as the system OCR helpers report it (packages/pdf2docx ocr-vision). */
export interface OcrLine {
  text: string
  /** 0-1 */
  confidence: number
  box: OcrBox
  /** per-character boxes; characters of one word share its box, spaces have empty boxes */
  chars?: { text: string; box: OcrBox }[]
}

/**
 * Recognize one page render (PNG, display orientation). Null when recognition
 * failed for this page; the page is then left as it was.
 */
export type OcrRecognizer = (
  png: Uint8Array,
  page: { widthPt: number; heightPt: number },
) => { lines: OcrLine[] } | null

export interface OcrOptions {
  /** page-range expression; empty means every page */
  pages?: string
}

export interface OcrStats {
  /** pages with no text, the ones recognition ran on */
  scanned: number
  /** scanned pages that got a text layer */
  recognized: number
}

/** A page with fewer non-space characters than this counts as a scan (same rule as the viewer). */
const MIN_PAGE_CHARS = 8
/** Lines the engine is this unsure about are stray marks, not text. */
const MIN_LINE_CONFIDENCE = 0.3
/** Long edge of the page render handed to the engine, in pixels. */
const RENDER_LONG_EDGE = 2048

interface Word {
  text: string
  box: OcrBox
}

const emptyBox = (b: OcrBox) => b.x1 - b.x0 <= 0 || b.y1 - b.y0 <= 0

/** Split a line into words from its character boxes, or keep it whole without them. */
function lineWords(line: OcrLine): Word[] {
  const whole = () => {
    const text = line.text.trim()
    return text && !emptyBox(line.box) ? [{ text, box: line.box }] : []
  }
  if (!line.chars || line.chars.length === 0) return whole()
  const words: Word[] = []
  let cur: Word | null = null
  for (const c of line.chars) {
    if (c.text.trim() === '' || emptyBox(c.box)) {
      if (cur) words.push(cur)
      cur = null
    } else if (cur && c.box.x0 === cur.box.x0 && c.box.x1 === cur.box.x1) {
      cur.text += c.text
    } else {
      if (cur) words.push(cur)
      cur = { text: c.text, box: { ...c.box } }
    }
  }
  if (cur) words.push(cur)
  // some engines box only part of the line: then the line box is the better guess
  const covered = words.reduce((n, w) => n + w.text.length, 0)
  return covered < line.text.replace(/\s/g, '').length * 0.5 ? whole() : words
}

/** True when the page has (almost) no extractable text. */
function isScanned(page: PDFPage) {
  const st = page.toStructuredText('')
  try {
    return st.asText().replace(/\s/g, '').length < MIN_PAGE_CHARS
  } finally {
    st.destroy()
  }
}

/**
 * Make scanned pages searchable: every page without text is rendered, handed
 * to the system OCR engine, and gets the recognized words as invisible text
 * (render mode 3) laid over the picture, so search, select and copy work in
 * any PDF reader while the page looks exactly the same. Pages that already
 * have text are left alone. Latin text only for now: words in other scripts
 * are left out rather than written as question marks.
 */
export function ocrPdf(
  m: Mupdf,
  input: ToolInput,
  recognize: OcrRecognizer,
  opts: OcrOptions = {},
  onPage: (done: number, total: number) => void = () => {},
): ToolOutput & { stats: OcrStats } {
  return withPdf(m, input, (doc) => {
    const selected = [...parsePageSet(opts.pages ?? '', doc.countPages())].sort((a, b) => a - b)
    const scanned = selected.filter((p) => {
      const page = doc.loadPage(p)
      try {
        return isScanned(page)
      } finally {
        page.destroy()
      }
    })
    if (scanned.length === 0) {
      throw new PdfToolError(
        'bad-input',
        `${input.name} already has text on every page, so there is nothing to recognize`,
        input.name,
      )
    }
    const { font, ref } = fontFor(m, doc, false)
    let recognized = 0
    try {
      scanned.forEach((p, i) => {
        onPage(i, scanned.length)
        const words = recognizePage(m, doc, p, recognize)
        if (words.length === 0) return
        overlayPage(m, doc, p, ref, 1, ({ width, height, font: f }) => {
          const ops = ['3 Tr']
          for (const w of words) {
            const bytes = encodeLatin(w.text)
            const x = w.box.x0 * width
            const y = w.box.y0 * height
            const h = (w.box.y1 - w.box.y0) * height
            const boxW = (w.box.x1 - w.box.x0) * width
            const size = Math.max(1, h)
            const natural = textWidth(font, bytes, size)
            const scale = natural > 0 ? (boxW / natural) * 100 : 100
            // the trailing space keeps words apart when a reader copies the line
            const text = pdfString(w.last ? bytes : [...bytes, 32])
            ops.push(
              `BT /${f} ${num(size)} Tf ${num(scale)} Tz ${num(x)} ${num(y + h * 0.2)} Td ${text} Tj ET`,
            )
          }
          return ops.join('\n')
        })
        recognized += 1
      })
      onPage(scanned.length, scanned.length)
    } finally {
      font.destroy()
    }
    if (recognized === 0) {
      throw new PdfToolError(
        'unsupported',
        `no text could be recognized on the scanned pages of ${input.name}`,
        input.name,
      )
    }
    return {
      name: `${stem(input.name)}-searchable.pdf`,
      bytes: savePdf(doc),
      stats: { scanned: scanned.length, recognized },
    }
  })
}

/** Render one page and recognize it; the Latin words that can be written, in reading order. */
function recognizePage(
  m: Mupdf,
  doc: PDFDocument,
  index: number,
  recognize: OcrRecognizer,
): (Word & { last: boolean })[] {
  const page = doc.loadPage(index)
  let png: Uint8Array
  let widthPt: number
  let heightPt: number
  try {
    const [x0, y0, x1, y1] = page.getBounds()
    widthPt = x1 - x0
    heightPt = y1 - y0
    const zoom = Math.min(4, Math.max(1.5, RENDER_LONG_EDGE / Math.max(widthPt, heightPt, 1)))
    const pix = page.toPixmap(m.Matrix.scale(zoom, zoom), m.ColorSpace.DeviceRGB, false, true)
    try {
      png = pix.asPNG().slice()
    } finally {
      pix.destroy()
    }
  } finally {
    page.destroy()
  }
  let result: ReturnType<OcrRecognizer>
  try {
    result = recognize(png, { widthPt, heightPt })
  } catch {
    return []
  }
  if (!result) return []
  const out: (Word & { last: boolean })[] = []
  for (const line of result.lines) {
    if (line.confidence < MIN_LINE_CONFIDENCE) continue
    const words = lineWords(line).filter((w) =>
      [...w.text].every((ch) => {
        const c = ch.codePointAt(0)!
        return (c >= 32 && c < 127) || (c >= 160 && c <= 255)
      }),
    )
    words.forEach((w, i) => out.push({ ...w, last: i === words.length - 1 }))
  }
  return out
}
