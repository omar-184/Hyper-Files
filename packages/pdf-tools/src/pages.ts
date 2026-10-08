import {
  type Mupdf,
  type PDFDocument,
  PdfToolError,
  type ToolInput,
  type ToolOutput,
  openPdf,
  savePdf,
  stem,
  withPdf,
} from './core'
import { formatPageList, parsePageList, parsePageRuns, parsePageSet } from './page-ranges'

/** Copy `pages` (0-based, in order) of `src` onto the end of `dst`. */
function appendPages(dst: PDFDocument, src: PDFDocument, pages: readonly number[]) {
  // one graft map per source shares fonts and images between its pages
  const map = dst.newGraftMap()
  try {
    for (const p of pages) map.graftPage(-1, src, p)
  } finally {
    map.destroy()
  }
}

export interface MergeItem extends ToolInput {
  /** page-range expression; empty means every page */
  pages?: string
}

/** Join PDFs in the order given, optionally taking only some pages of each. */
export function mergePdfs(m: Mupdf, items: readonly MergeItem[], outName?: string): ToolOutput {
  if (items.length === 0) throw new PdfToolError('bad-input', 'no files to merge')
  const out = new m.PDFDocument()
  try {
    for (const item of items) {
      withPdf(m, item, (src) => {
        appendPages(out, src, parsePageList(item.pages ?? '', src.countPages()))
      })
    }
    return {
      name: outName ?? `${stem(items[0].name)}-merged.pdf`,
      bytes: savePdf(out),
    }
  } finally {
    out.destroy()
  }
}

/** Build a new PDF from `pages` of an already-open document. */
function subset(m: Mupdf, src: PDFDocument, pages: readonly number[]): Uint8Array {
  const out = new m.PDFDocument()
  try {
    appendPages(out, src, pages)
    return savePdf(out)
  } finally {
    out.destroy()
  }
}

export type SplitMode =
  /** one file per `size` pages */
  | { kind: 'every'; size: number }
  /** one file per comma-separated range, e.g. "1-3, 4-10, 11-" */
  | { kind: 'ranges'; ranges: string }
  /** one file per page, only for the pages listed (empty: all) */
  | { kind: 'pages'; pages: string }

export function splitPdf(m: Mupdf, input: ToolInput, mode: SplitMode): ToolOutput[] {
  return withPdf(m, input, (src) => {
    const count = src.countPages()
    const groups: number[][] = []
    if (mode.kind === 'every') {
      const size = Math.floor(mode.size)
      if (!(size >= 1)) throw new PdfToolError('bad-input', 'pages per file must be at least 1')
      for (let p = 0; p < count; p += size) {
        groups.push(Array.from({ length: Math.min(size, count - p) }, (_, k) => p + k))
      }
    } else if (mode.kind === 'ranges') {
      for (const run of parsePageRuns(mode.ranges, count)) {
        const lo = Math.min(run.from, run.to)
        const hi = Math.max(run.from, run.to)
        groups.push(Array.from({ length: hi - lo + 1 }, (_, k) => lo + k))
      }
    } else {
      for (const p of [...parsePageSet(mode.pages, count)].sort((a, b) => a - b)) groups.push([p])
    }
    const base = stem(input.name)
    const width = String(count).length
    return groups.map((pages) => ({
      name:
        pages.length === 1
          ? `${base}-page-${String(pages[0] + 1).padStart(width, '0')}.pdf`
          : `${base}-pages-${formatPageList(pages)}.pdf`,
      bytes: subset(m, src, pages),
    }))
  })
}

/** Keep only the listed pages, in the order listed ("Extract pages"). */
export function extractPages(m: Mupdf, input: ToolInput, pages: string): ToolOutput {
  return withPdf(m, input, (src) => {
    const list = parsePageList(pages, src.countPages())
    return { name: `${stem(input.name)}-extract.pdf`, bytes: subset(m, src, list) }
  })
}

/** Rewrite a document in place with a new page order, keeping links and outlines. */
function rearranged(m: Mupdf, input: ToolInput, order: (count: number) => number[]): Uint8Array {
  const doc = openPdf(m, input)
  try {
    const list = order(doc.countPages())
    if (list.length === 0) {
      throw new PdfToolError('bad-input', 'the result would have no pages', input.name)
    }
    doc.rearrangePages(list)
    return savePdf(doc)
  } finally {
    doc.destroy()
  }
}

export function deletePages(m: Mupdf, input: ToolInput, pages: string): ToolOutput {
  return {
    name: `${stem(input.name)}-edited.pdf`,
    bytes: rearranged(m, input, (count) => {
      const drop = parsePageSet(pages, count)
      return Array.from({ length: count }, (_, p) => p).filter((p) => !drop.has(p))
    }),
  }
}

/**
 * Put pages in a new order. `order` is a page-range expression ("3, 1-2, 4-")
 * or the word "reverse"; pages it leaves out are dropped.
 */
export function reorderPages(m: Mupdf, input: ToolInput, order: string): ToolOutput {
  return {
    name: `${stem(input.name)}-reordered.pdf`,
    bytes: rearranged(m, input, (count) =>
      /^\s*reverse\s*$/i.test(order)
        ? Array.from({ length: count }, (_, p) => count - 1 - p)
        : parsePageList(order, count),
    ),
  }
}

export type RotateAngle = 90 | 180 | 270

/** Turn pages clockwise by `angle`; `pages` empty means every page. */
export function rotatePages(
  m: Mupdf,
  input: ToolInput,
  angle: RotateAngle,
  pages = '',
): ToolOutput {
  return withPdf(m, input, (doc) => {
    const set = parsePageSet(pages, doc.countPages())
    for (const p of set) {
      const obj = doc.findPage(p)
      const current = obj.getInheritable('Rotate')
      const was = current.isNumber() ? current.asNumber() : 0
      obj.put('Rotate', (((was + angle) % 360) + 360) % 360)
    }
    return { name: `${stem(input.name)}-rotated.pdf`, bytes: savePdf(doc) }
  })
}

/** Page count and size of the first page, for the tool panel's file list. */
export interface PdfInfo {
  pages: number
  encrypted: boolean
  /** first page size in points, as displayed (rotation applied) */
  width: number
  height: number
  title?: string
}

export function pdfInfo(m: Mupdf, input: ToolInput): PdfInfo {
  return withPdf(m, input, (doc) => {
    const page = doc.loadPage(0)
    try {
      const [x0, y0, x1, y1] = page.getBounds()
      const title = doc.getMetaData('info:Title')
      return {
        pages: doc.countPages(),
        encrypted: (doc.getMetaData('encryption') ?? 'None') !== 'None',
        width: x1 - x0,
        height: y1 - y0,
        ...(title ? { title } : {}),
      }
    } finally {
      page.destroy()
    }
  })
}
