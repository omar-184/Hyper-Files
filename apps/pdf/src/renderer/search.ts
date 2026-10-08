import type { PDFDocumentProxy } from 'pdfjs-dist'
import { foldCase } from '@genoffice/ui'
import { matchRanges, type FindTarget } from '../shared/text-match'

/** One hit: original page + PDF user-space rects (multiple when spanning several text items) */
export interface SearchMatch {
  pageIndex: number
  rects: [number, number, number, number][]
}

interface IndexedItem {
  start: number
  end: number
  x: number
  y: number
  w: number
  h: number
  /** Rotated run (tilted baseline) — excluded from block grouping */
  rot?: boolean
  /** pdf.js font id (e.g. 'g_d0_f7'); resolves to the run's font for edit previews */
  font?: string
}

export interface PageEntry {
  /** Original text (same length as lower; used for context excerpts) */
  text: string
  lower: string
  items: IndexedItem[]
}

export type SearchIndex = PageEntry[]

const MAX_MATCHES = 1000

interface RawTextItem {
  str?: string
  transform?: number[]
  width?: number
  height?: number
  hasEOL?: boolean
  fontName?: string
}

/** Concatenate text per page + record each item's char range and PDF-space box (built once, cached per doc by caller) */
export async function buildSearchIndex(doc: PDFDocumentProxy): Promise<SearchIndex> {
  const entries: PageEntry[] = []
  for (let n = 1; n <= doc.numPages; n++) {
    const page = await doc.getPage(n)
    const content = await page.getTextContent()
    let text = ''
    const items: IndexedItem[] = []
    for (const it of content.items as RawTextItem[]) {
      if (typeof it.str !== 'string') continue
      if (it.str.length > 0) {
        // Geometry is optional; without a transform the item still contributes its
        // characters, or it is unfindable while its hasEOL newline still lands.
        if (it.transform) {
          const h = it.height || Math.hypot(it.transform[2] ?? 0, it.transform[3] ?? 0)
          // Rotation tilts the baseline (b ≠ 0). A non-zero c alone is horizontal
          // shear — synthetic italics — which stays horizontally set and must keep
          // participating in block grouping.
          const rot = Math.abs(it.transform[1] ?? 0) > h * 1e-3
          items.push({
            start: text.length,
            end: text.length + it.str.length,
            x: it.transform[4] ?? 0,
            y: it.transform[5] ?? 0,
            w: it.width ?? 0,
            h,
            ...(rot ? { rot: true } : {}),
            ...(typeof it.fontName === 'string' ? { font: it.fontName } : {}),
          })
        }
        text += it.str
      }
      if (it.hasEOL) text += '\n'
    }
    entries.push({ text, lower: foldCase(text), items })
  }
  return entries
}

/** One text extraction per loaded document, shared by every consumer */
export interface SearchIndexCache {
  /** Base index of `doc`; the same promise for the same document until cleared */
  get(doc: PDFDocumentProxy): Promise<SearchIndex>
  /** Drop the cached extraction (a save-reload replaces the document content) */
  clear(): void
}

/** Building the index walks every page of the document, so it must happen once per
    loaded document: search, paragraph boxes and the auto-OCR pass all
    read this cache instead of extracting the text a second time. */
export function createSearchIndexCache(): SearchIndexCache {
  let entry: { doc: PDFDocumentProxy; promise: Promise<SearchIndex> } | null = null
  return {
    get(doc) {
      if (entry?.doc !== doc) entry = { doc, promise: buildSearchIndex(doc) }
      return entry.promise
    },
    clear() {
      entry = null
    },
  }
}

/** Rects of the char range [s, e) on one page; linearly interpolated within items by char ratio */
function rectsForRange(
  items: IndexedItem[],
  s: number,
  e: number,
): [number, number, number, number][] {
  const rects: [number, number, number, number][] = []
  for (const it of items) {
    if (it.end <= s || it.start >= e) continue
    const len = it.end - it.start
    const lo = (Math.max(s, it.start) - it.start) / len
    const hi = (Math.min(e, it.end) - it.start) / len
    const x1 = it.x + it.w * lo
    const x2 = it.x + it.w * hi
    if (x2 - x1 < 0.01) continue
    rects.push([x1, it.y, x2, it.y + it.h])
  }
  return rects
}

/** Case-insensitive full-text search; rects linearly interpolated within items by char ratio (approximate; bounding box for rotated glyphs) */
export function searchInIndex(index: SearchIndex, query: string): SearchMatch[] {
  return findInIndex(index, { query })
}

/** Every occurrence of a phrase or sensitive-text pattern, in page order */
export function findInIndex(index: SearchIndex, target: FindTarget): SearchMatch[] {
  const matches: SearchMatch[] = []
  for (let pageIndex = 0; pageIndex < index.length; pageIndex++) {
    const { text, lower, items } = index[pageIndex]!
    for (const [s, e] of matchRanges(text, target, MAX_MATCHES - matches.length, lower)) {
      const rects = rectsForRange(items, s, e)
      if (rects.length > 0) matches.push({ pageIndex, rects })
    }
    if (matches.length >= MAX_MATCHES) break
  }
  return matches
}

/** Text of a page lying inside the given PDF-space boxes (the words a highlight
    covers). Items count when at least half their height falls in a box; within an
    item the char range is cut by the same linear ratio search uses. */
export function textInRects(
  entry: PageEntry,
  boxes: readonly (readonly [number, number, number, number])[],
): string {
  const parts: string[] = []
  let lastItem = -1
  entry.items.forEach((it, i) => {
    if (it.w <= 0 || it.h <= 0) return
    let lo = Infinity
    let hi = -Infinity
    for (const [bx1, by1, bx2, by2] of boxes) {
      const overlapY =
        Math.min(it.y + it.h, Math.max(by1, by2)) - Math.max(it.y, Math.min(by1, by2))
      if (overlapY < it.h * 0.5) continue
      const x1 = Math.max(it.x, Math.min(bx1, bx2))
      const x2 = Math.min(it.x + it.w, Math.max(bx1, bx2))
      if (x2 <= x1) continue
      lo = Math.min(lo, x1)
      hi = Math.max(hi, x2)
    }
    if (hi <= lo) return
    const len = it.end - it.start
    const from = it.start + Math.round(((lo - it.x) / it.w) * len)
    const to = it.start + Math.round(((hi - it.x) / it.w) * len)
    const piece = entry.text.slice(from, to)
    if (!piece) return
    // Adjacent items continue a run; a gap means a new line or column
    if (parts.length > 0 && lastItem !== i - 1 && !/\s$/.test(parts[parts.length - 1]!))
      parts.push(' ')
    parts.push(piece)
    lastItem = i
  })
  return parts.join('').replace(/\s+/g, ' ').trim()
}
