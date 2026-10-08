/**
 * Find-and-redact on pdfium: locate every occurrence of a phrase or
 * sensitive-text pattern in the file and return the exact glyph boxes. The
 * viewer's search rects interpolate positions by character count, which can
 * miss the first or last glyph of a match in proportional fonts; a redaction
 * mark must cover every glyph, so marks come from pdfium's per-char boxes.
 */
import { chainPdfium, loadPdfium, withDocument, type Pdfium } from './text-edit'
import { matchRanges, type FindTarget } from '../shared/text-match'
import type { FindTextMatch } from '../shared/ipc'

/** Same cap as the viewer's search */
export const MAX_FIND_MATCHES = 1000

type Box = [number, number, number, number]

/** Raw textpage text: one UTF-16 unit per char index, pdfium's generated \r\n kept so offsets line up */
function rawPageText(m: Pdfium, textPage: number, count: number): string {
  if (count <= 0) return ''
  const buf = m._malloc((count + 1) * 2)
  if (!buf) return ''
  try {
    const written = m._FPDFText_GetText(textPage, 0, count, buf)
    if (written <= 0) return ''
    return Buffer.from(m.HEAPU8.subarray(buf, buf + Math.min(written, count) * 2)).toString(
      'utf16le',
    )
  } finally {
    m._free(buf)
  }
}

/** Loose box (font ascent to descent) of one char as [x1, y1, x2, y2]; null for generated or empty chars */
function charBox(m: Pdfium, textPage: number, index: number, ptr: number): Box | null {
  // FS_RECTF {left, top, right, bottom}
  if (!m._FPDFText_GetLooseCharBox(textPage, index, ptr)) return null
  const f = m.HEAPF32
  const l = f[ptr >> 2]!
  const t = f[(ptr >> 2) + 1]!
  const r = f[(ptr >> 2) + 2]!
  const b = f[(ptr >> 2) + 3]!
  const box: Box = [Math.min(l, r), Math.min(t, b), Math.max(l, r), Math.max(t, b)]
  if (!box.every(Number.isFinite) || box[2] - box[0] <= 0 || box[3] - box[1] <= 0) return null
  return box
}

/**
 * Union consecutive char boxes into one rect per line: a char starts a new rect
 * when its vertical middle leaves the current rect's band.
 */
export function lineRects(boxes: readonly Box[]): Box[] {
  const out: Box[] = []
  for (const b of boxes) {
    const cur = out[out.length - 1]
    const mid = (b[1] + b[3]) / 2
    if (cur && mid >= cur[1] && mid <= cur[3]) {
      cur[0] = Math.min(cur[0], b[0])
      cur[1] = Math.min(cur[1], b[1])
      cur[2] = Math.max(cur[2], b[2])
      cur[3] = Math.max(cur[3], b[3])
    } else out.push([...b])
  }
  return out
}

export function findTextBoxes(bytes: Uint8Array, target: FindTarget): Promise<FindTextMatch[]> {
  return chainPdfium(async () => {
    const m = await loadPdfium()
    return withDocument(m, bytes, async (doc) => {
      const matches: FindTextMatch[] = []
      const ptr = m._malloc(16)
      try {
        const pageCount = m._FPDF_GetPageCount(doc)
        for (let pageIndex = 0; pageIndex < pageCount; pageIndex++) {
          if (matches.length >= MAX_FIND_MATCHES) break
          const page = m._FPDF_LoadPage(doc, pageIndex)
          if (!page) continue
          try {
            const textPage = m._FPDFText_LoadPage(page)
            if (!textPage) continue
            try {
              const text = rawPageText(m, textPage, m._FPDFText_CountChars(textPage))
              for (const [s, e] of matchRanges(text, target, MAX_FIND_MATCHES - matches.length)) {
                const boxes: Box[] = []
                for (let i = s; i < e; i++) {
                  const box = charBox(m, textPage, i, ptr)
                  if (box) boxes.push(box)
                }
                if (boxes.length > 0) matches.push({ pageIndex, rects: lineRects(boxes) })
              }
            } finally {
              m._FPDFText_ClosePage(textPage)
            }
          } finally {
            m._FPDF_ClosePage(page)
          }
        }
      } finally {
        m._free(ptr)
      }
      return matches
    })
  })
}
