import { PDFHexString, PDFName, PDFNull, PDFNumber, PDFString } from 'pdf-lib'
import type { PDFDict, PDFDocument, PDFObject, PDFPage, PDFRef } from 'pdf-lib'
import type { OutlineEntryInput } from '../shared/ipc'

/** Same bound the viewer renders; deeper entries are dropped */
const MAX_DEPTH = 32
const MAX_ENTRIES = 20000
const FITS = new Set(['XYZ', 'Fit', 'FitH', 'FitV', 'FitR', 'FitB', 'FitBH', 'FitBV'])

/**
 * Replace the document outline (/Outlines) with the given tree. Destinations
 * are explicit [page /Fit args...] arrays on the original page objects, so a
 * later reorder in the same save keeps them pointing at the right page; an
 * entry whose page is deleted keeps its title without a destination.
 */
export function writeOutline(
  pdfDoc: PDFDocument,
  pages: PDFPage[],
  entries: OutlineEntryInput[],
  deleted: ReadonlySet<number>,
): void {
  const ctx = pdfDoc.context
  if (entries.length === 0) {
    pdfDoc.catalog.delete(PDFName.of('Outlines'))
    return
  }
  let budget = MAX_ENTRIES

  const destFor = (e: OutlineEntryInput): PDFObject | null => {
    if (e.pageIndex === null || deleted.has(e.pageIndex)) return null
    const page = pages[e.pageIndex]
    if (!page) return null
    const fit = e.fit && FITS.has(e.fit) ? e.fit : 'XYZ'
    const args = (e.fit && FITS.has(e.fit) ? (e.args ?? []) : [null, null, null]).map((v) =>
      typeof v === 'number' && Number.isFinite(v) ? PDFNumber.of(v) : PDFNull,
    )
    return ctx.obj([page.ref, PDFName.of(fit), ...args])
  }

  /** Writes the siblings under parent; returns [first, last, visible descendant count] */
  const writeLevel = (
    list: OutlineEntryInput[],
    parent: PDFRef,
    depth: number,
  ): [PDFRef, PDFRef, number] | null => {
    const refs: PDFRef[] = []
    const dicts: PDFDict[] = []
    let count = 0
    for (const e of list) {
      if (budget-- <= 0) break
      const dict = ctx.obj({}) as PDFDict
      const ref = ctx.register(dict)
      dict.set(PDFName.of('Title'), PDFHexString.fromText(e.title))
      dict.set(PDFName.of('Parent'), parent)
      if (e.url) {
        dict.set(PDFName.of('A'), ctx.obj({ S: 'URI', URI: PDFString.of(e.url) }))
      } else {
        const dest = destFor(e)
        if (dest) dict.set(PDFName.of('Dest'), dest)
      }
      const flags = (e.italic ? 1 : 0) | (e.bold ? 2 : 0)
      if (flags) dict.set(PDFName.of('F'), PDFNumber.of(flags))
      if (e.color && e.color.some((c) => c > 0)) dict.set(PDFName.of('C'), ctx.obj(e.color))
      if (depth + 1 < MAX_DEPTH && e.items.length > 0) {
        const kids = writeLevel(e.items, ref, depth + 1)
        if (kids) {
          dict.set(PDFName.of('First'), kids[0])
          dict.set(PDFName.of('Last'), kids[1])
          // Positive: shown expanded
          dict.set(PDFName.of('Count'), PDFNumber.of(kids[2]))
          count += kids[2]
        }
      }
      refs.push(ref)
      dicts.push(dict)
      count++
    }
    if (refs.length === 0) return null
    for (let i = 0; i < refs.length; i++) {
      if (i > 0) dicts[i]!.set(PDFName.of('Prev'), refs[i - 1]!)
      if (i < refs.length - 1) dicts[i]!.set(PDFName.of('Next'), refs[i + 1]!)
    }
    return [refs[0]!, refs[refs.length - 1]!, count]
  }

  const root = ctx.obj({ Type: 'Outlines' }) as PDFDict
  const rootRef = ctx.register(root)
  const top = writeLevel(entries, rootRef, 0)
  if (!top) {
    pdfDoc.catalog.delete(PDFName.of('Outlines'))
    return
  }
  root.set(PDFName.of('First'), top[0])
  root.set(PDFName.of('Last'), top[1])
  root.set(PDFName.of('Count'), PDFNumber.of(top[2]))
  pdfDoc.catalog.set(PDFName.of('Outlines'), rootRef)
}
