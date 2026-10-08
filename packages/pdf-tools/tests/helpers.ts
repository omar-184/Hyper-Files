import * as m from 'mupdf'
import type { ToolInput } from '../src'

export { m }

m.setLog(null)

/** A PDF whose page i reads "Page i+1", built with a real font resource. */
export function makePdf(
  pages: number,
  opts: { size?: [number, number]; rotate?: 0 | 90 | 180 | 270 } = {},
): Uint8Array {
  const [w, h] = opts.size ?? [595, 842]
  const doc = new m.PDFDocument()
  const font = doc.addSimpleFont(new m.Font('Helvetica'))
  for (let i = 0; i < pages; i++) {
    const res = doc.addObject({ Font: { F1: font } })
    const page = doc.addPage(
      [0, 0, w, h],
      opts.rotate ?? 0,
      res,
      `BT /F1 24 Tf 72 ${h - 100} Td (Page ${i + 1}) Tj ET`,
    )
    doc.insertPage(-1, page)
  }
  const bytes = doc.saveToBuffer('compress').asUint8Array().slice()
  doc.destroy()
  return bytes
}

export function input(bytes: Uint8Array, name = 'sample.pdf', password?: string): ToolInput {
  return password ? { name, bytes, password } : { name, bytes }
}

export function open(bytes: Uint8Array, password?: string) {
  const doc = new m.PDFDocument(bytes)
  if (password) doc.authenticatePassword(password)
  return doc
}

/** The text of every page, trimmed. */
export function pageTexts(bytes: Uint8Array, password?: string): string[] {
  const doc = open(bytes, password)
  const out: string[] = []
  for (let p = 0; p < doc.countPages(); p++) {
    const page = doc.loadPage(p)
    const st = page.toStructuredText('')
    out.push(st.asText().trim())
    st.destroy()
    page.destroy()
  }
  doc.destroy()
  return out
}

/** A w x h RGB noise image as PNG: noise does not compress, like a photo. */
export function makePng(w: number, h: number, seed = 1): Uint8Array {
  const pix = new m.Pixmap(m.ColorSpace.DeviceRGB, [0, 0, w, h], false)
  const px = pix.getPixels()
  // xorshift32: integer-exact, unlike a float LCG whose patterns would compress
  let s = seed | 1
  for (let i = 0; i < px.length; i++) {
    s ^= s << 13
    s ^= s >>> 17
    s ^= s << 5
    px[i] = s & 255
  }
  const png = pix.asPNG().slice()
  pix.destroy()
  return png
}

/** A one-page PDF holding one big image, as a scanner would produce. */
export function makeImagePdf(w: number, h: number): Uint8Array {
  const doc = new m.PDFDocument()
  const image = new m.Image(makePng(w, h))
  const res = doc.addObject({ XObject: { Im0: doc.addImage(image) } })
  doc.insertPage(-1, doc.addPage([0, 0, 595, 842], 0, res, 'q 595 0 0 842 0 0 cm /Im0 Do Q'))
  const bytes = doc.saveToBuffer('compress').asUint8Array().slice()
  doc.destroy()
  return bytes
}
