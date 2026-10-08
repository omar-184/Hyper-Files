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

export interface CompressOptions {
  /** images are scaled down to at most this resolution at full-page size */
  dpi: number
  /** JPEG quality 1-100 for re-encoded images */
  quality: number
}

export const COMPRESS_PRESETS = {
  /** smallest file; fine for reading on screen */
  strong: { dpi: 100, quality: 55 },
  /** good balance for sharing */
  balanced: { dpi: 150, quality: 70 },
  /** keeps print quality; mostly removes waste */
  light: { dpi: 220, quality: 85 },
} as const satisfies Record<string, CompressOptions>

export interface CompressResult extends ToolOutput {
  originalSize: number
  imagesRewritten: number
}

/** Largest page side in inches: an image wider than this at `dpi` is wasted. */
function maxPageInches(doc: PDFDocument): number {
  let max = 0
  const count = doc.countPages()
  for (let p = 0; p < count; p++) {
    const page = doc.loadPage(p)
    const [x0, y0, x1, y1] = page.getBounds()
    page.destroy()
    max = Math.max(max, x1 - x0, y1 - y0)
  }
  return max / 72
}

function nameOf(obj: PDFObject): string | null {
  const r = deref(obj)
  return r.isName() ? r.asName() : null
}

/** Image XObjects we can safely re-encode as JPEG. */
function isRewritable(dict: PDFObject): boolean {
  if (nameOf(dict.get('Subtype')) !== 'Image') return false
  // stencil masks and colour-key masks depend on exact sample values
  const imageMask = dict.get('ImageMask')
  if (imageMask.isBoolean() && imageMask.asBoolean()) return false
  if (deref(dict.get('Mask')).isArray()) return false
  // an image used as another image's soft mask is grey alpha, not a picture
  if (!dict.get('SMaskInData').isNull()) return false
  const bpc = dict.get('BitsPerComponent')
  if (bpc.isNumber() && bpc.asNumber() < 8) return false
  return true
}

/** Objects referenced as /SMask by some image: never recompress these lossy. */
function softMaskNumbers(doc: PDFDocument): Set<number> {
  const out = new Set<number>()
  const n = doc.countObjects()
  for (let num = 1; num < n; num++) {
    const obj = deref(doc.newIndirect(num))
    if (!obj.isStream() && !obj.isDictionary()) continue
    const sm = obj.get('SMask')
    if (sm.isIndirect()) out.add(sm.asIndirect())
  }
  return out
}

function rewriteImages(m: Mupdf, doc: PDFDocument, opts: CompressOptions): number {
  const maxPixels = Math.max(16, Math.round(maxPageInches(doc) * opts.dpi))
  const masks = softMaskNumbers(doc)
  let rewritten = 0
  const n = doc.countObjects()
  for (let num = 1; num < n; num++) {
    if (masks.has(num)) continue
    const ref = doc.newIndirect(num)
    const dict = deref(ref)
    // streams are only recognisable through their reference; resolve() yields the dictionary
    if (!ref.isStream() || !isRewritable(dict)) continue
    let image
    try {
      image = doc.loadImage(ref)
    } catch {
      continue // undecodable image: leave it as it is
    }
    try {
      const w = image.getWidth()
      const h = image.getHeight()
      const scale = Math.min(1, maxPixels / Math.max(w, h))
      const filter = nameOf(dict.get('Filter')) ?? ''
      // already a small JPEG with nothing to gain from scaling: leave it
      if (scale === 1 && filter === 'DCTDecode') continue
      let pix = image.toPixmap()
      const cs = pix.getColorSpace()
      const wantGray = cs?.isGray() ?? false
      const target = wantGray ? m.ColorSpace.DeviceGray : m.ColorSpace.DeviceRGB
      if (pix.getAlpha() || !cs || cs.getType() !== (wantGray ? 'Gray' : 'RGB')) {
        const conv = pix.convertToColorSpace(target, false)
        pix.destroy()
        pix = conv
      }
      if (scale < 1) {
        const nw = Math.max(1, Math.round(w * scale))
        const nh = Math.max(1, Math.round(h * scale))
        const warped = pix.warp(
          [
            [0, 0],
            [w, 0],
            [w, h],
            [0, h],
          ],
          nw,
          nh,
        )
        pix.destroy()
        pix = warped
      }
      const jpeg = pix.asJPEG(Math.round(Math.min(100, Math.max(1, opts.quality))))
      const nw = pix.getWidth()
      const nh = pix.getHeight()
      pix.destroy()
      const before = ref.readRawStream()
      const oldSize = before.getLength()
      before.destroy()
      if (jpeg.length >= oldSize) continue
      dict.put('Width', nw)
      dict.put('Height', nh)
      dict.put('BitsPerComponent', 8)
      dict.put('ColorSpace', doc.newName(wantGray ? 'DeviceGray' : 'DeviceRGB'))
      dict.put('Filter', doc.newName('DCTDecode'))
      dict.delete('DecodeParms')
      dict.delete('Decode')
      dict.delete('Intent')
      ref.writeRawStream(jpeg)
      rewritten++
    } finally {
      image.destroy()
    }
  }
  return rewritten
}

/**
 * Shrink a PDF: re-encode oversized images as JPEG at the chosen resolution,
 * then drop unused and duplicate objects, subset fonts and pack objects into
 * compressed object streams. Returns the original when nothing got smaller.
 */
export function compressPdf(m: Mupdf, input: ToolInput, opts: CompressOptions): CompressResult {
  return withPdf(m, input, (doc) => {
    const imagesRewritten = rewriteImages(m, doc, opts)
    try {
      doc.subsetFonts()
    } catch {
      // a font MuPDF cannot subset stays whole; the rest of the work still counts
    }
    const bytes = savePdf(
      doc,
      'garbage=deduplicate,compress,compress-fonts,compress-images,clean,objstms,encrypt=none',
    )
    const smaller = bytes.length < input.bytes.length
    return {
      name: `${stem(input.name)}-compressed.pdf`,
      bytes: smaller ? bytes : input.bytes,
      originalSize: input.bytes.length,
      imagesRewritten: smaller ? imagesRewritten : 0,
    }
  })
}
