import {
  deref,
  type Mupdf,
  PdfToolError,
  type ToolInput,
  type ToolOutput,
  savePdf,
  stem,
  withPdf,
} from './core'
import { parsePageSet } from './page-ranges'

/** Paper sizes in points, portrait. */
export const PAPER_SIZES = {
  a4: [595.28, 841.89],
  letter: [612, 792],
  legal: [612, 1008],
  a3: [841.89, 1190.55],
  a5: [419.53, 595.28],
} as const

export type PaperSize = keyof typeof PAPER_SIZES

export interface ImagesToPdfOptions {
  /** 'fit' makes each page the size of its image */
  pageSize: PaperSize | 'fit'
  orientation: 'auto' | 'portrait' | 'landscape'
  /** blank border in points around the image (paper sizes only) */
  margin: number
}

/** The image's size in points, from its own resolution (96 dpi when unset). */
function imageSizePt(width: number, height: number, xres: number, yres: number) {
  const xr = xres > 1 ? xres : 96
  const yr = yres > 1 ? yres : 96
  return [(width * 72) / xr, (height * 72) / yr] as const
}

/** One page per image, in the order given. JPEGs are embedded as they are. */
export function imagesToPdf(
  m: Mupdf,
  images: readonly ToolInput[],
  opts: ImagesToPdfOptions,
  outName?: string,
): ToolOutput {
  if (images.length === 0) throw new PdfToolError('bad-input', 'no images to convert')
  const doc = new m.PDFDocument()
  try {
    for (const input of images) {
      let image
      try {
        image = new m.Image(input.bytes)
      } catch {
        throw new PdfToolError(
          'bad-input',
          `${input.name} is not an image MuPDF can read`,
          input.name,
        )
      }
      try {
        const [iw, ih] = imageSizePt(
          image.getWidth(),
          image.getHeight(),
          image.getXResolution(),
          image.getYResolution(),
        )
        let pw: number
        let ph: number
        let margin = 0
        if (opts.pageSize === 'fit') {
          pw = iw
          ph = ih
        } else {
          const [a, b] = PAPER_SIZES[opts.pageSize]
          const landscape =
            opts.orientation === 'landscape' || (opts.orientation === 'auto' && iw > ih)
          pw = landscape ? b : a
          ph = landscape ? a : b
          margin = Math.max(0, Math.min(opts.margin, Math.min(pw, ph) / 2 - 1))
        }
        // fit inside the margins, never enlarging past the image's own size on paper
        const boxW = pw - 2 * margin
        const boxH = ph - 2 * margin
        const scale = opts.pageSize === 'fit' ? 1 : Math.min(boxW / iw, boxH / ih, 1)
        const dw = iw * scale
        const dh = ih * scale
        const x = (pw - dw) / 2
        const y = (ph - dh) / 2
        const resources = doc.addObject({ XObject: { Im0: doc.addImage(image) } })
        const content = `q ${dw.toFixed(3)} 0 0 ${dh.toFixed(3)} ${x.toFixed(3)} ${y.toFixed(3)} cm /Im0 Do Q`
        doc.insertPage(-1, doc.addPage([0, 0, pw, ph], 0, resources, content))
      } finally {
        image.destroy()
      }
    }
    return { name: outName ?? `${stem(images[0].name)}.pdf`, bytes: savePdf(doc) }
  } finally {
    doc.destroy()
  }
}

export interface PdfToImagesOptions {
  format: 'png' | 'jpeg'
  /** render resolution */
  dpi: number
  /** JPEG quality 1-100 */
  quality?: number
  /** page-range expression; empty means every page */
  pages?: string
}

/** Renders past this many pixels per page would need hundreds of MB on a 4 GB machine. */
const MAX_PIXELS = 50_000_000

/** Render pages to image files, one per page. */
export function pdfToImages(m: Mupdf, input: ToolInput, opts: PdfToImagesOptions): ToolOutput[] {
  return withPdf(m, input, (doc) => {
    const count = doc.countPages()
    const pages = [...parsePageSet(opts.pages ?? '', count)].sort((a, b) => a - b)
    const base = stem(input.name)
    const width = String(count).length
    const ext = opts.format === 'png' ? 'png' : 'jpg'
    const out: ToolOutput[] = []
    for (const p of pages) {
      const page = doc.loadPage(p)
      try {
        const [x0, y0, x1, y1] = page.getBounds()
        let zoom = opts.dpi / 72
        const px = (x1 - x0) * zoom * (y1 - y0) * zoom
        if (px > MAX_PIXELS) zoom *= Math.sqrt(MAX_PIXELS / px)
        // JPEG has no alpha; PNG keeps a white page too so viewers match the PDF
        const pix = page.toPixmap(m.Matrix.scale(zoom, zoom), m.ColorSpace.DeviceRGB, false, true)
        try {
          pix.setResolution(Math.round(zoom * 72), Math.round(zoom * 72))
          const bytes =
            opts.format === 'png'
              ? pix.asPNG().slice()
              : pix.asJPEG(Math.round(Math.min(100, Math.max(1, opts.quality ?? 85)))).slice()
          out.push({ name: `${base}-page-${String(p + 1).padStart(width, '0')}.${ext}`, bytes })
        } finally {
          pix.destroy()
        }
      } finally {
        page.destroy()
      }
    }
    return out
  })
}

/** Embedded images pulled out as they are stored, one file each. */
export function extractImages(m: Mupdf, input: ToolInput): ToolOutput[] {
  return withPdf(m, input, (doc) => {
    const out: ToolOutput[] = []
    const base = stem(input.name)
    const n = doc.countObjects()
    let index = 0
    for (let num = 1; num < n; num++) {
      const ref = doc.newIndirect(num)
      if (!ref.isStream()) continue
      const obj = deref(ref)
      const subtype = deref(obj.get('Subtype'))
      if (!subtype.isName() || subtype.asName() !== 'Image') continue
      const mask = obj.get('ImageMask')
      if (mask.isBoolean() && mask.asBoolean()) continue
      let image
      try {
        image = doc.loadImage(ref)
      } catch {
        continue
      }
      try {
        // skip tiny images: icons, bullets, 1-pixel spacers
        if (image.getWidth() < 16 || image.getHeight() < 16) continue
        index++
        const filter = deref(obj.get('Filter'))
        const isJpeg = filter.isName() && filter.asName() === 'DCTDecode'
        const cs = deref(obj.get('ColorSpace'))
        const cmyk = cs.isName() && cs.asName() === 'DeviceCMYK'
        if (isJpeg && !cmyk && obj.get('SMask').isNull()) {
          const raw = ref.readRawStream()
          out.push({ name: `${base}-image-${index}.jpg`, bytes: raw.asUint8Array().slice() })
          raw.destroy()
        } else {
          let pix = image.toPixmap()
          const pcs = pix.getColorSpace()
          if (pcs && !pcs.isRGB() && !pcs.isGray()) {
            const conv = pix.convertToColorSpace(m.ColorSpace.DeviceRGB, true)
            pix.destroy()
            pix = conv
          }
          out.push({ name: `${base}-image-${index}.png`, bytes: pix.asPNG().slice() })
          pix.destroy()
        }
      } finally {
        image.destroy()
      }
    }
    return out
  })
}
