import { describe, expect, it } from 'vitest'
import {
  COMPRESS_PRESETS,
  PageRangeError,
  PdfToolError,
  addPageNumbers,
  addWatermark,
  compressPdf,
  deletePages,
  extractImages,
  extractPages,
  flattenPdf,
  formatPageList,
  imagesToPdf,
  mergePdfs,
  parsePageList,
  pdfInfo,
  pdfToImages,
  permissionBits,
  protectPdf,
  readMetadata,
  reorderPages,
  repairPdf,
  rotatePages,
  splitPdf,
  unlockPdf,
  writeMetadata,
  ALL_PERMISSIONS,
} from '../src'
import { input, m, makeImagePdf, makePdf, makePng, open, pageTexts } from './helpers'

describe('page ranges', () => {
  it('reads lists, open ends and keywords', () => {
    expect(parsePageList('1-3, 5, 8-', 10)).toEqual([0, 1, 2, 4, 7, 8, 9])
    expect(parsePageList('', 3)).toEqual([0, 1, 2])
    expect(parsePageList('-2', 5)).toEqual([0, 1])
    expect(parsePageList('last', 4)).toEqual([3])
    expect(parsePageList('3-1', 4)).toEqual([2, 1, 0])
    expect(parsePageList('2-end', 3)).toEqual([1, 2])
  })

  it('rejects pages outside the document and junk', () => {
    expect(() => parsePageList('0', 3)).toThrow(PageRangeError)
    expect(() => parsePageList('4', 3)).toThrow(/outside/)
    expect(() => parsePageList('a-b', 3)).toThrow(PageRangeError)
    expect(() => parsePageList(',', 3)).toThrow(/no pages/)
  })

  it('formats runs compactly', () => {
    expect(formatPageList([0, 1, 2, 4, 6, 7])).toBe('1-3,5,7-8')
  })
})

describe('opening', () => {
  it('refuses a file that is not a PDF', () => {
    const bad = input(new TextEncoder().encode('hello'), 'notes.txt')
    expect(() => pdfInfo(m, bad)).toThrow(PdfToolError)
    try {
      pdfInfo(m, bad)
    } catch (err) {
      expect((err as PdfToolError).code).toBe('not-pdf')
    }
  })

  it('reports page count and size', () => {
    const info = pdfInfo(m, input(makePdf(3, { size: [612, 792] })))
    expect(info).toMatchObject({ pages: 3, encrypted: false, width: 612, height: 792 })
  })
})

describe('page tools', () => {
  it('merges files in order, honouring page selections', () => {
    const a = makePdf(3)
    const b = makePdf(2)
    const out = mergePdfs(m, [{ ...input(a, 'a.pdf'), pages: '2-3' }, input(b, 'b.pdf')])
    expect(out.name).toBe('a-merged.pdf')
    expect(pageTexts(out.bytes)).toEqual(['Page 2', 'Page 3', 'Page 1', 'Page 2'])
  })

  it('splits every N pages and by ranges', () => {
    const src = input(makePdf(5), 'doc.pdf')
    const every = splitPdf(m, src, { kind: 'every', size: 2 })
    expect(every.map((o) => o.name)).toEqual([
      'doc-pages-1-2.pdf',
      'doc-pages-3-4.pdf',
      'doc-page-5.pdf',
    ])
    expect(pageTexts(every[2].bytes)).toEqual(['Page 5'])
    const ranges = splitPdf(m, src, { kind: 'ranges', ranges: '1, 2-4' })
    expect(ranges.map((o) => pageTexts(o.bytes))).toEqual([
      ['Page 1'],
      ['Page 2', 'Page 3', 'Page 4'],
    ])
    const single = splitPdf(m, src, { kind: 'pages', pages: '' })
    expect(single).toHaveLength(5)
  })

  it('extracts, deletes and reorders', () => {
    const src = input(makePdf(4))
    expect(pageTexts(extractPages(m, src, '4, 1').bytes)).toEqual(['Page 4', 'Page 1'])
    expect(pageTexts(deletePages(m, src, '2-3').bytes)).toEqual(['Page 1', 'Page 4'])
    expect(pageTexts(reorderPages(m, src, 'reverse').bytes)).toEqual([
      'Page 4',
      'Page 3',
      'Page 2',
      'Page 1',
    ])
    expect(() => deletePages(m, src, '1-4')).toThrow(/no pages/)
  })

  it('rotates selected pages on top of their existing rotation', () => {
    const src = input(makePdf(3, { rotate: 90 }))
    const out = rotatePages(m, src, 90, '1, 3')
    const doc = open(out.bytes)
    const rot = [0, 1, 2].map((p) => doc.findPage(p).getInheritable('Rotate').asNumber())
    expect(rot).toEqual([180, 90, 180])
    doc.destroy()
  })
})

describe('passwords', () => {
  it('protects with AES-256 and unlocks again', () => {
    const src = input(makePdf(2))
    const locked = protectPdf(m, src, { userPassword: 'open-me', ownerPassword: 'owner' })
    expect(() => pdfInfo(m, input(locked.bytes))).toThrow(/needs a password/)
    expect(() => pdfInfo(m, input(locked.bytes, 'x.pdf', 'nope'))).toThrow(/wrong password/)
    const info = pdfInfo(m, input(locked.bytes, 'x.pdf', 'open-me'))
    expect(info.encrypted).toBe(true)
    const open = unlockPdf(m, input(locked.bytes, 'x.pdf', 'open-me'))
    expect(pdfInfo(m, input(open.bytes)).encrypted).toBe(false)
    expect(pageTexts(open.bytes)).toEqual(['Page 1', 'Page 2'])
  })

  it('writes restriction bits a reader can check', () => {
    const src = input(makePdf(1))
    const locked = protectPdf(m, src, {
      userPassword: '',
      ownerPassword: 'owner',
      permissions: { ...ALL_PERMISSIONS, print: false, copy: false },
    })
    const doc = m.Document.openDocument(locked.bytes, 'application/pdf')
    expect(doc.needsPassword()).toBe(false)
    expect(doc.hasPermission('print')).toBe(false)
    expect(doc.hasPermission('copy')).toBe(false)
    expect(doc.hasPermission('annotate')).toBe(true)
    doc.destroy()
    expect(permissionBits(ALL_PERMISSIONS) & 0b11).toBe(0)
  })

  it('refuses commas, which the save options cannot carry', () => {
    expect(() => protectPdf(m, input(makePdf(1)), { userPassword: 'a,b' })).toThrow(/comma/)
  })
})

describe('compress', () => {
  it('downsamples an oversized scan', () => {
    const src = input(makeImagePdf(2400, 3400))
    const out = compressPdf(m, src, COMPRESS_PRESETS.balanced)
    expect(out.imagesRewritten).toBe(1)
    expect(out.bytes.length).toBeLessThan(src.bytes.length / 4)
    const doc = open(out.bytes)
    const page = doc.loadPage(0)
    const pix = page.toPixmap(m.Matrix.identity, m.ColorSpace.DeviceRGB)
    expect(pix.getWidth()).toBe(595)
    pix.destroy()
    page.destroy()
    doc.destroy()
  })

  it('returns the original when nothing gets smaller', () => {
    const src = input(makePdf(1))
    const out = compressPdf(m, src, COMPRESS_PRESETS.light)
    expect(out.bytes.length).toBeLessThanOrEqual(src.bytes.length)
    expect(pageTexts(out.bytes)).toEqual(['Page 1'])
  })
})

describe('images', () => {
  it('turns images into pages, fitting them to paper', () => {
    const out = imagesToPdf(
      m,
      [input(makePng(400, 300), 'wide.png'), input(makePng(300, 400), 'tall.png')],
      { pageSize: 'a4', orientation: 'auto', margin: 20 },
    )
    expect(out.name).toBe('wide.pdf')
    const doc = open(out.bytes)
    expect(doc.countPages()).toBe(2)
    const [, , w0, h0] = doc.loadPage(0).getBounds()
    const [, , w1, h1] = doc.loadPage(1).getBounds()
    expect(w0).toBeGreaterThan(h0) // landscape for the wide image
    expect(h1).toBeGreaterThan(w1)
    doc.destroy()
  })

  it('makes a page the size of the image with "fit"', () => {
    const out = imagesToPdf(m, [input(makePng(96, 192), 'a.png')], {
      pageSize: 'fit',
      orientation: 'auto',
      margin: 0,
    })
    const doc = open(out.bytes)
    expect(doc.loadPage(0).getBounds()).toEqual([0, 0, 72, 144])
    doc.destroy()
  })

  it('refuses a file that is not an image', () => {
    expect(() =>
      imagesToPdf(m, [input(new Uint8Array([1, 2, 3]), 'x.png')], {
        pageSize: 'fit',
        orientation: 'auto',
        margin: 0,
      }),
    ).toThrow(PdfToolError)
  })

  it('renders pages to PNG and JPEG at the requested resolution', () => {
    const src = input(makePdf(3, { size: [72, 144] }), 'r.pdf')
    const png = pdfToImages(m, src, { format: 'png', dpi: 144, pages: '2-3' })
    expect(png.map((o) => o.name)).toEqual(['r-page-2.png', 'r-page-3.png'])
    const img = new m.Image(png[0].bytes)
    expect([img.getWidth(), img.getHeight()]).toEqual([144, 288])
    const jpg = pdfToImages(m, src, { format: 'jpeg', dpi: 72, quality: 80, pages: '1' })
    expect(jpg[0].name).toBe('r-page-1.jpg')
    expect([...jpg[0].bytes.slice(0, 2)]).toEqual([0xff, 0xd8])
  })

  it('pulls embedded images out', () => {
    const out = extractImages(m, input(makeImagePdf(64, 48), 'scan.pdf'))
    expect(out.map((o) => o.name)).toEqual(['scan-image-1.png'])
    const img = new m.Image(out[0].bytes)
    expect([img.getWidth(), img.getHeight()]).toEqual([64, 48])
  })
})

describe('stamps', () => {
  it('adds a watermark on the chosen pages without losing the content', () => {
    const src = input(makePdf(3))
    const out = addWatermark(m, src, {
      text: 'CONFIDENTIAL',
      fontSize: 60,
      color: '#cc0000',
      opacity: 0.3,
      rotation: 'diagonal',
      position: 'center',
      pages: '1, 3',
    })
    const texts = pageTexts(out.bytes)
    expect(texts[0]).toContain('Page 1')
    expect(texts[0]).toContain('CONFIDENTIAL')
    expect(texts[1]).toBe('Page 2')
    expect(texts[2]).toContain('CONFIDENTIAL')
  })

  it('keeps a watermark upright on a rotated page', () => {
    const src = input(makePdf(1, { rotate: 90 }))
    const out = addWatermark(m, src, {
      text: 'DRAFT',
      fontSize: 40,
      color: '#000000',
      opacity: 1,
      rotation: 0,
      position: 'center',
    })
    const doc = open(out.bytes)
    const page = doc.loadPage(0)
    const hits = page.search('DRAFT', {})
    expect(hits).toHaveLength(1)
    const [ulx, uly, urx, ury] = hits[0][0]
    // horizontal on screen: upper-left and upper-right share a y
    expect(Math.abs(uly - ury)).toBeLessThan(0.5)
    expect(urx).toBeGreaterThan(ulx)
    page.destroy()
    doc.destroy()
  })

  it('numbers pages from a chosen start, skipping a cover page', () => {
    const src = input(makePdf(4))
    const out = addPageNumbers(m, src, {
      format: 'Page {n} of {total}',
      position: 'bottom-center',
      fontSize: 10,
      margin: 24,
      firstNumber: 1,
      pages: '2-',
    })
    const texts = pageTexts(out.bytes)
    expect(texts[0]).toBe('Page 1')
    expect(texts[1]).toContain('Page 1 of 3')
    expect(texts[3]).toContain('Page 3 of 3')
  })
})

describe('other tools', () => {
  it('flattens annotations into the page', () => {
    const doc = open(makePdf(1))
    const page = doc.loadPage(0)
    const annot = page.createAnnotation('FreeText')
    annot.setRect([100, 100, 300, 140])
    annot.setContents('Stamped note')
    annot.update()
    const bytes = doc.saveToBuffer('').asUint8Array().slice()
    doc.destroy()
    const out = flattenPdf(m, input(bytes))
    const flat = open(out.bytes)
    expect(flat.loadPage(0).getAnnotations()).toHaveLength(0)
    flat.destroy()
    expect(pageTexts(out.bytes)[0]).toContain('Stamped note')
  })

  it('repairs a file with a broken cross-reference table', () => {
    const good = makePdf(2)
    const text = Buffer.from(good).toString('latin1')
    const broken = new Uint8Array(
      Buffer.from(text.replace(/startxref\s+\d+/, 'startxref\n999999'), 'latin1'),
    )
    const out = repairPdf(m, input(broken))
    expect(out.repaired).toBe(true)
    expect(pageTexts(out.bytes)).toEqual(['Page 1', 'Page 2'])
  })

  it('reads and writes document properties', () => {
    const out = writeMetadata(m, input(makePdf(1)), { title: 'Annual report', author: 'Omar' })
    expect(readMetadata(m, input(out.bytes))).toMatchObject({
      title: 'Annual report',
      author: 'Omar',
    })
  })
})
