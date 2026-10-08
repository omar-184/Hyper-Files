import { describe, expect, it } from 'vitest'
import { PdfToolError, mergePdfs, ocrPdf, type OcrRecognizer } from '../src'
import { input, m, makeImagePdf, open, pageTexts } from './helpers'

/** One page of real text, long enough not to count as a scan. */
function textPdf(): Uint8Array {
  const doc = new m.PDFDocument()
  const res = doc.addObject({ Font: { F1: doc.addSimpleFont(new m.Font('Helvetica')) } })
  doc.insertPage(
    -1,
    doc.addPage([0, 0, 595, 842], 0, res, 'BT /F1 24 Tf 72 700 Td (Quarterly report) Tj ET'),
  )
  const bytes = doc.saveToBuffer('compress').asUint8Array().slice()
  doc.destroy()
  return bytes
}

/** A fake engine: two lines in the upper half of every page it is shown. */
function fakeEngine(calls: { width: number; height: number }[]): OcrRecognizer {
  return (png, page) => {
    const img = new m.Image(png)
    calls.push({ width: img.getWidth(), height: img.getHeight() })
    expect(page.widthPt).toBeGreaterThan(0)
    return {
      lines: [
        {
          text: 'Invoice 2041',
          confidence: 0.95,
          box: { x0: 0.1, y0: 0.8, x1: 0.5, y1: 0.84 },
          chars: [
            ...'Invoice'
              .split('')
              .map((t) => ({ text: t, box: { x0: 0.1, y0: 0.8, x1: 0.32, y1: 0.84 } })),
            { text: ' ', box: { x0: 0, y0: 0, x1: 0, y1: 0 } },
            ...'2041'
              .split('')
              .map((t) => ({ text: t, box: { x0: 0.35, y0: 0.8, x1: 0.5, y1: 0.84 } })),
          ],
        },
        { text: 'Total due', confidence: 0.9, box: { x0: 0.1, y0: 0.7, x1: 0.4, y1: 0.73 } },
        { text: 'smudge', confidence: 0.1, box: { x0: 0.1, y0: 0.6, x1: 0.2, y1: 0.62 } },
        { text: 'مرحبا', confidence: 0.9, box: { x0: 0.6, y0: 0.6, x1: 0.8, y1: 0.62 } },
      ],
    }
  }
}

describe('ocr', () => {
  it('adds invisible text to scanned pages only and keeps the page looks', () => {
    const merged = mergePdfs(m, [
      input(textPdf(), 'text.pdf'),
      input(makeImagePdf(40, 40), 'scan.pdf'),
    ])
    const calls: { width: number; height: number }[] = []
    const progress: [number, number][] = []
    const out = ocrPdf(m, input(merged.bytes, 'mixed.pdf'), fakeEngine(calls), {}, (d, t) =>
      progress.push([d, t]),
    )
    expect(out.name).toBe('mixed-searchable.pdf')
    expect(out.stats).toEqual({ scanned: 1, recognized: 1 })
    // only the scan was rendered, at about 2k pixels on the long edge
    expect(calls).toHaveLength(1)
    expect(Math.max(calls[0].width, calls[0].height)).toBeGreaterThan(1900)
    expect(progress.at(-1)).toEqual([1, 1])

    const texts = pageTexts(out.bytes)
    expect(texts[0]).toBe('Quarterly report')
    expect(texts[1]).toContain('Invoice 2041')
    expect(texts[1]).toContain('Total due')
    expect(texts[1]).not.toContain('smudge')
    expect(texts[1]).not.toContain('?')

    // the text sits where the engine saw it: upper part of the page
    const doc = open(out.bytes)
    const page = doc.loadPage(1)
    const st = page.toStructuredText('')
    const hits = st.search('Invoice', '')
    expect(hits.length).toBe(1)
    const quad = hits[0][0]
    expect(quad[0]).toBeCloseTo(59.5, -1) // 0.1 * 595
    expect(quad[1]).toBeLessThan(842 * 0.25)
    // invisible: rendering the page draws nothing new
    const before = open(merged.bytes)
      .loadPage(1)
      .toPixmap(m.Matrix.identity, m.ColorSpace.DeviceGray)
    const after = page.toPixmap(m.Matrix.identity, m.ColorSpace.DeviceGray)
    expect(Buffer.from(after.getPixels()).equals(Buffer.from(before.getPixels()))).toBe(true)
    st.destroy()
    page.destroy()
    doc.destroy()
  })

  it('says so when every page already has text', () => {
    expect(() => ocrPdf(m, input(textPdf()), fakeEngine([]))).toThrow(PdfToolError)
  })

  it('fails when the engine recognizes nothing', () => {
    const err = (() => {
      try {
        ocrPdf(m, input(makeImagePdf(20, 20)), () => null)
      } catch (e) {
        return e as PdfToolError
      }
    })()
    expect(err?.code).toBe('unsupported')
  })
})
