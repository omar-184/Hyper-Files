import { describe, expect, it } from 'vitest'
import {
  PDFArray,
  PDFDict,
  PDFDocument,
  PDFName,
  PDFRawStream,
  decodePDFRawStream,
  degrees,
} from 'pdf-lib'
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs'
import { applySaveRequest } from '../src/main/save-pdf'
import { isWinAnsi, wrapTextBox } from '../src/shared/text-box'
import type { DrawingInput, SavePdfRequest } from '../src/shared/ipc'

/** Monospace stand-in: every char 6 pt wide */
const mono = (s: string) => s.length * 6

describe('wrapTextBox', () => {
  it('breaks at spaces, keeps paragraphs, and splits words wider than the box', () => {
    expect(wrapTextBox('aaa bbb ccc', 42, mono)).toEqual(['aaa bbb', 'ccc'])
    expect(wrapTextBox('one\ntwo', 100, mono)).toEqual(['one', 'two'])
    expect(wrapTextBox('abcdefghij', 24, mono)).toEqual(['abcd', 'efgh', 'ij'])
    expect(wrapTextBox('', 100, mono)).toEqual([''])
  })
})

describe('isWinAnsi', () => {
  it('accepts Latin-1 and cp1252 extras, rejects other scripts', () => {
    expect(isWinAnsi('Café — “quoted” €5\nnext')).toBe(true)
    expect(isWinAnsi('مرحبا')).toBe(false)
    expect(isWinAnsi('你好')).toBe(false)
  })
})

const request = (drawings: DrawingInput[]): SavePdfRequest => ({
  path: '/tmp/test.pdf',
  markups: [],
  drawings,
  formValues: [],
  stamps: [],
})

const freetext = (over: Partial<Extract<DrawingInput, { kind: 'freetext' }>> = {}) =>
  ({
    kind: 'freetext',
    pageIndex: 0,
    rect: [100, 600, 300, 640],
    contents: 'Please check this total',
    fontSize: 12,
    color: [0.86, 0.22, 0.18],
    author: 'Omar',
    ...over,
  }) as DrawingInput

async function saveOne(d: DrawingInput, rotate = 0) {
  const src = await PDFDocument.create()
  src.addPage([612, 792]).setRotation(degrees(rotate))
  const { bytes } = await applySaveRequest(await src.save(), request([d]))
  const doc = await PDFDocument.load(bytes)
  const annots = doc.getPage(0).node.lookup(PDFName.of('Annots'), PDFArray)
  const annot = annots.lookup(0, PDFDict)
  const ap = annot.lookup(PDFName.of('AP'), PDFDict).lookup(PDFName.of('N'), PDFRawStream)
  const content = new TextDecoder().decode(decodePDFRawStream(ap).decode())
  return { bytes, annot, ap, content }
}

describe('saving a text-box comment', () => {
  it('writes a FreeText annotation with Helvetica text other viewers can read', async () => {
    const { bytes, annot, ap, content } = await saveOne(freetext())
    expect(annot.lookup(PDFName.of('Subtype'), PDFName).decodeText()).toBe('FreeText')
    expect(content).toContain('/Helv 12 Tf')
    expect(content).toContain('Tj')
    const fonts = ap.dict
      .lookup(PDFName.of('Resources'), PDFDict)
      .lookup(PDFName.of('Font'), PDFDict)
    expect(fonts.has(PDFName.of('Helv'))).toBe(true)
    const pdf = await getDocument({ data: bytes }).promise
    const [a] = await (await pdf.getPage(1)).getAnnotations()
    expect(a.annotationType).toBe(3)
    expect(a.contentsObj.str).toBe('Please check this total')
    expect(a.titleObj.str).toBe('Omar')
    expect(a.hasAppearance).toBe(true)
  })

  it('uses the rendered image for text Helvetica cannot encode', async () => {
    const png =
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
    const { ap, content } = await saveOne(freetext({ contents: 'مرحبا', image: png }))
    expect(content).toContain('/Im0 Do')
    expect(content).not.toContain('Tj')
    const res = ap.dict.lookup(PDFName.of('Resources'), PDFDict)
    expect(res.has(PDFName.of('XObject'))).toBe(true)
  })

  it('draws upright on a rotated page', async () => {
    const { annot, content } = await saveOne(freetext(), 90)
    expect(annot.lookup(PDFName.of('Rotate'))?.toString()).toBe('90')
    expect(content.startsWith('q 0 1 -1 0 200 0 cm')).toBe(true)
  })
})
