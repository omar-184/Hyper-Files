import { describe, expect, it } from 'vitest'
import { PDFDocument } from 'pdf-lib'
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs'
import { applySaveRequest } from '../src/main/save-pdf'
import { normalizeLinkUrl } from '../src/renderer/LinkDialog'
import type { DrawingInput, SavePdfRequest } from '../src/shared/ipc'

describe('normalizeLinkUrl', () => {
  it('completes bare domains and addresses, keeps web and mail links, rejects the rest', () => {
    expect(normalizeLinkUrl('example.com/page')).toBe('https://example.com/page')
    expect(normalizeLinkUrl(' http://example.com ')).toBe('http://example.com/')
    expect(normalizeLinkUrl('omar@example.com')).toBe('mailto:omar@example.com')
    expect(normalizeLinkUrl('mailto:a@b.co')).toBe('mailto:a@b.co')
    expect(normalizeLinkUrl('javascript:alert(1)')).toBeNull()
    expect(normalizeLinkUrl('file:///etc/passwd')).toBeNull()
    expect(normalizeLinkUrl('two words')).toBeNull()
    expect(normalizeLinkUrl('')).toBeNull()
  })
})

describe('link save', () => {
  it('writes web and page links; a link to a deleted page is dropped', async () => {
    const src = await PDFDocument.create()
    for (let i = 0; i < 3; i++) src.addPage([300, 400])
    const links: DrawingInput[] = [
      { kind: 'link', pageIndex: 0, rect: [10, 10, 100, 30], url: 'https://example.com/' },
      { kind: 'link', pageIndex: 0, rect: [10, 50, 100, 70], targetPage: 2 },
      { kind: 'link', pageIndex: 0, rect: [10, 90, 100, 110], targetPage: 1 },
    ]
    const req: SavePdfRequest = {
      path: '/tmp/t.pdf',
      markups: [],
      drawings: links,
      formValues: [],
      stamps: [],
      deletedPages: [1],
    }
    const { bytes } = await applySaveRequest(await src.save(), req)
    const doc = await getDocument({ data: bytes }).promise
    const annots = (await (await doc.getPage(1)).getAnnotations()) as {
      subtype: string
      url?: string
      dest?: unknown[]
    }[]
    const found = await Promise.all(
      annots
        .filter((a) => a.subtype === 'Link')
        .map(
          async (a) =>
            a.url ?? (await doc.getPageIndex(a.dest![0] as { num: number; gen: number })),
        ),
    )
    // Page 2 (index 2) is index 1 after page 1 is deleted
    expect(found).toEqual(['https://example.com/', 1])
  })
})
