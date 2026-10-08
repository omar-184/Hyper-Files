import { describe, expect, it } from 'vitest'
import { PDFDocument, PDFName } from 'pdf-lib'
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs'
import { applySaveRequest } from '../src/main/save-pdf'
import type { OutlineEntryInput, SavePdfRequest } from '../src/shared/ipc'

const entry = (
  title: string,
  pageIndex: number | null,
  items: OutlineEntryInput[] = [],
  over: Partial<OutlineEntryInput> = {},
): OutlineEntryInput => ({ title, pageIndex, items, ...over })

async function save(outline: OutlineEntryInput[], extra: Partial<SavePdfRequest> = {}) {
  const src = await PDFDocument.create()
  for (let i = 0; i < 4; i++) src.addPage([300, 400])
  const req: SavePdfRequest = {
    path: '/tmp/t.pdf',
    markups: [],
    drawings: [],
    formValues: [],
    stamps: [],
    outline,
    ...extra,
  }
  return (await applySaveRequest(await src.save(), req)).bytes
}

interface Read {
  title: string
  page: number | null
  url?: string
  bold: boolean
  items: Read[]
}

async function readOutline(bytes: Uint8Array): Promise<Read[]> {
  const doc = await getDocument({ data: bytes }).promise
  const walk = async (
    nodes: { title: string; dest: unknown; url?: string; bold?: boolean; items: unknown[] }[],
  ): Promise<Read[]> =>
    Promise.all(
      nodes.map(async (n) => ({
        title: n.title,
        page: Array.isArray(n.dest)
          ? await doc.getPageIndex(n.dest[0] as { num: number; gen: number })
          : null,
        ...(n.url ? { url: n.url } : {}),
        bold: n.bold === true,
        items: await walk(n.items as typeof nodes),
      })),
    )
  return walk(((await doc.getOutline()) ?? []) as never)
}

describe('writeOutline', () => {
  it('writes a nested tree with page destinations, links and styles', async () => {
    const bytes = await save([
      entry('Intro', 0, [], { bold: true }),
      entry('Chapter 1', 1, [
        entry('Section 1.1', 2),
        entry('Site', null, [], { url: 'https://example.com/' }),
      ]),
      entry('Arabic عنوان', 3),
    ])
    expect(await readOutline(bytes)).toEqual([
      { title: 'Intro', page: 0, bold: true, items: [] },
      {
        title: 'Chapter 1',
        page: 1,
        bold: false,
        items: [
          { title: 'Section 1.1', page: 2, bold: false, items: [] },
          { title: 'Site', page: null, url: 'https://example.com/', bold: false, items: [] },
        ],
      },
      { title: 'Arabic عنوان', page: 3, bold: false, items: [] },
    ])
  })

  it('follows pages through a reorder and drops the destination of a deleted page', async () => {
    const bytes = await save([entry('Was first', 0), entry('Gone', 2)], {
      deletedPages: [2],
      pageOrder: [3, 1, 0],
    })
    const read = await readOutline(bytes)
    expect(read.map((r) => [r.title, r.page])).toEqual([
      ['Was first', 2],
      ['Gone', null],
    ])
  })

  it('removes the outline when given an empty tree', async () => {
    const withTree = await save([entry('A', 0)])
    const doc = await PDFDocument.load(withTree)
    expect(doc.catalog.get(PDFName.of('Outlines'))).toBeDefined()
    const req: SavePdfRequest = {
      path: '/tmp/t.pdf',
      markups: [],
      drawings: [],
      formValues: [],
      stamps: [],
      outline: [],
    }
    const cleared = (await applySaveRequest(withTree, req)).bytes
    expect((await PDFDocument.load(cleared)).catalog.get(PDFName.of('Outlines'))).toBeUndefined()
  })
})
