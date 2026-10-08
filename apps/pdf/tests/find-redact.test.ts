import { describe, expect, it } from 'vitest'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import { PDFDocument, StandardFonts } from 'pdf-lib'
import { buildSearchIndex, findInIndex, type SearchIndex } from '../src/renderer/search'
import { marksFromMatches, padMatchRect } from '../src/renderer/redact-marks'
import { MAX_REDACTION_REGIONS } from '../src/shared/ipc'
import { luhnValid, type RedactPattern } from '../src/shared/text-match'
import { findTextBoxes, lineRects } from '../src/main/find-text'

/** One page per string, one 600-pt-wide item per line so char ratios are easy to read */
async function indexOf(...pages: string[][]): Promise<SearchIndex> {
  const doc = {
    numPages: pages.length,
    getPage: async (n: number) => ({
      getTextContent: async () => ({
        items: pages[n - 1]!.map((line, i) => ({
          str: line,
          transform: [1, 0, 0, 1, 0, 700 - i * 20],
          width: line.length * 6,
          height: 10,
          hasEOL: true,
        })),
      }),
    }),
  } as unknown as PDFDocumentProxy
  return buildSearchIndex(doc)
}

const found = (index: SearchIndex, pattern: RedactPattern) =>
  findInIndex(index, { pattern }).map((m) => {
    const entry = index[m.pageIndex]!
    const item = entry.items.find((it) => Math.abs(it.y - m.rects[0]![1]) < 0.01)!
    const start = item.start + Math.round(m.rects[0]![0] / 6)
    const end = item.start + Math.round(m.rects[m.rects.length - 1]![2] / 6)
    return entry.text.slice(start, end)
  })

describe('findInIndex with patterns', () => {
  it('finds email addresses', async () => {
    const index = await indexOf(['Write to jane.doe@example.co.uk or ops@host.io today'])
    expect(found(index, 'email')).toEqual(['jane.doe@example.co.uk', 'ops@host.io'])
  })

  it('finds phone numbers in common layouts and skips dates and short numbers', async () => {
    const index = await indexOf([
      'Call +1 (555) 123-4567 or 020 7946 0958.',
      'Office 555.867.5309, ext 12, invoice date 2024-01-02, room 42',
    ])
    expect(found(index, 'phone')).toEqual(['+1 (555) 123-4567', '020 7946 0958', '555.867.5309'])
  })

  it('does not join digit groups across a line break', async () => {
    const index = await indexOf(['Total 1200', '3400 units'])
    expect(found(index, 'phone')).toEqual([])
  })

  it('finds card numbers that pass the Luhn check only', async () => {
    const index = await indexOf(['Paid with 4111 1111 1111 1111; ref 4111 1111 1111 1112'])
    expect(found(index, 'card')).toEqual(['4111 1111 1111 1111'])
  })

  it('reports matches on their own pages', async () => {
    const index = await indexOf(['none here'], ['mail a@b.org'])
    expect(findInIndex(index, { pattern: 'email' }).map((m) => m.pageIndex)).toEqual([1])
  })
})

describe('luhnValid', () => {
  it('accepts valid and rejects altered numbers', () => {
    expect(luhnValid('4111111111111111')).toBe(true)
    expect(luhnValid('5500005555555559')).toBe(true)
    expect(luhnValid('4111111111111112')).toBe(false)
  })
})

describe('marksFromMatches', () => {
  const match = (pageIndex: number, x: number) => ({
    pageIndex,
    rects: [[x, 100, x + 30, 110] as [number, number, number, number]],
  })

  it('pads each rect to cover descenders and accents', () => {
    expect(padMatchRect([10, 100, 40, 110])).toEqual([9.5, 97, 40.5, 111])
  })

  it('adds one mark per rect and skips rects already marked', () => {
    const first = marksFromMatches([match(0, 10), match(1, 50)], [])
    expect(first.added).toHaveLength(2)
    expect(first.overCap).toBe(0)
    const again = marksFromMatches([match(0, 10), match(1, 50), match(1, 200)], first.added)
    expect(again.added).toEqual([{ pageIndex: 1, rect: [200, 100, 230, 110] }])
  })

  it('stops at the apply cap and counts what it skipped', () => {
    const many = Array.from({ length: MAX_REDACTION_REGIONS + 5 }, (_, i) => match(i, 10))
    const result = marksFromMatches(many, [])
    expect(result.added).toHaveLength(MAX_REDACTION_REGIONS)
    expect(result.overCap).toBe(5)
  })
})

describe('findTextBoxes', () => {
  it('covers every glyph of a match with the exact line box', async () => {
    const pdf = await PDFDocument.create()
    const page = pdf.addPage([300, 200])
    const font = await pdf.embedFont(StandardFonts.Helvetica)
    page.drawText('Contact jane@example.com today', { x: 20, y: 150, size: 12, font })
    page.drawText('second jane@example.com', { x: 20, y: 60, size: 12, font })
    const matches = await findTextBoxes(await pdf.save(), { pattern: 'email' })
    expect(matches).toHaveLength(2)
    const [first] = matches
    expect(first!.pageIndex).toBe(0)
    expect(first!.rects).toHaveLength(1)
    const [x1, y1, x2, y2] = first!.rects[0]!
    const start = 20 + font.widthOfTextAtSize('Contact ', 12)
    const end = 20 + font.widthOfTextAtSize('Contact jane@example.com', 12)
    expect(x1).toBeLessThanOrEqual(start + 0.5)
    expect(x1).toBeGreaterThan(start - 3)
    expect(x2).toBeGreaterThanOrEqual(end - 0.5)
    expect(x2).toBeLessThan(end + 3)
    // Loose boxes span the font's descent to ascent around the baseline
    expect(y1).toBeLessThan(150)
    expect(y2).toBeGreaterThan(158)
  })

  it('finds a typed phrase case-insensitively', async () => {
    const pdf = await PDFDocument.create()
    const page = pdf.addPage([300, 200])
    const font = await pdf.embedFont(StandardFonts.Helvetica)
    page.drawText('Top SECRET plan', { x: 20, y: 150, size: 12, font })
    const matches = await findTextBoxes(await pdf.save(), { query: 'secret' })
    expect(matches).toHaveLength(1)
  })
})

describe('lineRects', () => {
  it('unions boxes on a line and splits at a line change', () => {
    expect(
      lineRects([
        [0, 100, 5, 110],
        [5, 99, 10, 111],
        [0, 80, 5, 90],
      ]),
    ).toEqual([
      [0, 99, 10, 111],
      [0, 80, 5, 90],
    ])
  })
})
