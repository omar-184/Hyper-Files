import { describe, expect, it } from 'vitest'

import {
  expandToPrimitiveOps,
  filteredCopySourceRows,
  matchableCellText,
  workbookOperationSchema,
  type WorkbookOperation,
} from '@genoffice/xlsx-gateway/domain/workbook-dsl'

/// copy_range filterColumn/filterValues: row extraction for splitting data
/// by a column's values — matching rows land compacted at the target as
/// static values. The executors share filteredCopySourceRows /
/// matchableCellText, tested here; the lazy chunked wiring is exercised by
/// the real-app driver.

describe('copy_range filter schema and geometry validation', () => {
  const base = {
    op: 'copy_range' as const,
    sheetId: 'sheet-1',
    source: 'A1:D6',
    target: 'F1',
  }

  it('parses filterColumn/filterValues', () => {
    const parsed = workbookOperationSchema.parse({
      ...base,
      filterColumn: 'D',
      filterValues: ['ja', 'ko'],
    })
    expect(parsed).toMatchObject({ filterColumn: 'D', filterValues: ['ja', 'ko'] })
  })

  it('accepts filter values up to the cell text cap (real cell contents are legal)', () => {
    const longValue = 'x'.repeat(300)
    const maxValue = 'y'.repeat(32_767)
    const parsed = workbookOperationSchema.parse({
      ...base,
      filterColumn: 'D',
      filterValues: [longValue, maxValue],
    })
    expect(parsed).toMatchObject({ filterValues: [longValue, maxValue] })
    expect(() =>
      workbookOperationSchema.parse({
        ...base,
        filterColumn: 'D',
        filterValues: ['z'.repeat(32_768)],
      }),
    ).toThrow()
  })

  it('matches a long filter value against the cell text (trimmed, case-insensitive)', () => {
    const longText = `Note: ${'x'.repeat(300)}`
    const rows = filteredCopySourceRows(
      {
        op: 'copy_range',
        sheetId: 'sheet-1',
        source: 'A1:D6',
        target: 'F1',
        filterColumn: 'D',
        filterValues: [` ${longText.toUpperCase()} `],
      },
      (row, column) => (column === 3 && row === 2 ? matchableCellText(longText) : ''),
    )
    expect(rows).toEqual([2])
  })

  it('rejects filterValues without filterColumn (and vice versa)', () => {
    expect(() =>
      expandToPrimitiveOps([{ ...base, filterValues: ['ja'] } as WorkbookOperation]),
    ).toThrow(/provided together/)
    expect(() =>
      expandToPrimitiveOps([{ ...base, filterColumn: 'D' } as WorkbookOperation]),
    ).toThrow(/provided together/)
  })

  it('rejects a filterColumn outside the source range', () => {
    expect(() =>
      expandToPrimitiveOps([
        { ...base, filterColumn: 'E', filterValues: ['ja'] } as WorkbookOperation,
      ]),
    ).toThrow(/outside the source range/)
  })

  it('rejects a multi-cell target on a filtered copy', () => {
    expect(() =>
      expandToPrimitiveOps([
        {
          ...base,
          target: 'F1:I6',
          filterColumn: 'D',
          filterValues: ['ja'],
        } as WorkbookOperation,
      ]),
    ).toThrow(/top-left cell/)
  })
})

describe('matchableCellText', () => {
  it('trims and lowercases text', () => {
    expect(matchableCellText('  JA ')).toBe('ja')
  })
  it('coerces numbers and booleans, and maps null to ""', () => {
    expect(matchableCellText(12.5)).toBe('12.5')
    expect(matchableCellText(true)).toBe('true')
    expect(matchableCellText(false)).toBe('false')
    expect(matchableCellText(null)).toBe('')
  })
})

describe('filteredCopySourceRows', () => {
  const values = new Map<string, string>([
    ['1:3', 'locale'],
    ['2:3', 'ja'],
    ['3:3', 'ko'],
    ['4:3', ' JA '],
    ['5:3', 'en'],
  ])
  const cellText = (row: number, column: number): string =>
    matchableCellText(values.get(`${row}:${column}`) ?? null)

  it('returns every source row when unfiltered', () => {
    const op = workbookOperationSchema.parse({
      op: 'copy_range',
      sheetId: 's1',
      source: 'A2:D4',
      target: 'F1',
    })
    expect(filteredCopySourceRows(op as never, cellText)).toEqual([1, 2, 3])
  })

  it('keeps only matching rows, trimmed and case-insensitive', () => {
    const op = workbookOperationSchema.parse({
      op: 'copy_range',
      sheetId: 's1',
      source: 'A1:D6',
      target: 'F1',
      filterColumn: 'D',
      filterValues: ['ja'],
    })
    // Screen rows 2 and 4 hold "ja" / " JA " (the header row does not match).
    expect(filteredCopySourceRows(op as never, cellText)).toEqual([2, 4])
  })

  it('returns null when nothing matches', () => {
    const op = workbookOperationSchema.parse({
      op: 'copy_range',
      sheetId: 's1',
      source: 'A1:D6',
      target: 'F1',
      filterColumn: 'D',
      filterValues: ['zh-CN'],
    })
    expect(filteredCopySourceRows(op as never, cellText)).toBeNull()
  })
})

describe('add_sheet rows/columns', () => {
  it('parses the optional grid size', () => {
    expect(
      workbookOperationSchema.parse({ op: 'add_sheet', name: 'ja', rows: 6700, columns: 8 }),
    ).toMatchObject({ rows: 6700, columns: 8 })
    expect(() => workbookOperationSchema.parse({ op: 'add_sheet', name: 'ja', rows: 0 })).toThrow()
    expect(() =>
      workbookOperationSchema.parse({ op: 'add_sheet', name: 'ja', rows: 2_000_000 }),
    ).toThrow()
  })
})
