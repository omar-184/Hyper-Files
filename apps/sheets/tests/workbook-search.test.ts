import { describe, expect, it, vi } from 'vitest'

import { ERROR_VALUE_RE, selectWorkbookRange } from '../src/renderer/workbook-search'
import { ensureLazyRangeLoaded } from '../src/renderer/univer-sync'
import type { WorkbookReadContext } from '../src/renderer/workbook-readers'

vi.mock('../src/renderer/univer-sync', () => ({
  ensureLazyRangeLoaded: vi.fn().mockResolvedValue(true),
}))

describe('ERROR_VALUE_RE', () => {
  it('recognizes every error value Excel can display, modern data-type errors included', () => {
    for (const value of [
      '#REF!',
      '#DIV/0!',
      '#VALUE!',
      '#NAME?',
      '#N/A',
      '#NUM!',
      '#NULL!',
      '#SPILL!',
      '#CALC!',
      '#FIELD!',
      '#CONNECT!',
      '#BLOCKED!',
      '#UNKNOWN!',
      '#GETTING_DATA',
    ]) {
      expect(ERROR_VALUE_RE.test(value)).toBe(true)
    }
    for (const value of ['#FOO!', 'N/A', '#REF', '', '#12345!']) {
      expect(ERROR_VALUE_RE.test(value)).toBe(false)
    }
  })
})

function lazyState(journalCells: Map<string, Map<string, unknown>>) {
  return {
    file: {
      sessionId: 'session-1',
      sheets: [{ id: 'sh1', name: 'Data', rowCount: 4, columnCount: 2 }],
    },
    editJournal: { cells: journalCells, structuralOps: new Map() },
  }
}

describe('selectWorkbookRange', () => {
  function selectionCtx() {
    const range = { activate: vi.fn() }
    const worksheet = {
      getSheetId: () => 'sh1',
      getSheetName: () => 'Data',
      getRange: vi.fn().mockReturnValue(range),
      scrollToCell: vi.fn(),
    }
    const workbook = {
      getActiveSheet: () => worksheet,
      getSheetBySheetId: (id: string) => (id === 'sh1' ? worksheet : null),
      setActiveSheet: vi.fn(),
    }
    const ctx = {
      univerRef: { current: { univerAPI: { getActiveWorkbook: () => workbook } } },
      lazyWorkbookRef: { current: null },
      adapterRef: { current: { getSnapshot: () => ({ revision: 0, sheets: [] }) } },
    } as unknown as WorkbookReadContext
    return { ctx, worksheet, workbook, range }
  }

  it('selects and scrolls to the range on the active sheet', async () => {
    const { ctx, worksheet, range } = selectionCtx()
    const result = await selectWorkbookRange(
      ctx,
      undefined,
      { startRow: 1, startColumn: 1, endRow: 3, endColumn: 2 },
      () => {},
    )
    expect(result).toEqual({ ok: true, sheetName: 'Data' })
    expect(worksheet.getRange).toHaveBeenCalledWith(1, 1, 3, 2)
    expect(range.activate).toHaveBeenCalled()
    expect(worksheet.scrollToCell).toHaveBeenCalledWith(1, 1)
  })

  it('loads the target range first on lazy workbooks', async () => {
    const { ctx, worksheet } = selectionCtx()
    ;(ctx.lazyWorkbookRef as { current: unknown }).current = lazyState(new Map())
    const result = await selectWorkbookRange(
      ctx,
      'sh1',
      { startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 },
      () => {},
    )
    expect(result.ok).toBe(true)
    expect(vi.mocked(ensureLazyRangeLoaded)).toHaveBeenCalled()
    expect(worksheet.scrollToCell).toHaveBeenCalledWith(0, 0)
  })

  it('reports an unknown sheet', async () => {
    const { ctx } = selectionCtx()
    const result = await selectWorkbookRange(
      ctx,
      'ghost',
      { startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 },
      () => {},
    )
    expect(result).toEqual({ ok: false, error: 'Unknown sheet: ghost' })
  })
})
