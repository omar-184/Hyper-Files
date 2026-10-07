import { describe, expect, it } from 'vitest'

import { lazyGateFailure } from '../src/renderer/plan-operations'
import type { LazyWorkbookState } from '../src/renderer/univer-state'
import type { WorkbookOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'

/// The BeforeCommandExecute gates in App.tsx cancel these facade commands
/// silently; lazyGateFailure mirrors them so apply fails loud instead.

function lazyGateError(
  state: LazyWorkbookState,
  operation: WorkbookOperation,
): string | null {
  return lazyGateFailure(state, operation)?.reason ?? null
}

function lazyState(overrides: Record<string, unknown> = {}): LazyWorkbookState {
  return {
    file: {
      sessionId: 'session-1',
      sheets: [
        { id: 'sh1', name: 'Data', rowCount: 10, columnCount: 5, pivotRanges: [] },
        {
          id: 'sh2',
          name: 'Pivot',
          rowCount: 10,
          columnCount: 5,
          pivotRanges: [{ startRow: 0, endRow: 4, startColumn: 0, endColumn: 2 }],
        },
      ],
      visuals: [],
    },
    editJournal: {
      cells: new Map(),
      structuralOps: new Map(),
      sheets: { added: new Set(), removed: new Set() },
      visualAdds: [],
      tableAdds: [],
    },
    loadedRanges: new Map(),
    formulaMode: true,
    flags: { preloadComplete: true },
    filterOrigins: new Map(),
    appliedDvSheets: new Set(['sh1', 'sh2']),
    ...overrides,
  } as unknown as LazyWorkbookState
}

/// sh1 carries a file table over A1:D6 (0-based rows 0-5, columns 0-3).
function tableState(): LazyWorkbookState {
  const state = lazyState()
  ;(state.file.sheets[0] as { tables?: unknown[] }).tables = [
    {
      name: 'Form_Responses',
      range: { startRow: 0, endRow: 5, startColumn: 0, endColumn: 3 },
      headerRowCount: 1,
      showRowStripes: true,
      showColumnStripes: false,
    },
  ]
  return state
}

describe('lazyGateFailure', () => {
  it.each(['insert_rows', 'delete_rows'] as const)(
    'blocks %s on a pivot sheet and names it',
    (op) => {
      const state = lazyState()
      const error = lazyGateError(state, { op, sheetId: 'sh2', row: 6, count: 1 })
      expect(error).toContain('Pivot')
      expect(error).toContain('PivotTable')
      expect(lazyGateError(state, { op, sheetId: 'sh1', row: 6, count: 1 })).toBeNull()
    },
  )

  it('blocks merges on a pivot sheet regardless of the pivot region', () => {
    const error = lazyGateError(lazyState(), {
      op: 'merge_cells',
      sheetId: 'sh2',
      range: 'D8:E9',
    })
    expect(error).toContain('PivotTable')
  })

  it('blocks merges that overlap a table and names it for the ribbon', () => {
    const state = tableState()
    const gate = lazyGateFailure(state, { op: 'merge_cells', sheetId: 'sh1', range: 'B2:C3' })
    expect(gate?.reason).toContain('table "Form_Responses"')
    expect(gate?.messageKey).toBe('appMergeOverTable')
    // touching the table's last row/column still counts as inside
    expect(lazyGateError(state, { op: 'merge_cells', sheetId: 'sh1', range: 'D6:E7' })).toContain(
      'Form_Responses',
    )
    expect(lazyGateError(state, { op: 'merge_cells', sheetId: 'sh1', range: 'F8:G9' })).toBeNull()
    expect(lazyGateError(state, { op: 'merge_cells', sheetId: 'sh1', range: 'A7:D8' })).toBeNull()
    expect(lazyGateError(state, { op: 'unmerge_cells', sheetId: 'sh1', range: 'B2:C3' })).toBeNull()
  })

  it('checks file tables at their post-structural-edit coordinates', () => {
    const state = tableState()
    state.editJournal.structuralOps.set('sh1', [{ kind: 'insert-rows', index: 0, count: 2 }])
    // the table now sits on A3:D8: its old top rows are plain cells
    expect(lazyGateError(state, { op: 'merge_cells', sheetId: 'sh1', range: 'A1:D2' })).toBeNull()
    expect(lazyGateError(state, { op: 'merge_cells', sheetId: 'sh1', range: 'B7:C8' })).toContain(
      'Form_Responses',
    )
    expect(lazyGateError(state, { op: 'merge_cells', sheetId: 'sh1', range: 'A9:D9' })).toBeNull()
  })

  it('keeps rows inserted inside a table in the gate after the original lines are deleted', () => {
    const state = tableState()
    // insert 2 rows inside the table (rows 2-3), then delete the original rows
    // 2-5 that now sit at 4-7: the table survives on A1:D4 with the inserted rows
    state.editJournal.structuralOps.set('sh1', [
      { kind: 'insert-rows', index: 2, count: 2 },
      { kind: 'remove-rows', index: 4, count: 4 },
    ])
    expect(lazyGateError(state, { op: 'merge_cells', sheetId: 'sh1', range: 'B3:C4' })).toContain(
      'Form_Responses',
    )
    expect(lazyGateError(state, { op: 'merge_cells', sheetId: 'sh1', range: 'A5:D5' })).toBeNull()
  })

  it('blocks merges over a table added this session', () => {
    const state = lazyState({
      editJournal: {
        cells: new Map(),
        structuralOps: new Map(),
        sheets: { added: new Set(), removed: new Set() },
        visualAdds: [],
        tableAdds: [
          {
            sheetId: 'sh1',
            name: 'Added1',
            area: { startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 },
            columnNames: ['a', 'b'],
          },
        ],
      },
    })
    expect(lazyGateError(state, { op: 'merge_cells', sheetId: 'sh1', range: 'A1:B1' })).toContain(
      'Added1',
    )
    expect(lazyGateError(state, { op: 'merge_cells', sheetId: 'sh1', range: 'C1:D1' })).toBeNull()
  })

  it('blocks filter ops until the workbook is fully loaded', () => {
    const loading = lazyState({ flags: { preloadComplete: false } })
    expect(lazyGateError(loading, { op: 'set_filter', sheetId: 'sh1', range: 'A1:B5' })).toContain(
      'still loading',
    )
    const streamed = lazyState({ formulaMode: false, flags: { preloadComplete: false } })
    expect(lazyGateError(streamed, { op: 'clear_filter', sheetId: 'sh1' })).toContain(
      'fully-loaded mode',
    )
    expect(
      lazyGateError(lazyState(), {
        op: 'set_filter_criteria',
        sheetId: 'sh1',
        column: 'A',
        values: null,
      }),
    ).toBeNull()
  })

  it('blocks writes on a sheet whose XML is above the save-patch cap', () => {
    const state = lazyState({
      file: {
        sessionId: 'session-1',
        sheets: [
          {
            id: 'sh1',
            name: 'Huge',
            rowCount: 10,
            columnCount: 5,
            pivotRanges: [],
            sourceXmlBytes: 501 * 1024 * 1024,
          },
        ],
        visuals: [],
      },
    })
    const error = lazyGateError(state, { op: 'set_cell', sheetId: 'sh1', address: 'A1', value: 1 })
    expect(error).toContain('read-only')
    expect(error).toContain('500MB')
    // workbook.xml-only ops stay allowed
    expect(lazyGateError(state, { op: 'rename_sheet', sheetId: 'sh1', name: 'X' })).toBeNull()
    expect(
      lazyGateError(state, { op: 'set_sheet_hidden', sheetId: 'sh1', hidden: true }),
    ).toBeNull()
    // at/below the cap: no gate
    const under = lazyState()
    expect(
      lazyGateError(under, { op: 'set_cell', sheetId: 'sh1', address: 'A1', value: 1 }),
    ).toBeNull()
  })

  it('gates add_pivot on its output sheet, not its source', () => {
    const state = lazyState({
      file: {
        sessionId: 'session-1',
        sheets: [
          { id: 'small', name: 'Small', rowCount: 10, columnCount: 5, pivotRanges: [] },
          {
            id: 'huge',
            name: 'Huge',
            rowCount: 10,
            columnCount: 5,
            pivotRanges: [],
            sourceXmlBytes: 501 * 1024 * 1024,
          },
        ],
        visuals: [],
      },
    })
    const pivot = (sheetId: string, targetSheetId?: string): WorkbookOperation => ({
      op: 'add_pivot',
      sheetId,
      ...(targetSheetId === undefined ? {} : { targetSheetId }),
      sourceRange: 'A1:B5',
      targetCell: 'D1',
      rowFields: 'h',
      values: [{ field: 'h', agg: 'count' }],
    })
    // small source → oversized output: blocked
    expect(lazyGateError(state, pivot('small', 'huge'))).toContain('read-only')
    // oversized source read → small output: allowed
    expect(lazyGateError(state, pivot('huge', 'small'))).toBeNull()
    // default output = the source sheet
    expect(lazyGateError(state, pivot('huge'))).toContain('read-only')
  })

  it('allows filter ops on sheets added this session even while streaming', () => {
    const state = lazyState({
      formulaMode: false,
      flags: { preloadComplete: false },
      editJournal: {
        cells: new Map(),
        structuralOps: new Map(),
        sheets: { added: new Set(['new1']), removed: new Set() },
        visualAdds: [],
        tableAdds: [],
      },
    })
    expect(lazyGateError(state, { op: 'set_filter', sheetId: 'new1', range: 'A1:B5' })).toBeNull()
  })

  it('blocks filter edits on table-owned filters', () => {
    const state = lazyState({
      filterOrigins: new Map([
        [
          'sh1',
          { origin: 'table', range: { startRow: 0, endRow: 5, startColumn: 0, endColumn: 3 } },
        ],
      ]),
    })
    expect(
      lazyGateError(state, {
        op: 'set_filter_criteria',
        sheetId: 'sh1',
        column: 'A',
        values: ['x'],
      }),
    ).toContain('table')
  })

  it('blocks set_data_validation until the sheet is indexed', () => {
    const state = lazyState({ appliedDvSheets: new Set() })
    expect(
      lazyGateError(state, {
        op: 'set_data_validation',
        sheetId: 'sh1',
        range: 'A1:A5',
        validation: null,
      }),
    ).toContain('still being indexed')
    expect(
      lazyGateError(lazyState(), {
        op: 'set_data_validation',
        sheetId: 'sh1',
        range: 'A1:A5',
        validation: null,
      }),
    ).toBeNull()
  })
})
