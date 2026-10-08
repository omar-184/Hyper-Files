import { describe, expect, it } from 'vitest'
import { createEditJournal, recordNeutralStyleEdit } from '../src/renderer/edit-journal'
import {
  cellIsLocked,
  matrixHasLockedCell,
  PROTECTION_CHECK_BUDGET,
  protectedPlanFailure,
  rangeHasLockedCell,
  sheetIsProtected,
} from '../src/renderer/sheet-protection'
import type { LazyWorkbookState } from '../src/renderer/univer-state'

const area = (startRow: number, startColumn: number, endRow: number, endColumn: number) => ({
  startRow,
  startColumn,
  endRow,
  endColumn,
})

// File locks: B2 is unlocked; column D is unlocked except D1.
function makeState(
  options: {
    fileProtected?: boolean
    ranges?: { name: string; sqref: string; hasPassword: boolean }[]
    truncated?: boolean
  } = {},
): LazyWorkbookState {
  return {
    editJournal: createEditJournal(),
    sheetProtections: new Map([
      ['s1', { protected: options.fileProtected ?? true, hasPassword: false }],
    ]),
    sheetProtectedRanges: new Map([['s1', options.ranges ?? []]]),
    sheetCellLocks: new Map([
      [
        's1',
        {
          unlocked: [area(1, 1, 1, 1), area(0, 3, 1_048_575, 3)],
          locked: [area(0, 3, 0, 3)],
          truncated: options.truncated ?? false,
        },
      ],
    ]),
  } as unknown as LazyWorkbookState
}

describe('sheet protection', () => {
  it('follows the session toggle over the file state', () => {
    const state = makeState({ fileProtected: true })
    expect(sheetIsProtected(state, 's1')).toBe(true)
    state.editJournal.sheetProtection.set('s1', false)
    expect(sheetIsProtected(state, 's1')).toBe(false)
    expect(sheetIsProtected(state, 'unknown')).toBe(false)
    expect(sheetIsProtected(null, 's1')).toBe(false)
  })

  it('locks every cell the file does not unlock', () => {
    const state = makeState()
    expect(cellIsLocked(state, 's1', 0, 0)).toBe(true)
    expect(cellIsLocked(state, 's1', 1, 1)).toBe(false)
    expect(cellIsLocked(state, 's1', 500, 3)).toBe(false)
    // a locked xf inside the unlocked column
    expect(cellIsLocked(state, 's1', 0, 3)).toBe(true)
    // nothing known yet for this sheet
    expect(cellIsLocked(state, 's2', 1, 1)).toBe(true)
  })

  it('maps screen cells back to file cells across inserted rows', () => {
    const state = makeState()
    state.editJournal.structuralOps.set('s1', [{ kind: 'insert-rows', index: 0, count: 2 }])
    // file B2 now sits at screen row 3; the inserted rows lock like new cells
    expect(cellIsLocked(state, 's1', 3, 1)).toBe(false)
    expect(cellIsLocked(state, 's1', 1, 1)).toBe(true)
  })

  it('honors Format Cells lock changes and Clear Formats from this session', () => {
    const state = makeState()
    recordNeutralStyleEdit(state.editJournal, 's1', 0, 0, { protectionLocked: false })
    recordNeutralStyleEdit(state.editJournal, 's1', 1, 1, { protectionLocked: true })
    expect(cellIsLocked(state, 's1', 0, 0)).toBe(false)
    expect(cellIsLocked(state, 's1', 1, 1)).toBe(true)
    state.editJournal.cells
      .get('s1')
      ?.set('5:3', { row: 5, column: 3, hasValue: true, value: null, styleReset: true })
    expect(cellIsLocked(state, 's1', 5, 3)).toBe(true)
  })

  it('opens allow-edit ranges without a password only', () => {
    const state = makeState({
      ranges: [
        { name: 'Inputs', sqref: '$F$3:$G$4 H1', hasPassword: false },
        { name: 'Secret', sqref: 'J1:J9', hasPassword: true },
      ],
    })
    expect(cellIsLocked(state, 's1', 2, 5)).toBe(false)
    expect(cellIsLocked(state, 's1', 3, 6)).toBe(false)
    expect(cellIsLocked(state, 's1', 0, 7)).toBe(false)
    expect(cellIsLocked(state, 's1', 4, 6)).toBe(true)
    expect(cellIsLocked(state, 's1', 0, 9)).toBe(true)
  })

  it('stops enforcing cell locks when the file had too many areas', () => {
    const state = makeState({ truncated: true })
    expect(cellIsLocked(state, 's1', 0, 0)).toBe(false)
  })

  it('checks ranges and mutation matrices cell by cell', () => {
    const state = makeState()
    expect(rangeHasLockedCell(state, 's1', [area(1, 1, 1, 1)])).toBe(false)
    expect(rangeHasLockedCell(state, 's1', [area(1, 1, 1, 2)])).toBe(true)
    expect(matrixHasLockedCell(state, 's1', { 1: { 1: { v: 'x' } } })).toBe(false)
    expect(matrixHasLockedCell(state, 's1', { 1: { 1: {}, 2: {} } })).toBe(true)
  })

  it('treats a range too large to check as locked', () => {
    const state = makeState()
    expect(rangeHasLockedCell(state, 's1', [area(1, 3, PROTECTION_CHECK_BUDGET, 3)])).toBe(true)
  })

  it('fails plans that reshape or write locked cells of a protected sheet', () => {
    const state = makeState()
    expect(protectedPlanFailure(state, [{ op: 'insert_rows', sheetId: 's1' } as never], [])).toBe(
      'appSheetProtectedStructure',
    )
    expect(
      protectedPlanFailure(
        state,
        [{ op: 'clear_range', sheetId: 's1', range: 'A1:B2' } as never],
        [],
      ),
    ).toBe('appCellProtected')
    expect(protectedPlanFailure(state, [], [{ sheetId: 's1', address: 'B2' }])).toBe(null)
    expect(protectedPlanFailure(state, [], [{ sheetId: 's1', address: 'A1' }])).toBe(
      'appCellProtected',
    )
    // Unprotecting in the same batch defers to the per-command gate.
    expect(
      protectedPlanFailure(
        state,
        [{ op: 'protect_sheet', sheetId: 's1', protected: false } as never],
        [{ sheetId: 's1', address: 'A1' }],
      ),
    ).toBe(null)
  })

  it('leaves unprotected sheets alone', () => {
    const state = makeState({ fileProtected: false })
    expect(
      protectedPlanFailure(
        state,
        [{ op: 'delete_cols', sheetId: 's1' } as never],
        [{ sheetId: 's1', address: 'A1' }],
      ),
    ).toBe(null)
  })
})
