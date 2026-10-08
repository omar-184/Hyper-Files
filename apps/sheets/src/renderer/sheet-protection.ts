/**
 * Protect Sheet enforcement: which cells of a protected sheet may change.
 * Excel locks every cell by default; a cell stays editable when the file
 * unlocks it (the sidecar reports those areas once the sheet is indexed),
 * when Format Cells unlocked it this session, or when it sits in an
 * allow-edit range without a password.
 */
import { parseRange, type RangeBounds } from '@genoffice/xlsx-gateway/domain/cell-address'
import { type LazyWorkbookState } from './univer-state'
import { screenToFile } from './view-transform'

/// Above this much work (cells x lock areas) a check gives up and treats the
/// range as locked rather than stall the UI on a whole-column selection.
export const PROTECTION_CHECK_BUDGET = 2_000_000

/// Effective protection: the session's toggle, else the file's state.
export function sheetIsProtected(state: LazyWorkbookState | null, sheetId: string): boolean {
  if (!state) return false
  return (
    state.editJournal.sheetProtection.get(sheetId) ??
    state.sheetProtections.get(sheetId)?.protected ??
    false
  )
}

const parsedRanges = new WeakMap<object, RangeBounds[]>()

/// Allow-edit ranges without a password; password ones cannot be unlocked
/// here, so their cells keep their own lock.
function openEditAreas(state: LazyWorkbookState, sheetId: string): RangeBounds[] {
  const ranges = state.sheetProtectedRanges.get(sheetId)
  if (!ranges) return []
  let areas = parsedRanges.get(ranges)
  if (!areas) {
    areas = []
    for (const range of ranges) {
      if (range.hasPassword) continue
      for (const part of range.sqref.split(/\s+/)) {
        if (!part) continue
        try {
          areas.push(parseRange(part.replaceAll('$', '')))
        } catch {
          // An unparsable area grants nothing.
        }
      }
    }
    parsedRanges.set(ranges, areas)
  }
  return areas
}

function inAreas(areas: readonly RangeBounds[], row: number, column: number): boolean {
  return areas.some(
    (area) =>
      row >= area.startRow &&
      row <= area.endRow &&
      column >= area.startColumn &&
      column <= area.endColumn,
  )
}

/// The file's own lock for a screen cell. Lines inserted this session have
/// no file style and lock like a fresh cell; a sheet whose lock areas were
/// too many to report stays editable rather than block real input.
function fileCellLocked(
  state: LazyWorkbookState,
  sheetId: string,
  row: number,
  column: number,
): boolean {
  const locks = state.sheetCellLocks.get(sheetId)
  if (!locks) return true
  if (locks.truncated) return false
  const ops = state.editJournal.structuralOps.get(sheetId) ?? []
  const fileRow = ops.length === 0 ? row : screenToFile(ops, 'row', row)
  const fileColumn = ops.length === 0 ? column : screenToFile(ops, 'column', column)
  if (fileRow === null || fileColumn === null) return true
  if (inAreas(locks.locked, fileRow, fileColumn)) return true
  return !inAreas(locks.unlocked, fileRow, fileColumn)
}

export function cellIsLocked(
  state: LazyWorkbookState,
  sheetId: string,
  row: number,
  column: number,
): boolean {
  if (inAreas(openEditAreas(state, sheetId), row, column)) return false
  const entry = state.editJournal.cells.get(sheetId)?.get(`${row}:${column}`)
  const override = entry?.style?.protectionLocked
  if (override !== undefined) return override
  // Clear Formats resets the cell to the default xf, which is locked.
  if (entry?.styleReset) return true
  return fileCellLocked(state, sheetId, row, column)
}

function checkCost(state: LazyWorkbookState, sheetId: string, cells: number): number {
  const locks = state.sheetCellLocks.get(sheetId)
  const areas =
    1 +
    openEditAreas(state, sheetId).length +
    (locks ? locks.unlocked.length + locks.locked.length : 0)
  return cells * areas
}

/// True when any cell of the areas is locked (or the areas are too large to
/// check). Callers check sheetIsProtected first.
export function rangeHasLockedCell(
  state: LazyWorkbookState,
  sheetId: string,
  areas: readonly RangeBounds[],
): boolean {
  let total = 0
  for (const area of areas) {
    total += (area.endRow - area.startRow + 1) * (area.endColumn - area.startColumn + 1)
  }
  if (checkCost(state, sheetId, total) > PROTECTION_CHECK_BUDGET) return true
  for (const area of areas) {
    for (let row = area.startRow; row <= area.endRow; row += 1) {
      for (let column = area.startColumn; column <= area.endColumn; column += 1) {
        if (cellIsLocked(state, sheetId, row, column)) return true
      }
    }
  }
  return false
}

/// Cells named by a set-range-values mutation's sparse matrix ({row: {col}}).
export function matrixHasLockedCell(
  state: LazyWorkbookState,
  sheetId: string,
  matrix: unknown,
): boolean {
  if (!matrix || typeof matrix !== 'object') return false
  const budget = PROTECTION_CHECK_BUDGET / checkCost(state, sheetId, 1)
  let checked = 0
  for (const [rowKey, columns] of Object.entries(matrix as Record<string, unknown>)) {
    if (!columns || typeof columns !== 'object') continue
    const row = Number(rowKey)
    for (const columnKey of Object.keys(columns as Record<string, unknown>)) {
      checked += 1
      if (checked > budget) return true
      if (cellIsLocked(state, sheetId, row, Number(columnKey))) return true
    }
  }
  return false
}

/// Plan ops that reshape the grid; Excel blocks them on a protected sheet
/// unless the protection explicitly allows them.
export const PROTECTED_STRUCTURE_OPS = new Set([
  'insert_rows',
  'delete_rows',
  'insert_cols',
  'delete_cols',
  'merge_cells',
  'unmerge_cells',
  'sort_range',
  'add_table',
  'add_table_row',
  'add_table_column',
  'delete_table_row',
  'delete_table_column',
])

/// Plan ops that write cells, keyed to the field naming their target range.
export const PROTECTED_RANGE_OPS: Readonly<Record<string, string>> = {
  clear_range: 'range',
  fill_range: 'target',
  copy_range: 'target',
  convert_to_values: 'range',
  format_range: 'range',
  find_replace: 'range',
}

/// Univer commands that reshape the grid of the active sheet.
export const PROTECTED_STRUCTURE_COMMANDS = new Set([
  'sheet.command.insert-row',
  'sheet.command.insert-row-before',
  'sheet.command.insert-row-after',
  'sheet.command.insert-multi-rows-above',
  'sheet.command.insert-multi-rows-after',
  'sheet.command.insert-col',
  'sheet.command.insert-col-before',
  'sheet.command.insert-col-after',
  'sheet.command.insert-multi-cols-before',
  'sheet.command.insert-multi-cols-right',
  'sheet.command.remove-row',
  'sheet.command.remove-col',
  'sheet.command.remove-row-confirm',
  'sheet.command.remove-col-confirm',
  'sheet.command.insert-range-move-down',
  'sheet.command.insert-range-move-right',
  'sheet.command.delete-range-move-left',
  'sheet.command.delete-range-move-up',
  'sheet.command.move-rows',
  'sheet.command.move-cols',
  'sheet.command.add-worksheet-merge',
  'sheet.command.add-worksheet-merge-all',
  'sheet.command.add-worksheet-merge-vertical',
  'sheet.command.add-worksheet-merge-horizontal',
  'sheet.command.remove-worksheet-merge',
  'sheet.command.insert-row-by-range',
  'sheet.command.insert-col-by-range',
  'sheet.command.remove-row-by-range',
  'sheet.command.remove-col-by-range',
  'sheet.command.insert-range-move-down-confirm',
  'sheet.command.insert-range-move-right-confirm',
  'sheet.command.delete-range-move-left-confirm',
  'sheet.command.delete-range-move-up-confirm',
  'sheet.command.move-range',
  'sheet.command.sort-range',
  'sheet.command.reorder-range',
])

/// Whole-batch check for a change plan, run before anything applies so a
/// blocked edit fails loud instead of being silently cancelled midway.
/// A batch that toggles a sheet's protection is left to the per-command gate.
export function protectedPlanFailure(
  state: LazyWorkbookState,
  ops: readonly { readonly op: string }[],
  cellChanges: readonly { readonly sheetId: string; readonly address: string }[],
): 'appCellProtected' | 'appSheetProtectedStructure' | null {
  const toggled = new Set<string>()
  for (const op of ops) {
    const sheetId = (op as { sheetId?: unknown }).sheetId
    if (op.op === 'protect_sheet' && typeof sheetId === 'string') toggled.add(sheetId)
  }
  const guarded = (sheetId: unknown): sheetId is string =>
    typeof sheetId === 'string' && !toggled.has(sheetId) && sheetIsProtected(state, sheetId)
  const locked = (sheetId: string, range: string): boolean => {
    let bounds: RangeBounds
    try {
      bounds = parseRange(range)
    } catch {
      return false // the op's own validation reports a bad range
    }
    return rangeHasLockedCell(state, sheetId, [bounds])
  }
  for (const op of ops) {
    const record = op as Record<string, unknown>
    if (!guarded(record.sheetId)) continue
    if (PROTECTED_STRUCTURE_OPS.has(op.op)) return 'appSheetProtectedStructure'
    const field = PROTECTED_RANGE_OPS[op.op]
    const range = field === undefined ? undefined : record[field]
    if (typeof range === 'string' && locked(record.sheetId, range)) return 'appCellProtected'
  }
  for (const change of cellChanges) {
    if (guarded(change.sheetId) && locked(change.sheetId, change.address)) {
      return 'appCellProtected'
    }
  }
  return null
}
