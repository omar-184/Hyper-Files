/**
 * Cell reads over the open workbook (live Univer grid for opened files, the
 * in-memory snapshot for the blank demo workbook). Extracted from App.tsx;
 * callers pass a WorkbookReadContext closing over the App refs.
 */
import type { InMemoryWorkbookAdapter } from '@genoffice/xlsx-gateway/domain/in-memory-workbook'
import type { CellScalar } from '@genoffice/xlsx-gateway/domain/workbook.types'
import { lazyCellReader } from './univer-sync'
import type { LazyWorkbookState, UniverRuntime } from './univer-state'

/** The App refs the readers need; passed per call so they never go stale. */
export interface WorkbookReadContext {
  univerRef: { readonly current: UniverRuntime | null }
  lazyWorkbookRef: { readonly current: LazyWorkbookState | null }
  adapterRef: { readonly current: InMemoryWorkbookAdapter }
}

export function readCells(
  ctx: WorkbookReadContext,
  addresses: readonly string[],
  sheetId?: string,
): Record<string, { value: CellScalar; formula?: string; rawValue?: CellScalar | undefined }> {
  const result: Record<
    string,
    { value: CellScalar; formula?: string; rawValue?: CellScalar | undefined }
  > = {}
  const workbook = ctx.univerRef.current?.univerAPI.getActiveWorkbook()
  const state = ctx.lazyWorkbookRef.current
  if (state) {
    const worksheet =
      sheetId === undefined ? workbook?.getActiveSheet() : workbook?.getSheetBySheetId(sheetId)
    if (!worksheet) return result
    const reader = lazyCellReader(worksheet)
    for (const address of addresses) {
      const cell = reader(address)
      // `value` is the rendered text and `rawValue` the model value behind it;
      // machine-facing callers (the save pipeline) need the latter — see
      // modelCellValue in univer-sync.ts.
      result[address] = cell.formula
        ? { value: cell.value, formula: cell.formula, rawValue: cell.rawValue }
        : { value: cell.value, rawValue: cell.rawValue }
    }
    return result
  }
  const sheets = ctx.adapterRef.current.getSnapshot().sheets
  const targetId = sheetId ?? workbook?.getActiveSheet()?.getSheetId()
  const sheet =
    sheets.find((s) => s.id === targetId) ?? (sheetId === undefined ? sheets[0] : undefined)
  if (!sheet) return result
  // The in-memory model stores only value:null for formula cells; computed
  // values live in Univer's formula engine, backfilled by reading the grid
  const worksheet =
    sheetId === undefined ? workbook?.getActiveSheet() : workbook?.getSheetBySheetId(sheetId)
  for (const address of addresses) {
    const cell = sheet.cells[address] ?? { value: null }
    if (cell.formula) {
      let computed = cell.value
      if (computed === null && worksheet) {
        try {
          computed = worksheet.getRange(address).getValue() ?? null
        } catch {
          /* fail-open */
        }
      }
      result[address] = { value: computed, formula: cell.formula }
    } else {
      result[address] = { value: cell.value }
    }
  }
  return result
}
