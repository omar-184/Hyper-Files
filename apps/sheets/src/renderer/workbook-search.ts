/**
 * Workbook-wide scan limits and view navigation shared by Error Checking,
 * the streamed Find bridge, and the Find & Replace panel's jump-to-hit.
 */
import type { RangeBounds } from '@genoffice/xlsx-gateway/domain/cell-address'
import { ensureLazyRangeLoaded } from './univer-sync'
import type { WorkbookReadContext } from './workbook-readers'

/** Every error value Excel can display (ECMA-376 ST_CellType plus the modern
    data-type errors): Error Checking shares this taxonomy so no evaluation
    error is silently skipped. */
export const ERROR_VALUE_RE =
  /^#(?:REF!|DIV\/0!|VALUE!|NAME\?|N\/A|NUM!|NULL!|SPILL!|CALC!|FIELD!|CONNECT!|BLOCKED!|UNKNOWN!|GETTING_DATA)$/
/** Total cells (by scanned extent) one workbook-wide scan may cover. */
export const MAX_SCAN_CELLS = 400_000
/** Row batches sized to stay under the sidecar's per-read cell budget. */
export const FILE_READ_BATCH_CELLS = 18_000

export interface SelectRangeOutcome {
  readonly ok: boolean
  readonly sheetName?: string
  readonly error?: string
}

/** Activate the target sheet, select the range, and scroll it into view. */
export async function selectWorkbookRange(
  ctx: WorkbookReadContext,
  sheetId: string | undefined,
  bounds: RangeBounds,
  setMessage: (message: string) => void,
): Promise<SelectRangeOutcome> {
  const runtime = ctx.univerRef.current
  const workbook = runtime?.univerAPI.getActiveWorkbook()
  if (!runtime || !workbook) return { ok: false, error: 'No workbook is currently open.' }
  const worksheet = sheetId ? workbook.getSheetBySheetId(sheetId) : workbook.getActiveSheet()
  if (!worksheet) return { ok: false, error: `Unknown sheet: ${sheetId}` }
  try {
    if (worksheet.getSheetId() !== workbook.getActiveSheet()?.getSheetId()) {
      workbook.setActiveSheet(worksheet)
    }
    if (ctx.lazyWorkbookRef.current) {
      // Best-effort: selection works on not-yet-streamed cells too, but
      // loading the range means the user sees data, not an empty grid
      await ensureLazyRangeLoaded(
        runtime,
        ctx.lazyWorkbookRef,
        worksheet,
        { ...bounds },
        setMessage,
      )
    }
    worksheet
      .getRange(
        bounds.startRow,
        bounds.startColumn,
        bounds.endRow - bounds.startRow + 1,
        bounds.endColumn - bounds.startColumn + 1,
      )
      .activate()
    worksheet.scrollToCell(bounds.startRow, bounds.startColumn)
    return { ok: true, sheetName: worksheet.getSheetName() }
  } catch (error) {
    return {
      ok: false,
      error: error instanceof Error ? error.message : 'Selection failed',
    }
  }
}
