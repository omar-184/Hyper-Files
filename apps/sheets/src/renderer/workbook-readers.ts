/**
 * The App refs that whole-workbook readers (range aggregation, workbook
 * search) close over: the live Univer grid for opened files, the in-memory
 * snapshot for the blank demo workbook.
 */
import type { InMemoryWorkbookAdapter } from '@genoffice/xlsx-gateway/domain/in-memory-workbook'
import type { LazyWorkbookState, UniverRuntime } from './univer-state'

/** The App refs the readers need; passed per call so they never go stale. */
export interface WorkbookReadContext {
  univerRef: { readonly current: UniverRuntime | null }
  lazyWorkbookRef: { readonly current: LazyWorkbookState | null }
  adapterRef: { readonly current: InMemoryWorkbookAdapter }
}
