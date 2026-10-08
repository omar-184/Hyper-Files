import { MAX_REDACTION_REGIONS } from '../shared/ipc'
import type { RedactionInput } from '../shared/ipc'
import type { SearchMatch } from './search'
import { rectsNear } from './edit-state'

type Rect = RedactionInput['rect']

/**
 * Viewer search rects (OCR pages) run from the baseline to the item height, so descenders (g, p, y)
 * and the tops of accented capitals would survive a mark drawn on them as-is.
 * Grow each rect to cover the full glyph box.
 */
export function padMatchRect([x1, y1, x2, y2]: Rect): Rect {
  const h = Math.abs(y2 - y1)
  return [x1 - h * 0.05, y1 - h * 0.3, x2 + h * 0.05, y2 + h * 0.1]
}

export interface MatchMarks {
  /** New marks to append, in match order */
  added: RedactionInput[]
  /** Match rects left out because the per-apply cap was reached */
  overCap: number
}

/**
 * Redaction marks for every rect of every match, skipping rects already marked
 * (running the same search twice adds nothing) and stopping at the apply cap.
 * Rects are used as given: callers pad approximate ones with padMatchRect.
 */
export function marksFromMatches(
  matches: readonly SearchMatch[],
  existing: readonly RedactionInput[],
): MatchMarks {
  const added: RedactionInput[] = []
  let overCap = 0
  const taken = (pageIndex: number, rect: Rect) =>
    [...existing, ...added].some((m) => m.pageIndex === pageIndex && rectsNear(m.rect, rect, 0.5))
  for (const match of matches) {
    for (const rect of match.rects) {
      if (taken(match.pageIndex, rect)) continue
      if (existing.length + added.length >= MAX_REDACTION_REGIONS) {
        overCap++
        continue
      }
      added.push({ pageIndex: match.pageIndex, rect })
    }
  }
  return { added, overCap }
}
