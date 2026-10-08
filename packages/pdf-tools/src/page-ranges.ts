/**
 * Page-range expressions as people type them: "1-3, 5, 8-", "last", "2-end",
 * "-4". Pages are 1-based in the expression and 0-based in the result.
 */

export class PageRangeError extends Error {
  constructor(
    message: string,
    /** the piece of the expression that failed, for the UI to point at */
    readonly part: string,
  ) {
    super(message)
    this.name = 'PageRangeError'
  }
}

function parsePageNumber(token: string, pageCount: number, part: string): number {
  const t = token.trim().toLowerCase()
  if (t === 'last' || t === 'end' || t === 'z') return pageCount
  if (!/^\d+$/.test(t)) throw new PageRangeError(`not a page number: "${token}"`, part)
  const n = Number(t)
  if (n < 1 || n > pageCount) {
    throw new PageRangeError(`page ${n} is outside 1-${pageCount}`, part)
  }
  return n
}

/** One contiguous run, 0-based and inclusive; `from > to` runs backwards. */
export interface PageRun {
  from: number
  to: number
}

/**
 * Parse an expression into runs, in the order written. An empty expression
 * means every page. Throws PageRangeError on anything it cannot read.
 */
export function parsePageRuns(expr: string, pageCount: number): PageRun[] {
  if (pageCount < 1) return []
  const trimmed = expr.trim()
  if (trimmed === '' || /^all$/i.test(trimmed)) return [{ from: 0, to: pageCount - 1 }]
  const runs: PageRun[] = []
  for (const raw of trimmed.split(/[,;]/)) {
    const part = raw.trim()
    if (part === '') continue
    const dash = part.indexOf('-')
    if (dash === -1) {
      const n = parsePageNumber(part, pageCount, part)
      runs.push({ from: n - 1, to: n - 1 })
      continue
    }
    const left = part.slice(0, dash).trim()
    const right = part.slice(dash + 1).trim()
    const from = left === '' ? 1 : parsePageNumber(left, pageCount, part)
    const to = right === '' ? pageCount : parsePageNumber(right, pageCount, part)
    runs.push({ from: from - 1, to: to - 1 })
  }
  if (runs.length === 0) throw new PageRangeError('no pages selected', trimmed)
  return runs
}

/** Expand runs into a page list in expression order (duplicates kept). */
export function expandRuns(runs: readonly PageRun[]): number[] {
  const pages: number[] = []
  for (const { from, to } of runs) {
    const step = from <= to ? 1 : -1
    for (let p = from; p !== to + step; p += step) pages.push(p)
  }
  return pages
}

/** Pages an expression names, in the order written (duplicates kept). */
export function parsePageList(expr: string, pageCount: number): number[] {
  return expandRuns(parsePageRuns(expr, pageCount))
}

/** Pages an expression names as a sorted set, for "apply to these pages". */
export function parsePageSet(expr: string, pageCount: number): Set<number> {
  return new Set(parsePageList(expr, pageCount))
}

/** "1-3, 5" for 0-based [0, 1, 2, 4]: the short label used in output names. */
export function formatPageList(pages: readonly number[]): string {
  const parts: string[] = []
  let i = 0
  while (i < pages.length) {
    let j = i
    while (j + 1 < pages.length && pages[j + 1] === pages[j] + 1) j++
    parts.push(j === i ? `${pages[i] + 1}` : `${pages[i] + 1}-${pages[j] + 1}`)
    i = j + 1
  }
  return parts.join(',')
}
