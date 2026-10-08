/// Streaming range statistics for the status bar: footer counts on an
/// 88k-row streamed sheet must come from batched reads, never from
/// COUNTIF-style array formulas (quadratic main-thread evaluation — the
/// app-freeze incident).

import type { CellScalar } from '@genoffice/xlsx-gateway/domain/workbook.types'

/// Distinct tracking stops (and the result says so) past this many unique
/// values; counts and numeric stats stay exact.
const MAX_DISTINCT_TRACKED = 200_000

export interface RangeAggregate {
  readonly cells: number
  readonly nonEmpty: number
  /** null when distinct tracking overflowed MAX_DISTINCT_TRACKED */
  readonly distinct: number | null
  readonly numericCount: number
  readonly sum: number
  readonly min: number | null
  readonly max: number | null
  readonly average: number | null
  /** most frequent values, descending; empty when tracking overflowed */
  readonly topValues: readonly { value: string; count: number }[]
}

interface RangeAggregator {
  add(value: CellScalar): void
  addRepeated(value: CellScalar, count: number): void
  addEmpty(count: number): void
  finish(topValueCount: number): RangeAggregate
}

export function createRangeAggregator(): RangeAggregator {
  let cells = 0
  let nonEmpty = 0
  let numericCount = 0
  let sum = 0
  let min: number | null = null
  let max: number | null = null
  let overflowed = false
  const counts = new Map<string, number>()

  const addRepeated = (value: CellScalar, count: number): void => {
    if (count <= 0) return
    cells += count
    if (value === null || value === '') return
    nonEmpty += count
    if (typeof value === 'number') {
      numericCount += count
      sum += value * count
      min = min === null ? value : Math.min(min, value)
      max = max === null ? value : Math.max(max, value)
    }
    if (overflowed) return
    const key = `${typeof value}:${String(value)}`
    const existing = counts.get(key)
    if (existing !== undefined) {
      counts.set(key, existing + count)
    } else if (counts.size >= MAX_DISTINCT_TRACKED) {
      overflowed = true
      counts.clear()
    } else {
      counts.set(key, count)
    }
  }

  return {
    add(value: CellScalar): void {
      addRepeated(value, 1)
    },
    addRepeated,
    addEmpty(count: number): void {
      if (count <= 0) return
      cells += count
    },
    finish(topValueCount: number): RangeAggregate {
      const topValues = overflowed
        ? []
        : [...counts.entries()]
            .sort((left, right) => right[1] - left[1])
            .slice(0, Math.max(0, topValueCount))
            .map(([key, count]) => ({ value: key.slice(key.indexOf(':') + 1), count }))
      return {
        cells,
        nonEmpty,
        distinct: overflowed ? null : counts.size,
        numericCount,
        sum,
        min,
        max,
        average: numericCount > 0 ? sum / numericCount : null,
        topValues,
      }
    },
  }
}
