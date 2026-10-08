/**
 * Text matching shared by the viewer's search bar and the main-process
 * find-and-redact pass, so both find the same occurrences.
 */
import { foldCase } from '../../../../packages/ui/src/find-text'

/** Kinds of sensitive text the find-and-redact tool can look for besides a typed phrase */
export type RedactPattern = 'email' | 'phone' | 'card'

/** What to find: a typed phrase (case-insensitive) or a sensitive-text pattern */
export type FindTarget = { query: string } | { pattern: RedactPattern }

const PATTERN_SOURCES: Record<RedactPattern, string> = {
  email: String.raw`[A-Za-z0-9._%+-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)*\.[A-Za-z]{2,}`,
  // Optional country code and area code, then two to five digit groups split by
  // space, dot or dash. Line breaks never join groups: a phone number set across
  // two lines is rare, a column of unrelated numbers is not.
  phone: String.raw`(?<![\w+])(?:\+\d{1,3}[ .-]?)?(?:\(\d{1,4}\)[ .-]?)?\d{2,4}(?:[ .-]?\d{2,4}){1,4}(?!\w)`,
  card: String.raw`(?<!\d)\d(?:[ -]?\d){12,18}(?!\d)`,
}

/** Luhn checksum: separates real card numbers from other long digit runs */
export function luhnValid(digits: string): boolean {
  let sum = 0
  for (let i = 0; i < digits.length; i++) {
    let d = digits.charCodeAt(digits.length - 1 - i) - 48
    if (i % 2 === 1) {
      d *= 2
      if (d > 9) d -= 9
    }
    sum += d
  }
  return sum % 10 === 0
}

function acceptPattern(pattern: RedactPattern, text: string): boolean {
  const digits = text.replace(/\D/g, '')
  if (pattern === 'phone') {
    // ISO dates share the digit-group shape of a phone number
    if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return false
    return digits.length >= 7 && digits.length <= 15
  }
  if (pattern === 'card') return luhnValid(digits)
  return true
}

/** [start, end) code-unit ranges of every non-overlapping occurrence in `text`.
    `folded` is foldCase(text) when the caller already has it. */
export function matchRanges(
  text: string,
  target: FindTarget,
  limit = Infinity,
  folded?: string,
): [number, number][] {
  const out: [number, number][] = []
  if ('query' in target) {
    const q = foldCase(target.query)
    if (!q) return out
    const lower = folded ?? foldCase(text)
    for (let s = lower.indexOf(q); s >= 0 && out.length < limit; s = lower.indexOf(q, s + q.length))
      out.push([s, s + q.length])
    return out
  }
  const re = new RegExp(PATTERN_SOURCES[target.pattern], 'g')
  for (let m = re.exec(text); m && out.length < limit; m = re.exec(text)) {
    if (m[0].length === 0) {
      re.lastIndex++
      continue
    }
    if (acceptPattern(target.pattern, m[0])) out.push([m.index, m.index + m[0].length])
  }
  return out
}
