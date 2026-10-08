import { PluginKey } from '@tiptap/pm/state'
import type { DecorationSet } from '@tiptap/pm/view'
import createHyphenator from 'hyphen'
import enUsPatterns from 'hyphen/patterns/en-us'

/**
 * English (US) word hyphenation: Liang's TeX patterns (the ones Word's and
 * LibreOffice's en-US dictionaries derive from). Electron ships no Chromium
 * hyphenation dictionaries, so CSS hyphens:auto never breaks a word; the
 * hyphenation extension asks this table where a line-end word may split.
 */

const MARK = '‧'
/** at least two letters before the hyphen and three after it (TeX en-US
 *  \lefthyphenmin / \righthyphenmin; Word never leaves "ly" on a line either) */
const LEFT_MIN = 2
const RIGHT_MIN = 3
const CACHE_MAX = 5_000

let hyphenator: ((text: string) => string) | null = null
const cache = new Map<string, number[]>()

/** letters the en-US patterns know: plain Latin, case-folded by the hyphenator */
const WORD_RE = /^[A-Za-z]+$/

/** split offsets (characters before each break) inside a letters-only word,
 *  ascending; [] when the word may not break */
export function hyphenBreaks(word: string): number[] {
  if (word.length < LEFT_MIN + RIGHT_MIN || !WORD_RE.test(word)) return []
  const hit = cache.get(word)
  if (hit) return hit
  hyphenator ??= createHyphenator(enUsPatterns, { hyphenChar: MARK, minWordLength: 1 })
  const marked = hyphenator(word)
  const out: number[] = []
  let letters = 0
  for (const ch of marked) {
    if (ch === MARK) {
      if (letters >= LEFT_MIN && word.length - letters >= RIGHT_MIN) out.push(letters)
    } else letters++
  }
  if (cache.size >= CACHE_MAX) cache.clear()
  cache.set(word, out)
  return out
}

/** the hyphenation extension's soft hyphen widgets (here so the line
 *  measurers it shares wrap points with can read them without an import cycle) */
export const hyphenationPluginKey = new PluginKey<DecorationSet>('hyphenation')

/** soft hyphen widget positions inside [from, to], relative to from */
export function hyphenPointsIn(set: DecorationSet | undefined, from: number, to: number): number[] {
  if (!set) return []
  return set.find(from, to).map((d) => d.from - from)
}
