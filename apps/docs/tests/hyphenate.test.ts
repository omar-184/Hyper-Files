import { describe, expect, it } from 'vitest'
import { hyphenBreaks } from '../src/renderer/editor/hyphenate'
import { hyphenWords } from '../src/renderer/editor/hyphenation'

const show = (word: string) => {
  let out = ''
  let last = 0
  for (const at of hyphenBreaks(word)) {
    out += word.slice(last, at) + '-'
    last = at
  }
  return out + word.slice(last)
}

describe('hyphenBreaks', () => {
  it('splits English words at the en-US pattern points', () => {
    expect(show('hyphenation')).toBe('hy-phen-ation')
    expect(show('representative')).toBe('rep-re-sen-ta-tive')
    expect(show('documentation')).toBe('doc-u-men-ta-tion')
  })

  it('keeps two letters before and three after every break', () => {
    expect(show('extraordinarily')).toBe('ex-tra-or-di-nar-ily')
    expect(show('international')).toBe('in-ter-na-tional')
    expect(show('happily')).toBe('hap-pily')
  })

  it('handles capitals like lowercase', () => {
    expect(show('INTERNATIONAL')).toBe('IN-TER-NA-TIONAL')
    expect(show('Microsoft')).toBe('Mi-crosoft')
  })

  it('leaves short words and non-letters alone', () => {
    expect(hyphenBreaks('the')).toEqual([])
    expect(hyphenBreaks('abc123def')).toEqual([])
    expect(hyphenBreaks("don't")).toEqual([])
  })
})

describe('hyphenWords', () => {
  it('finds the letters core of each run with its offset', () => {
    const words = hyphenWords('See (documentation), then representative.', false)
    expect(words.map((w) => [w.runStart, w.start, w.word])).toEqual([
      [4, 5, 'documentation'],
      [26, 26, 'representative'],
    ])
  })

  it('skips runs with hyphens, soft hyphens, digits or inline objects', () => {
    expect(hyphenWords('well-documented inter­national ab12cdefg ￼documentation', false)).toEqual(
      [],
    )
  })

  it('skips capitals under doNotHyphenateCaps', () => {
    expect(hyphenWords('INTERNATIONAL International', true).map((w) => w.word)).toEqual([
      'International',
    ])
  })
})
