import { describe, expect, it } from 'vitest'
import { deriveAutoFileName } from '../src/renderer/document/auto-name'

describe('deriveAutoFileName', () => {
  it('prefers the document title', () => {
    const html =
      '<html><head><title> Riverside &amp; Books </title></head><body><h1>Hero</h1></body></html>'
    expect(deriveAutoFileName(html)).toBe('Riverside & Books')
  })

  it('falls back to the first h1, stripping inline markup', () => {
    expect(deriveAutoFileName('<body><h1 class="x">A <em>good</em> book</h1></body>')).toBe(
      'A good book',
    )
  })

  it('falls back to the first words of the body', () => {
    const words = Array.from({ length: 12 }, (_, i) => `w${i}`).join(' ')
    expect(deriveAutoFileName(`<body><p>${words}</p></body>`)).toBe('w0 w1 w2 w3 w4 w5 w6 w7')
  })

  it('caps the length', () => {
    expect(deriveAutoFileName(`<title>${'x'.repeat(100)}</title>`)).toHaveLength(60)
  })

  it('leaves out-of-range numeric entities alone instead of throwing', () => {
    expect(deriveAutoFileName('<title>a &#1114112; b &#x110000; c &#65; </title>')).toBe(
      'a &#1114112; b &#x110000; c A',
    )
  })

  it('returns empty for an empty document', () => {
    expect(deriveAutoFileName('')).toBe('')
  })
})
