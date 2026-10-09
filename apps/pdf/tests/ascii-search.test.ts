import { describe, expect, it } from 'vitest'
import { containsAscii, indexOfAscii } from '../src/shared/ascii-search'

const enc = (s: string) => new TextEncoder().encode(s)

describe('indexOfAscii', () => {
  it('finds the first occurrence', () => {
    expect(indexOfAscii(enc('abcabc'), 'bc')).toBe(1)
    expect(indexOfAscii(enc('abcabc'), 'bc', 2)).toBe(4)
    expect(indexOfAscii(enc('abcabc'), 'abc', 1)).toBe(3)
  })

  it('skips partial matches of the first byte', () => {
    expect(indexOfAscii(enc('GGenGenOffice'), 'GenOffice')).toBe(4)
  })

  it('returns -1 when absent or cut off at the end', () => {
    expect(indexOfAscii(enc('abcab'), 'abd')).toBe(-1)
    expect(indexOfAscii(enc('xxGenOff'), 'GenOffice')).toBe(-1)
    expect(indexOfAscii(new Uint8Array(), 'a')).toBe(-1)
  })

  it('finds a match that ends exactly at the last byte', () => {
    expect(indexOfAscii(enc('xx/XFA'), '/XFA')).toBe(2)
  })
})

describe('containsAscii', () => {
  it('reports presence', () => {
    expect(
      containsAscii(enc('<< /GenOfficeStaticFormFills <...> >>'), 'GenOfficeStaticFormFills'),
    ).toBe(true)
    expect(containsAscii(enc('<< /Type /Catalog >>'), 'GenOfficeStaticFormFills')).toBe(false)
  })
})
