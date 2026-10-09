/**
 * Whole-file byte searches for large PDFs. Typed-array indexOf runs natively, so stepping
 * from one candidate first byte to the next is an order of magnitude faster than a
 * byte-by-byte JavaScript loop over a file of tens or hundreds of megabytes.
 */

/** Index of the first occurrence of an ASCII `text` at or after `from`, or -1. */
export function indexOfAscii(bytes: Uint8Array, text: string, from = 0): number {
  if (text.length === 0) return Math.min(Math.max(from, 0), bytes.length)
  const first = text.charCodeAt(0)
  const last = bytes.length - text.length
  for (let i = bytes.indexOf(first, from); i !== -1 && i <= last; i = bytes.indexOf(first, i + 1)) {
    let k = 1
    while (k < text.length && bytes[i + k] === text.charCodeAt(k)) k++
    if (k === text.length) return i
  }
  return -1
}

export function containsAscii(bytes: Uint8Array, text: string): boolean {
  return indexOfAscii(bytes, text) !== -1
}
