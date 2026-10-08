/**
 * Layout of a text-box comment (FreeText annotation), shared by the editor's
 * on-page preview and the saved appearance stream so both break lines alike.
 * The face is Helvetica (Arial on Windows has the same metrics).
 */

/** Inner padding between the border and the text, in points */
export const TEXT_BOX_PAD = 4
/** Line advance as a multiple of the font size */
export const TEXT_BOX_LEADING = 1.2
/** First baseline below the inner top edge, as a multiple of the font size */
export const TEXT_BOX_ASCENT = 0.85
export const TEXT_BOX_FONT_SIZES = [9, 10, 12, 14, 18, 24] as const
export const TEXT_BOX_DEFAULT_SIZE = 12
export const TEXT_BOX_DEFAULT_WIDTH = 180

/**
 * Break text into lines no wider than maxWidth: paragraphs at newlines, then
 * at spaces; a word wider than the box breaks between characters.
 */
export function wrapTextBox(
  text: string,
  maxWidth: number,
  measure: (s: string) => number,
): string[] {
  const out: string[] = []
  for (const para of text.replace(/\r\n?/g, '\n').split('\n')) {
    let line = ''
    for (const word of para.split(' ')) {
      const candidate = line ? `${line} ${word}` : word
      if (measure(candidate) <= maxWidth || !line) {
        line = candidate
      } else {
        out.push(line)
        line = word
      }
      // A single word wider than the box: split it by characters
      while (measure(line) > maxWidth && line.length > 1) {
        let cut = line.length - 1
        while (cut > 1 && measure(line.slice(0, cut)) > maxWidth) cut--
        out.push(line.slice(0, cut))
        line = line.slice(cut)
      }
    }
    out.push(line)
  }
  return out
}

/** Height (points) the wrapped lines need, padding included */
export const textBoxHeight = (lines: number, fontSize: number): number =>
  lines * fontSize * TEXT_BOX_LEADING + 2 * TEXT_BOX_PAD

/** cp1252 code points beyond Latin-1 that WinAnsiEncoding also covers */
const WIN_ANSI_EXTRA = new Set('€‚ƒ„…†‡ˆ‰Š‹ŒŽ‘’“”•–—˜™š›œžŸ')

/** Whether the standard Helvetica (WinAnsiEncoding) can draw every character */
export function isWinAnsi(text: string): boolean {
  for (const ch of text) {
    const c = ch.codePointAt(0)!
    if (ch === '\n' || ch === '\r') continue
    if ((c >= 0x20 && c <= 0x7e) || (c >= 0xa0 && c <= 0xff) || WIN_ANSI_EXTRA.has(ch)) continue
    return false
  }
  return true
}
