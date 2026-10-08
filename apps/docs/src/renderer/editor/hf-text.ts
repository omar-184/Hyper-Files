/**
 * Header/footer paragraph helpers shared by the on-canvas editing surface and
 * the page renderer. The invisible PAGE field sentinel shows as a visible
 * {PAGE} token in hints.
 */
import {
  PAGE_MARK,
  type HeaderFooter,
  type HfImage,
  type HfParagraph,
  type Run,
} from '@genoffice/docx-engine'

export const PAGE_TOKEN = '{PAGE}'

/** effective paragraphs: rich paras when present, else the legacy single line */
export function hfParasOf(value: HeaderFooter, images?: HfImage[] | null): HfParagraph[] {
  if (value.paras?.length) return value.paras
  const runs: Run[] = value.text ? [{ text: value.text }] : []
  if (value.pageNumber && !value.text.includes('#') && !value.text.includes(PAGE_MARK)) {
    runs.push({ text: runs.length > 0 ? ` ${PAGE_MARK}` : PAGE_MARK })
  }
  // Word reserves exactly the picture height for a picture-only part: no text line
  if (runs.length === 0 && images?.some((im) => !im.floating)) return []
  return [{ align: 'center', runs }]
}
