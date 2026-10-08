import { TEXT_BOX_ASCENT, TEXT_BOX_LEADING, TEXT_BOX_PAD, wrapTextBox } from '../shared/text-box'

/** Helvetica on macOS, Arial (same metrics) on Windows: what the saved Helvetica draws like */
export const TEXT_BOX_CSS_FONT = "Helvetica, Arial, 'Liberation Sans', sans-serif"

let ctx: CanvasRenderingContext2D | null | undefined

/** Text width in points at the given size; a rough fallback where canvas is unavailable (tests) */
export function measureTextBox(fontSize: number): (s: string) => number {
  if (ctx === undefined) {
    try {
      ctx = document.createElement('canvas').getContext('2d')
    } catch {
      ctx = null
    }
  }
  const c = ctx
  if (!c) return (s) => s.length * fontSize * 0.5
  return (s) => {
    c.font = `${fontSize}px ${TEXT_BOX_CSS_FONT}`
    return c.measureText(s).width
  }
}

/** Lines of a text box whose inner width is `width` points */
export const textBoxLines = (text: string, width: number, fontSize: number): string[] =>
  wrapTextBox(text, width - 2 * TEXT_BOX_PAD, measureTextBox(fontSize))

/**
 * Render the text of an upright box (visual size w × h points) to a PNG for
 * text the standard Helvetica cannot encode; the border is drawn by the saved
 * appearance itself. Returns base64 without the data: prefix.
 */
export function rasterTextBox(
  text: string,
  w: number,
  h: number,
  fontSize: number,
  color: readonly [number, number, number],
  pxPerPt = 3,
): string {
  const canvas = document.createElement('canvas')
  canvas.width = Math.max(1, Math.ceil(w * pxPerPt))
  canvas.height = Math.max(1, Math.ceil(h * pxPerPt))
  const g = canvas.getContext('2d')!
  g.scale(pxPerPt, pxPerPt)
  g.font = `${fontSize}px ${TEXT_BOX_CSS_FONT}`
  g.fillStyle = `rgb(${color.map((v) => Math.round(v * 255)).join(',')})`
  textBoxLines(text, w, fontSize).forEach((line, i) => {
    g.fillText(
      line,
      TEXT_BOX_PAD,
      TEXT_BOX_PAD + fontSize * TEXT_BOX_ASCENT + i * fontSize * TEXT_BOX_LEADING,
    )
  })
  return canvas.toDataURL('image/png').replace(/^data:image\/png;base64,/, '')
}
