import type { ReactElement } from 'react'
import type { ToolId } from '../../../shared/pdf-tools-api'

/** Glyph drawn inside the page outline for each tool (24×24, stroke icons). */
const GLYPHS: Record<ToolId, ReactElement> = {
  merge: <path d="M8 9h8M8 12h8M12 14.5v3M10.2 16l1.8 1.8 1.8-1.8" />,
  split: <path d="M8 10h8M8 15h8M6 12.5h2M10 12.5h1M13 12.5h1M16 12.5h2" />,
  extract: <path d="M9 10h6v6H9zM12 7v3" />,
  delete: <path d="M9.5 10.5l5 5M14.5 10.5l-5 5" />,
  rotate: <path d="M15.5 11.5a3.6 3.6 0 1 0 .3 2.6M15.8 9.2v2.6h-2.6" />,
  reorder: <path d="M10 9v8M8.3 15.3L10 17l1.7-1.7M14 17V9M12.3 10.7L14 9l1.7 1.7" />,
  compress: (
    <path d="M12 8.5v3M10.3 10l1.7 1.7 1.7-1.7M12 17.5v-3M10.3 16l1.7-1.7 1.7 1.7M9 13h6" />
  ),
  repair: <path d="M14.8 9.2a2.4 2.4 0 0 0-3 3L8.8 15.2l1 1 3-3a2.4 2.4 0 0 0 3-3l-1.4 1.4-1-1z" />,
  flatten: <path d="M8.5 11.5h7M8.5 14.5h7M10 17h4" />,
  'images-to-pdf': <path d="M8.5 15.5l2.2-2.6 1.8 2 1.3-1.4 1.7 2M9 9.5h6v7H9z" />,
  'pdf-to-images': <path d="M8.5 15.5l2.2-2.6 1.8 2 1.3-1.4 1.7 2M13.8 10.6h.01" />,
  'extract-images': <path d="M9 11h6v5H9zM9 15l2-2 2 2 2-1.5M12 8v2" />,
  protect: <path d="M9.5 12.5h5v4h-5zM10.5 12.5v-1.3a1.5 1.5 0 0 1 3 0v1.3" />,
  unlock: <path d="M9.5 12.5h5v4h-5zM10.5 12.5v-1.3a1.5 1.5 0 0 1 2.9-.5" />,
  watermark: <path d="M8.5 16l7-7M10 9.5h2M13 15.5h2" />,
  'page-numbers': <path d="M11 15.5h2M12 15.5v-4.5l-1 .8M8.5 9h7" />,
  properties: <path d="M9 10h6M9 12.5h6M9 15h3.5" />,
  ocr: <path d="M8.5 11V9.5H10M14 9.5h1.5V11M15.5 15v1.5H14M10 16.5H8.5V15M10.5 12h3M10.5 14h2" />,
}

export function ToolIcon({ id, size = 28 }: { id: ToolId; size?: number }) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.3"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      <path d="M6.5 3.5h7.5l3.5 3.5v13.5H6.5z" />
      <path d="M14 3.5V7h3.5" />
      {GLYPHS[id]}
    </svg>
  )
}
