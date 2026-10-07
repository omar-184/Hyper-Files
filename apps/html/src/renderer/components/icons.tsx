/**
 * Icon set for the html app: everything shared re-exports from the docs
 * ribbon library so glyph style stays uniform across the suite; the html-only
 * glyphs below follow the same 16-grid / pinned-stroke contract.
 */

import type { ReactNode } from 'react'

export {
  IconAlignCenter,
  IconAlignLeft,
  IconAlignRight,
  IconBullets,
  IconCopy,
  IconCrop,
  IconFlipH,
  IconFlipV,
  IconLink,
  IconLock,
  IconPicture,
  IconPilcrow,
  IconRedo,
  IconRemoveBg,
  IconReplacePicture,
  IconRotateLeft,
  IconRotateRight,
  IconSave,
  IconSearch,
  IconTable,
  IconTrash,
  IconUndo,
} from '../../../../docs/src/renderer/components/icons'

interface IconProps {
  size?: number
}

function pinnedStroke(size: number): number {
  const painted = size >= 20 ? 1.5 : size >= 13 ? 1.25 : 1.1
  return (painted * 16) / size
}

function Svg({ size = 20, children }: IconProps & { children: ReactNode }) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 16 16"
      fill="none"
      stroke="currentColor"
      strokeWidth={pinnedStroke(size)}
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden
    >
      {children}
    </svg>
  )
}

export function IconPlus(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M8 3v10M3 8h10" />
    </Svg>
  )
}

/** insert: heading (an H letterform) */
export function IconHeading(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M4 3v10M12 3v10M4 8h8" />
    </Svg>
  )
}

/** insert: button (a pill with its label line) */
export function IconButton(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="2" y="5" width="12" height="6" rx="3" />
      <path d="M5.5 8h5" />
    </Svg>
  )
}

/** insert: section (a block with a heading bar and body text) */
export function IconSection(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="2.5" y="2.5" width="11" height="11" rx="1" />
      <path d="M2.5 6h11M5 9h6" />
    </Svg>
  )
}

/** insert: divider (a rule between two text lines) */
export function IconDivider(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M2.5 8h11M5 4.5h6M5 11.5h6" />
    </Svg>
  )
}

export function IconMoveUp(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M8 13V4.2M4.4 7.6 8 4l3.6 3.6" />
    </Svg>
  )
}

export function IconMoveDown(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M8 3v8.8M11.6 8.4 8 12 4.4 8.4" />
    </Svg>
  )
}

export function IconCode(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M5.2 4.6 2 8l3.2 3.4M10.8 4.6 14 8l-3.2 3.4M9.4 3.2 6.6 12.8" />
    </Svg>
  )
}

export function IconPreview(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="2.2" y="3" width="11.6" height="10" rx="1.4" />
      <path d="M2.2 6h11.6" />
    </Svg>
  )
}

export function IconSplitView(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="2.2" y="3" width="11.6" height="10" rx="1.4" />
      <path d="M8 3v10" />
    </Svg>
  )
}

export function IconSliders(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M2.5 4.5h11M2.5 8h11M2.5 11.5h11" />
      <circle cx="6" cy="4.5" r="1.4" fill="var(--surface)" />
      <circle cx="10.5" cy="8" r="1.4" fill="var(--surface)" />
      <circle cx="5" cy="11.5" r="1.4" fill="var(--surface)" />
    </Svg>
  )
}

export function IconChevronDown(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="m4.5 6.5 3.5 3.5 3.5-3.5" />
    </Svg>
  )
}

export function IconExpand(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M9.5 2.5h4v4M13.5 2.5 9 7M6.5 13.5h-4v-4M2.5 13.5 7 9" />
    </Svg>
  )
}

export function IconPlay(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M4.5 3.2v9.6L12.5 8z" />
    </Svg>
  )
}

export function IconGlobe(props: IconProps) {
  return (
    <Svg {...props}>
      <circle cx="8" cy="8" r="5.5" />
      <path d="M2.5 8h11M8 2.5c1.8 1.6 2.6 3.4 2.6 5.5S9.8 12 8 13.5C6.2 11.9 5.4 10.1 5.4 8S6.2 4.1 8 2.5Z" />
    </Svg>
  )
}
