import type { LocalMarkup } from './annotations'
import type { LocalDrawing } from './DrawLayer'
import type { LocalAnnotDelete, SavedMarkupAnnot, SavedShapeAnnot } from './edit-state'
import type { NoteThreadItem } from './note-threads'
import type { MarkupType, ShapeAnnotType } from '../shared/ipc'

export type CommentKind = MarkupType | 'note' | ShapeAnnotType | 'arrow'

/** Pending drawing tools → the kind their saved annotation lists as */
const DRAWING_KIND: Record<
  'ink' | 'rect' | 'ellipse' | 'line' | 'arrow' | 'freetext',
  CommentKind
> = {
  freetext: 'freetext',
  ink: 'ink',
  rect: 'square',
  ellipse: 'circle',
  line: 'line',
  arrow: 'arrow',
}

/** What a row in the comments list acts on */
export type CommentTarget =
  | { type: 'savedMarkup'; annot: SavedMarkupAnnot }
  | { type: 'savedShape'; annot: SavedShapeAnnot }
  | { type: 'markup'; id: string }
  | { type: 'note'; root: NoteThreadItem }
  | { type: 'drawing'; id: string }

export interface CommentEntry {
  key: string
  kind: CommentKind
  /** Original page index */
  pageIndex: number
  /** PDF-space point the row scrolls to (top-left of the annotation) */
  at: [number, number]
  author: string
  /** Note text, or the words a text markup covers */
  text: string
  timeMs: number | null
  replies: number
  /** Not yet written to the file */
  pending: boolean
  target: CommentTarget
}

type Box = [number, number, number, number]

/** Bounding box of quads in the [x1,yTop,x2,yTop,x1,yBottom,x2,yBottom] layout */
export function quadBoxes(quads: readonly (readonly number[])[]): Box[] {
  return quads.flatMap((q) => {
    if (q.length < 8) return []
    const xs = [q[0]!, q[2]!, q[4]!, q[6]!]
    const ys = [q[1]!, q[3]!, q[5]!, q[7]!]
    return [[Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)] as Box]
  })
}

const topLeft = (boxes: Box[]): [number, number] => [
  Math.min(...boxes.map((b) => b[0])),
  Math.max(...boxes.map((b) => b[3])),
]

function drawingBox(input: LocalDrawing['input']): Box | null {
  switch (input.kind) {
    case 'rect':
    case 'ellipse':
    case 'image':
    case 'freetext':
      return input.rect
    case 'line':
    case 'arrow':
      return [
        Math.min(input.from[0], input.to[0]),
        Math.min(input.from[1], input.to[1]),
        Math.max(input.from[0], input.to[0]),
        Math.max(input.from[1], input.to[1]),
      ]
    case 'ink': {
      const xs = input.paths.flatMap((p) => p.filter((_, i) => i % 2 === 0))
      const ys = input.paths.flatMap((p) => p.filter((_, i) => i % 2 === 1))
      if (xs.length === 0) return null
      return [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)]
    }
    default:
      return null
  }
}

const countReplies = (item: NoteThreadItem): number =>
  item.replies.reduce((n, r) => n + 1 + countReplies(r), 0)

export interface CommentSources {
  /** Original page indices in on-screen order (pages pending deletion left out) */
  pages: readonly number[]
  savedMarkups: ReadonlyMap<number, readonly SavedMarkupAnnot[]>
  savedShapes: ReadonlyMap<number, readonly SavedShapeAnnot[]>
  annotDeletes: readonly LocalAnnotDelete[]
  markups: readonly LocalMarkup[]
  drawings: readonly LocalDrawing[]
  /** Note threads of a page, saved and pending, deletions and edits applied */
  noteThreads: (pageIndex: number) => NoteThreadItem[]
  /** Words under the given boxes on a page; '' while the text index is loading */
  textUnder: (pageIndex: number, boxes: Box[]) => string
}

/**
 * Every comment and markup in the document, saved or pending, in reading order:
 * page order on screen, then top to bottom, then left to right. Image stamps and
 * signatures placed this session are not comments and stay out, as do drawings
 * bound to form fields.
 */
export function buildCommentList(src: CommentSources): CommentEntry[] {
  const deleted = new Set(src.annotDeletes.map((d) => d.annot.objNum))
  const order = new Map(src.pages.map((p, i) => [p, i]))
  const out: CommentEntry[] = []
  for (const pageIndex of src.pages) {
    for (const annot of src.savedMarkups.get(pageIndex) ?? []) {
      if (deleted.has(annot.objNum)) continue
      const boxes = quadBoxes(annot.quads)
      out.push({
        key: `S${annot.objNum}`,
        kind: annot.type,
        pageIndex,
        at: boxes.length > 0 ? topLeft(boxes) : [annot.rect[0], annot.rect[3]],
        author: '',
        text: src.textUnder(pageIndex, boxes),
        timeMs: null,
        replies: 0,
        pending: false,
        target: { type: 'savedMarkup', annot },
      })
    }
    for (const annot of src.savedShapes.get(pageIndex) ?? []) {
      if (deleted.has(annot.objNum)) continue
      out.push({
        key: `S${annot.objNum}`,
        kind: annot.type,
        pageIndex,
        at: [Math.min(annot.rect[0], annot.rect[2]), Math.max(annot.rect[1], annot.rect[3])],
        author: annot.author,
        text: annot.contents,
        timeMs: null,
        replies: 0,
        pending: false,
        target: { type: 'savedShape', annot },
      })
    }
    for (const root of src.noteThreads(pageIndex)) {
      out.push({
        key: root.key,
        kind: 'note',
        pageIndex,
        at: root.at,
        author: root.author,
        text: root.contents,
        timeMs: root.timeMs,
        replies: countReplies(root),
        pending: root.saved === null,
        target: { type: 'note', root },
      })
    }
  }
  for (const mk of src.markups) {
    if (!order.has(mk.pageIndex)) continue
    const boxes = quadBoxes(mk.quads)
    if (boxes.length === 0) continue
    out.push({
      key: `M${mk.id}`,
      kind: mk.type,
      pageIndex: mk.pageIndex,
      at: topLeft(boxes),
      author: '',
      text: src.textUnder(mk.pageIndex, boxes),
      timeMs: null,
      replies: 0,
      pending: true,
      target: { type: 'markup', id: mk.id },
    })
  }
  for (const d of src.drawings) {
    const input = d.input
    if (
      d.formWidgetId ||
      input.kind === 'note' ||
      input.kind === 'image' ||
      input.kind === 'field' ||
      input.kind === 'link'
    )
      continue
    if (!order.has(input.pageIndex)) continue
    const box = drawingBox(input)
    if (!box) continue
    out.push({
      key: `D${d.id}`,
      kind: DRAWING_KIND[input.kind],
      pageIndex: input.pageIndex,
      at: [box[0], box[3]],
      author: input.kind === 'freetext' ? (input.author ?? '') : '',
      text: input.kind === 'freetext' ? input.contents : '',
      timeMs: input.kind === 'freetext' ? (input.createdMs ?? null) : null,
      replies: 0,
      pending: true,
      target: { type: 'drawing', id: d.id },
    })
  }
  return out.sort(
    (a, b) =>
      order.get(a.pageIndex)! - order.get(b.pageIndex)! || b.at[1] - a.at[1] || a.at[0] - b.at[0],
  )
}
