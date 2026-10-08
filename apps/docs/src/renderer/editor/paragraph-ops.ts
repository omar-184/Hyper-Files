import type { Editor } from '@tiptap/core'
import type { Node as PmDocNode } from '@tiptap/pm/model'
import type { Transaction } from '@tiptap/pm/state'

/**
 * Paragraph-format ops issued by the ribbon, shortcuts and dialogs. Each call
 * validates the whole batch up front and applies it as one ProseMirror
 * transaction, so one undo reverts the batch. The docx round trip is
 * untouched — changed nodes become dirty via the normal signature comparison
 * and regenerate on save.
 */

/** Which top-level blocks an op reaches; at least one condition is required */
export interface Target {
  /** By top-level block index (0-based, current PM doc order) */
  blockIndexes?: number[]
  /** 'selection' = only blocks covered by the current selection */
  scope?: 'selection' | 'document'
  /** explicit ProseMirror range standing in for the selection (callers that captured a range earlier) */
  range?: { from: number; to: number }
}

export interface Op {
  op: string
  target?: Target
  [key: string]: unknown
}

interface SelRange {
  from: number
  to: number
}

interface TopBlock {
  pos: number
  node: PmDocNode
}

interface OpDef {
  validate(op: Op, where: string): string | null
  /** returns how many paragraphs changed */
  apply(op: Op, tr: Transaction, sel: SelRange): number
}

function validateTarget(target: unknown, where: string): string | null {
  if (!target || typeof target !== 'object') return `${where}: missing target`
  const tg = target as Target
  if (
    tg.blockIndexes !== undefined &&
    (!Array.isArray(tg.blockIndexes) || tg.blockIndexes.some((i) => !Number.isInteger(i) || i < 0))
  ) {
    return `${where}: blockIndexes must be an array of non-negative integers`
  }
  if (tg.range !== undefined) {
    const r = tg.range as { from?: unknown; to?: unknown } | null
    if (!r || typeof r.from !== 'number' || typeof r.to !== 'number' || r.from > r.to) {
      return `${where}: range must be { from, to } positions with from ≤ to`
    }
  }
  const hasCondition =
    (Array.isArray(tg.blockIndexes) && tg.blockIndexes.length > 0) ||
    tg.scope === 'selection' ||
    tg.range !== undefined
  if (!hasCondition) return `${where}: target requires at least one condition`
  return null
}

function validateShape(op: Op, keys: readonly string[], where: string): string | null {
  const error = validateTarget(op.target, where)
  if (error) return error
  const unknown = Object.keys(op).filter((k) => k !== 'op' && k !== 'target' && !keys.includes(k))
  if (unknown.length > 0) {
    return `${where}: unknown field(s) ${unknown.join(', ')}; allowed: ${keys.join(', ')}`
  }
  return null
}

/** the range a selection-style target resolves against: an explicit range, else the selection */
function scopedRange(target: Target, sel: SelRange): SelRange | null {
  if (target.range) return target.range
  return target.scope === 'selection' ? sel : null
}

function matchTarget(doc: PmDocNode, target: Target, sel: SelRange): TopBlock[] {
  const scoped = scopedRange(target, sel)
  const out: TopBlock[] = []
  doc.forEach((node, pos, index) => {
    if (target.blockIndexes && !target.blockIndexes.includes(index)) return
    if (scoped) {
      const to = pos + node.nodeSize
      const overlaps = to > scoped.from && pos < scoped.to
      const caretInside = scoped.from === scoped.to && scoped.from >= pos && scoped.from <= to
      if (!overlaps && !caretInside) return
    }
    out.push({ pos, node })
  })
  return out
}

const PARAGRAPH_TYPES = new Set(['docParagraph', 'docHeading', 'docListItem'])

/**
 * Paragraph-like nodes an op formats inside a matched block: the block itself
 * when it is a text block; for containers (tables) the nested paragraphs, kept
 * to the scoped range when the target is selection-based — the ribbon's
 * "align" inside a table cell must reach that cell's paragraph only.
 */
function paragraphsIn(b: TopBlock, scoped: SelRange | null): TopBlock[] {
  if (b.node.type.name === 'docProtected') return []
  if (b.node.isTextblock) return [b]
  const hits: TopBlock[] = []
  b.node.descendants((node, offset) => {
    if (!PARAGRAPH_TYPES.has(node.type.name)) return true
    const pos = b.pos + 1 + offset
    const end = pos + node.nodeSize
    if (scoped) {
      const overlaps = end > scoped.from && pos < scoped.to
      const caretInside = scoped.from === scoped.to && scoped.from >= pos && scoped.from <= end
      if (!overlaps && !caretInside) return false
    }
    hits.push({ node, pos })
    return false
  })
  return hits
}

/** every paragraph the op's target reaches, in document order */
function targetParagraphs(op: Op, tr: Transaction, sel: SelRange): TopBlock[] {
  const target = op.target as Target
  const scoped = scopedRange(target, sel)
  return matchTarget(tr.doc, target, sel).flatMap((b) => paragraphsIn(b, scoped))
}

/**
 * Raw paragraph attrs (any schema attr, e.g. lineRule / tabStops / dropCap) on
 * every targeted paragraph; `align` also lands on targeted image blocks as their
 * w:jc. An explicit spacing value turns Word's "Auto" spacing off.
 */
function applyParagraphAttrs(op: Op, tr: Transaction, sel: SelRange): number {
  const patch = { ...(op.attrs as Record<string, unknown>) }
  const target = op.target as Target
  const scoped = scopedRange(target, sel)
  if ('spaceBefore' in patch && !('spaceBeforeAuto' in patch)) patch.spaceBeforeAuto = false
  if ('spaceAfter' in patch && !('spaceAfterAuto' in patch)) patch.spaceAfterAuto = false
  let changed = 0
  for (const b of matchTarget(tr.doc, target, sel)) {
    if (b.node.type.name === 'docProtected') {
      if ('align' in patch && b.node.attrs.blockType === 'image') {
        const align = patch.align ?? null
        if (b.node.attrs.imageAlign !== align) {
          tr.setNodeMarkup(b.pos, undefined, { ...b.node.attrs, imageAlign: align })
          changed++
        }
      }
      continue
    }
    for (const p of paragraphsIn(b, scoped)) {
      const attrs = { ...p.node.attrs, ...patch }
      const dirty = Object.keys(patch).some((k) => attrs[k] !== p.node.attrs[k])
      if (!dirty) continue
      tr.setNodeMarkup(p.pos, undefined, attrs)
      changed++
    }
  }
  return changed
}

/** Word's ribbon indent step: half an inch, in twips */
const INDENT_STEP = 720

/** ribbon increase/decrease indent — list items change level, paragraphs and headings snap to the next half-inch stop */
function applyStepIndent(op: Op, tr: Transaction, sel: SelRange): number {
  const delta = Number(op.delta)
  let changed = 0
  for (const p of targetParagraphs(op, tr, sel)) {
    if (p.node.type.name === 'docListItem') {
      const ilvl = Number(p.node.attrs.ilvl) || 0
      const next = Math.min(Math.max(ilvl + delta, 0), 8)
      if (next === ilvl) continue
      tr.setNodeMarkup(p.pos, undefined, { ...p.node.attrs, ilvl: next })
      changed++
      continue
    }
    const cur = Number(p.node.attrs.indentLeft) || 0
    const next =
      delta > 0
        ? Math.floor(cur / INDENT_STEP) * INDENT_STEP + INDENT_STEP
        : Math.max(Math.ceil(cur / INDENT_STEP) * INDENT_STEP - INDENT_STEP, 0)
    if (next === cur) continue
    tr.setNodeMarkup(p.pos, undefined, { ...p.node.attrs, indentLeft: next || null })
    changed++
  }
  return changed
}

/**
 * Word's Ctrl+T / Ctrl+Shift+T — a hanging indent moves the left indent one
 * stop right while pulling the first line back by the same amount, so
 * continuation lines sit under the body text; the reverse walks it back and
 * only clears the negative first line once the left indent is home.
 */
function applyStepHangingIndent(op: Op, tr: Transaction, sel: SelRange): number {
  const delta = Number(op.delta)
  let changed = 0
  for (const p of targetParagraphs(op, tr, sel)) {
    const left = Number(p.node.attrs.indentLeft) || 0
    const nextLeft = Math.max(
      delta > 0
        ? Math.floor(left / INDENT_STEP) * INDENT_STEP + INDENT_STEP
        : Math.ceil(left / INDENT_STEP) * INDENT_STEP - INDENT_STEP,
      0,
    )
    const nextHanging = nextLeft > 0 ? -Math.min(nextLeft, INDENT_STEP) : 0
    if (nextLeft === left && nextHanging === (Number(p.node.attrs.indentFirstLine) || 0)) continue
    tr.setNodeMarkup(p.pos, undefined, {
      ...p.node.attrs,
      indentLeft: nextLeft || null,
      indentFirstLine: nextHanging || null,
    })
    changed++
  }
  return changed
}

/** Paragraph dialog outline level — a direct w:outlineLvl, the style stays (Word never restyles for it) */
function applyOutlineLevel(op: Op, tr: Transaction, sel: SelRange): number {
  const schema = tr.doc.type.schema
  const level = Number(op.level)
  let changed = 0
  for (const p of targetParagraphs(op, tr, sel)) {
    const name = p.node.type.name
    const isHeading = name === 'docHeading'
    // a styled heading's level belongs to its style; list items have no level slot
    if (isHeading ? !p.node.attrs.outlineOnly : name !== 'docParagraph') continue
    if (level === 0) {
      if (!isHeading) continue
      const { level: _l, outlineOnly: _o, ...rest } = p.node.attrs
      tr.setNodeMarkup(p.pos, schema.nodes.docParagraph, rest)
    } else {
      if (isHeading && Number(p.node.attrs.level) === level) continue
      tr.setNodeMarkup(p.pos, schema.nodes.docHeading, {
        ...p.node.attrs,
        level,
        outlineOnly: true,
      })
    }
    changed++
  }
  return changed
}

const validDelta = (op: Op, where: string): string | null =>
  op.delta === 1 || op.delta === -1 ? null : `${where}: delta must be 1 or -1`

const REGISTRY: Record<string, OpDef> = {
  setParagraphAttrs: {
    validate(op, where) {
      const shape = validateShape(op, ['attrs'], where)
      if (shape) return shape
      if (!op.attrs || typeof op.attrs !== 'object' || Object.keys(op.attrs).length === 0) {
        return `${where}: attrs must be a non-empty object`
      }
      return null
    },
    apply: applyParagraphAttrs,
  },
  setOutlineLevel: {
    validate(op, where) {
      const shape = validateShape(op, ['level'], where)
      if (shape) return shape
      if (!Number.isInteger(op.level) || Number(op.level) < 0 || Number(op.level) > 9) {
        return `${where}: level must be an integer between 0 and 9`
      }
      return null
    },
    apply: applyOutlineLevel,
  },
  stepIndent: {
    validate: (op, where) => validateShape(op, ['delta'], where) ?? validDelta(op, where),
    apply: applyStepIndent,
  },
  stepHangingIndent: {
    validate: (op, where) => validateShape(op, ['delta'], where) ?? validDelta(op, where),
    apply: applyStepHangingIndent,
  },
}

/**
 * Apply a batch of paragraph ops as one transaction and restore focus like a
 * tiptap chain would. Returns true when something changed. A validation
 * failure is a programming error in the caller, so it is logged and nothing
 * is applied.
 */
export function runUiOps(editor: Editor, ops: Op[], opts: { focus?: boolean } = {}): boolean {
  const errors: string[] = []
  ops.forEach((op, i) => {
    const where = `op #${i} ${op.op}`
    const def = Object.hasOwn(REGISTRY, op.op) ? REGISTRY[op.op] : undefined
    const error = def ? def.validate(op, where) : `${where}: unknown op`
    if (error) errors.push(error)
  })
  if (errors.length > 0) {
    console.error('runUiOps rejected', errors.join('\n'))
    return false
  }
  const tr = editor.state.tr
  const selection = editor.state.selection
  let changed = 0
  for (const op of ops) {
    // selection positions mapped through the steps applied so far
    const sel = { from: tr.mapping.map(selection.from), to: tr.mapping.map(selection.to) }
    changed += REGISTRY[op.op]!.apply(op, tr, sel)
  }
  if (tr.steps.length > 0) editor.view.dispatch(tr)
  if (opts.focus !== false) editor.commands.focus()
  return changed > 0
}
