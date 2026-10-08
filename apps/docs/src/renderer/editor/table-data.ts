import { Fragment, type Node as PmNode, type Schema } from '@tiptap/pm/model'
import type { Command, EditorState } from '@tiptap/pm/state'
import { TextSelection } from '@tiptap/pm/state'
import { TableMap, isInTable, selectedRect } from '@tiptap/pm/tables'
import { tableModelToPmNode, type PmNode as JsonNode } from './convert'

/** Word's Table Layout > Data group: Sort, Convert to Text, and Insert > Table >
 *  Convert Text to Table. Each one is a single replace step, so one undo reverts it. */

export type SortKeyType = 'text' | 'number' | 'date'

export interface SortKey {
  /** grid column index (0-based) */
  column: number
  type: SortKeyType
  descending: boolean
}

export interface SortSpec {
  keys: SortKey[]
  /** leave the first row (or every leading repeat-header row) in place */
  hasHeader: boolean
  caseSensitive?: boolean
}

export type CellSeparator = 'tab' | 'comma' | 'paragraph' | { other: string }

export type SortRefusal = 'notInTable' | 'mergedRows' | 'nothingToSort'

interface TableAt {
  node: PmNode
  pos: number
}

function tableAtSelection(state: EditorState): TableAt | null {
  if (!isInTable(state)) return null
  try {
    const rect = selectedRect(state)
    const pos = rect.tableStart - 1
    const node = state.doc.nodeAt(pos)
    return node?.type.name === 'docTable' ? { node, pos } : null
  } catch {
    return null
  }
}

/** leading rows that stay put: every repeat-header row, or just the first row */
export function headerRowCount(table: PmNode, hasHeader: boolean): number {
  if (!hasHeader) return 0
  let count = 0
  while (count < table.childCount && table.child(count).attrs.repeatHeader) count++
  return Math.min(table.childCount, Math.max(1, count))
}

/** Word's default for "My list has a header row": the first row is a repeat
 *  header, or every cell in it is a header cell / bold. */
export function guessHasHeader(state: EditorState): boolean {
  const table = tableAtSelection(state)?.node
  const first = table?.firstChild
  if (!table || !first || table.childCount < 2) return false
  if (first.attrs.repeatHeader) return true
  let header = true
  first.forEach((cell) => {
    if (cell.type.name !== 'docTableHeader' && !cell.attrs.bold) header = false
  })
  return header
}

/** Column labels for the Sort dialog: header text when there is a header row,
 *  otherwise "Column N" (the caller formats the fallback). */
export function sortColumnLabels(state: EditorState, hasHeader: boolean): (string | null)[] {
  const table = tableAtSelection(state)?.node
  if (!table) return []
  const map = TableMap.get(table)
  return Array.from({ length: map.width }, (_, c) => {
    if (!hasHeader) return null
    const text = cellText(table.nodeAt(map.map[c])!).trim()
    return text || null
  })
}

function cellText(cell: PmNode): string {
  return cell.textBetween(0, cell.content.size, ' ', ' ')
}

const NUMBER_RE = /[-−]?\(?\d[\d,\s]*(?:\.\d+)?\)?|[-−]?\(?\.\d+\)?/

/** Numeric sort key the way Word reads a cell: the first number in it, with
 *  thousands separators and currency ignored and (123) read as negative. */
export function numberKey(text: string): number | null {
  const m = NUMBER_RE.exec(text)
  if (!m) return null
  const raw = m[0].replace(/[\s,]/g, '')
  const negative = /^[-−]/.test(raw) || (raw.includes('(') && raw.includes(')'))
  const value = Number(raw.replace(/[-−()]/g, ''))
  if (!Number.isFinite(value)) return null
  return negative ? -value : value
}

/** Date sort key: ISO and English month-name forms via Date.parse, and
 *  numeric d/m/y or m/d/y (the day/month order guessed from values over 12). */
export function dateKey(text: string): number | null {
  const s = text.trim()
  if (!s) return null
  const numeric = /^(\d{1,4})[./-](\d{1,2})[./-](\d{1,4})$/.exec(s)
  if (numeric) {
    const [a, b] = [Number(numeric[1]), Number(numeric[2])]
    let c = Number(numeric[3])
    let year: number, month: number, day: number
    if (numeric[1].length === 4) [year, month, day] = [a, b, c]
    else {
      if (c < 100) c += c < 50 ? 2000 : 1900
      year = c
      // US order unless the first part can only be a day
      if (a > 12) [day, month] = [a, b]
      else [month, day] = [a, b]
    }
    const time = Date.UTC(year, month - 1, day)
    return Number.isFinite(time) && month >= 1 && month <= 12 && day >= 1 && day <= 31 ? time : null
  }
  const parsed = Date.parse(s)
  return Number.isFinite(parsed) ? parsed : null
}

type Key = string | number | null

function keyOf(text: string, type: SortKeyType): Key {
  if (type === 'number') return numberKey(text)
  if (type === 'date') return dateKey(text)
  return text.trim()
}

/** null when the table can be sorted, otherwise why not */
export function sortRefusal(state: EditorState, hasHeader: boolean): SortRefusal | null {
  const found = tableAtSelection(state)
  if (!found) return 'notInTable'
  const { node } = found
  const header = headerRowCount(node, hasHeader)
  if (node.childCount - header < 2) return 'nothingToSort'
  // a row-spanning cell ties rows together: Word refuses to sort those too
  let merged = false
  node.forEach((row) =>
    row.forEach((cell) => {
      if ((cell.attrs.rowspan as number) > 1) merged = true
    }),
  )
  return merged ? 'mergedRows' : null
}

/** Sort the table's rows by up to three columns (Word's Sort dialog). Rows move
 *  whole, keeping their height, shading and cell formatting. */
export function sortTableRows(spec: SortSpec): Command {
  return (state, dispatch) => {
    if (spec.keys.length === 0 || sortRefusal(state, spec.hasHeader)) return false
    const { node: table, pos } = tableAtSelection(state)!
    const map = TableMap.get(table)
    const header = headerRowCount(table, spec.hasHeader)
    const collator = new Intl.Collator(undefined, {
      sensitivity: spec.caseSensitive ? 'case' : 'base',
      numeric: true,
    })
    const keys = spec.keys.filter((k) => k.column >= 0 && k.column < map.width)
    if (keys.length === 0) return false
    const rows: { row: PmNode; keys: Key[] }[] = []
    for (let r = header; r < table.childCount; r++) {
      rows.push({
        row: table.child(r),
        keys: keys.map((k) =>
          keyOf(cellText(table.nodeAt(map.map[r * map.width + k.column])!), k.type),
        ),
      })
    }
    const compare = (a: Key, b: Key): number => {
      if (typeof a === 'number' && typeof b === 'number') return a - b
      return collator.compare(String(a), String(b))
    }
    rows.sort((x, y) => {
      for (let i = 0; i < keys.length; i++) {
        const a = x.keys[i]
        const b = y.keys[i]
        // empty and unreadable values go to the bottom in either direction, like Excel
        const aMissing = a === null || a === ''
        const bMissing = b === null || b === ''
        if (aMissing || bMissing) {
          if (aMissing && bMissing) continue
          return aMissing ? 1 : -1
        }
        const diff = compare(a, b)
        if (diff !== 0) return keys[i].descending ? -diff : diff
      }
      return 0
    })
    const order = rows.map((r) => r.row)
    let unchanged = true
    order.forEach((row, i) => {
      if (row !== table.child(header + i)) unchanged = false
    })
    if (unchanged) return true
    if (!dispatch) return true
    const kept: PmNode[] = []
    for (let r = 0; r < header; r++) kept.push(table.child(r))
    const next = table.type.create(table.attrs, [...kept, ...order])
    const tr = state.tr.replaceWith(pos, pos + table.nodeSize, next)
    tr.setSelection(TextSelection.near(tr.doc.resolve(pos + 1)))
    dispatch(tr.scrollIntoView())
    return true
  }
}

function separatorText(sep: CellSeparator): string | null {
  if (sep === 'tab') return '\t'
  if (sep === 'comma') return ','
  if (sep === 'paragraph') return null
  return sep.other.slice(0, 1) || '\t'
}

function blankParagraph(schema: Schema, like: PmNode | null): PmNode {
  const attrs =
    like?.type.name === 'docParagraph' ? { ...like.attrs, docxIndex: null } : { docxIndex: null }
  return schema.nodes.docParagraph.create(attrs)
}

/** whether Convert to Text would drop content it cannot flatten (nested tables, anchored boxes) */
export function tableHasUnflattenable(state: EditorState): boolean {
  const table = tableAtSelection(state)?.node
  if (!table) return false
  let found = false
  table.descendants((node) => {
    if (node.type.name === 'docNestedTable' || node.type.name === 'docCellBoxes') found = true
    return !found
  })
  return found
}

/** Replace the table with paragraphs: one per row with cells joined by the
 *  separator, or one per cell for the paragraph-mark separator. Cell text keeps
 *  its character formatting. */
export function convertTableToText(sep: CellSeparator): Command {
  return (state, dispatch) => {
    const found = tableAtSelection(state)
    if (!found) return false
    if (!dispatch) return true
    const { node: table, pos } = found
    const schema = state.schema
    const joint = separatorText(sep)
    const out: PmNode[] = []
    const map = TableMap.get(table)
    const seen = new Set<number>()
    for (let r = 0; r < map.height; r++) {
      let current: PmNode | null = null
      const append = (like: PmNode | null, content: Fragment) => {
        current = current
          ? current.copy(current.content.append(content))
          : blankParagraph(schema, like).copy(content)
      }
      let k = 0
      for (let c = 0; c < map.width; c++) {
        const offset = map.map[r * map.width + c]
        if (seen.has(offset)) continue
        seen.add(offset)
        const cell = table.nodeAt(offset)!
        const paragraphs: PmNode[] = []
        cell.forEach((child) => {
          if (child.isTextblock) paragraphs.push(child)
        })
        if (joint === null) {
          for (const p of paragraphs) out.push(blankParagraph(schema, p).copy(p.content))
          if (paragraphs.length === 0) out.push(blankParagraph(schema, null))
          continue
        }
        if (k++ > 0) append(null, Fragment.from(schema.text(joint)))
        paragraphs.forEach((p, i) => {
          // a paragraph mark inside a cell stays a paragraph break
          if (i > 0 && current) {
            out.push(current)
            current = null
          }
          append(p, p.content)
        })
      }
      if (joint !== null) out.push(current ?? blankParagraph(schema, null))
    }
    const cleaned = out
    if (cleaned.length === 0) cleaned.push(blankParagraph(schema, null))
    const tr = state.tr.replaceWith(pos, pos + table.nodeSize, cleaned)
    tr.setSelection(TextSelection.near(tr.doc.resolve(pos + 1)))
    dispatch(tr.scrollIntoView())
    return true
  }
}

export interface TextToTablePlan {
  rows: number
  cols: number
}

/** top-level textblocks the selection covers, or null when it covers anything else */
function selectedParagraphs(
  state: EditorState,
): { nodes: PmNode[]; from: number; to: number } | null {
  if (isInTable(state)) return null
  const { from, to } = state.selection
  const nodes: PmNode[] = []
  let start = -1
  let end = -1
  let ok = true
  state.doc.forEach((node, offset) => {
    const nodeEnd = offset + node.nodeSize
    if (nodeEnd <= from || offset >= to) {
      // a collapsed caret still selects the paragraph it sits in
      if (!(from === to && offset < from && nodeEnd > from)) return
    }
    if (!node.isTextblock) ok = false
    nodes.push(node)
    if (start < 0) start = offset
    end = nodeEnd
  })
  return ok && nodes.length > 0 ? { nodes, from: start, to: end } : null
}

/** split one paragraph's inline content at the separator character */
function splitInline(schema: Schema, para: PmNode, joint: string): PmNode[][] {
  const fields: PmNode[][] = [[]]
  para.forEach((child) => {
    if (!child.isText) {
      fields[fields.length - 1].push(child)
      return
    }
    const parts = child.text!.split(joint)
    parts.forEach((part, i) => {
      if (i > 0) fields.push([])
      if (part) fields[fields.length - 1].push(schema.text(part, child.marks))
    })
  })
  return fields
}

/** The grid Convert Text to Table would build, for the dialog's preview */
export function textToTablePlan(
  state: EditorState,
  sep: CellSeparator,
  columns?: number,
): TextToTablePlan | null {
  const sel = selectedParagraphs(state)
  if (!sel) return null
  const joint = separatorText(sep)
  if (joint === null) {
    const cols = Math.max(1, Math.min(63, columns ?? 1))
    return { rows: Math.ceil(sel.nodes.length / cols), cols }
  }
  const cols = Math.max(...sel.nodes.map((p) => splitInline(state.schema, p, joint).length))
  return { rows: sel.nodes.length, cols: Math.min(63, Math.max(cols, columns ?? 1)) }
}

/** Turn the selected paragraphs into a table: one row per paragraph, a cell per
 *  separated field (or one cell per paragraph, filled row by row). */
export function convertTextToTable(sep: CellSeparator, columns?: number): Command {
  return (state, dispatch) => {
    const sel = selectedParagraphs(state)
    const plan = textToTablePlan(state, sep, columns)
    if (!sel || !plan) return false
    if (!dispatch) return true
    const schema = state.schema
    const joint = separatorText(sep)
    const grid: PmNode[][][] = []
    if (joint === null) {
      sel.nodes.forEach((p, i) => {
        if (i % plan.cols === 0) grid.push([])
        grid[grid.length - 1].push(p.content.content.slice())
      })
    } else {
      for (const p of sel.nodes) grid.push(splitInline(schema, p, joint))
    }
    const line = { style: 'single', szEighths: 4, color: 'auto' }
    const json = tableModelToPmNode({
      rows: Array.from({ length: plan.rows }, () =>
        Array.from({ length: plan.cols }, () => ({ paras: [''] })),
      ),
      colWidthsPct: Array.from({ length: plan.cols }, () => 100 / plan.cols),
      widthPct: 100,
      autoFit: 'window',
      borders: { top: line, bottom: line, left: line, right: line, insideH: line, insideV: line },
    })
    json.content?.forEach((row: JsonNode, r) =>
      row.content?.forEach((cell: JsonNode, c) => {
        const inline = grid[r]?.[c] ?? []
        const para = cell.content?.[0]
        if (para && inline.length > 0) para.content = inline.map((n) => n.toJSON() as JsonNode)
      }),
    )
    const table = schema.nodeFromJSON(json)
    const tr = state.tr.replaceWith(sel.from, sel.to, table)
    tr.setSelection(TextSelection.near(tr.doc.resolve(sel.from + 1)))
    dispatch(tr.scrollIntoView())
    return true
  }
}
