import type { Node as PmNode } from '@tiptap/pm/model'
import type { Transaction } from '@tiptap/pm/state'
import { TableMap, addColSpan, rowIsHeader, tableNodeTypes, type Rect } from '@tiptap/pm/tables'

/**
 * Grid-level table edits shared by the table commands: column widths in px per
 * grid slot, and row / column insertion that keeps merged cells and the
 * neighbouring formatting consistent.
 */

/** px per grid column when the table carries absolute widths */
export function columnWidthsPx(table: PmNode, map: TableMap): number[] | null {
  const widths: number[] = []
  const first = table.firstChild
  if (!first) return null
  for (let i = 0; i < first.childCount; i++) {
    const cw = first.child(i).attrs.colwidth as number[] | null
    if (!cw?.length) return null
    widths.push(...cw)
  }
  if (widths.length !== map.width || widths.some((w) => !(w > 0))) return null
  return widths
}

/**
 * Re-derive every cell's colwidth and the table's width attrs from one widths
 * array (px per grid column). `widths === null` keeps the table percentage
 * based: cells lose their px widths and colWidthsPct is renormalized.
 */
export function reflowColumns(
  tr: Transaction,
  tablePos: number,
  widths: number[] | null,
  pct: number[],
): void {
  const table = tr.doc.nodeAt(tablePos)!
  const map = TableMap.get(table)
  const start = tablePos + 1
  const seen = new Set<number>()
  for (let i = 0; i < map.map.length; i++) {
    const pos = map.map[i]!
    if (seen.has(pos)) continue
    seen.add(pos)
    const cell = table.nodeAt(pos)!
    const rect = map.findCell(pos)
    const colwidth = widths ? widths.slice(rect.left, rect.right) : null
    if (JSON.stringify(colwidth) === JSON.stringify(cell.attrs.colwidth)) continue
    tr.setNodeMarkup(start + pos, undefined, { ...cell.attrs, colwidth })
  }
  const total = pct.reduce((s, p) => s + p, 0) || 1
  tr.setNodeMarkup(tablePos, undefined, {
    ...table.attrs,
    colWidthsPct: pct.map((p) => (p / total) * 100),
    widthPx:
      widths && table.attrs.widthPx != null
        ? widths.reduce((s, w) => s + w, 0)
        : table.attrs.widthPx,
  })
}

export function pctOf(table: PmNode, map: TableMap, widths: number[] | null): number[] {
  if (widths) {
    const total = widths.reduce((s, w) => s + w, 0) || 1
    return widths.map((w) => (w / total) * 100)
  }
  const pct = table.attrs.colWidthsPct as number[] | null
  return pct?.length === map.width
    ? [...pct]
    : Array.from({ length: map.width }, () => 100 / map.width)
}

/** prosemirror-tables' addColumn maps through the whole transaction; this one only through its own steps */
export function insertColumnAt(tr: Transaction, tablePos: number, col: number): void {
  const table = tr.doc.nodeAt(tablePos)!
  const map = TableMap.get(table)
  const start = tablePos + 1
  const mapStart = tr.mapping.maps.length
  const refColumn = col > 0 ? -1 : 0
  const cellType = tableNodeTypes(table.type.schema).cell
  for (let row = 0; row < map.height; row++) {
    const index = row * map.width + col
    if (col > 0 && col < map.width && map.map[index - 1] === map.map[index]) {
      const pos = map.map[index]!
      const cell = table.nodeAt(pos)!
      tr.setNodeMarkup(
        tr.mapping.slice(mapStart).map(start + pos),
        undefined,
        addColSpan(cell.attrs as Parameters<typeof addColSpan>[0], col - map.colCount(pos)),
      )
      row += Number(cell.attrs.rowspan) - 1
    } else {
      const refPos = map.map[index + refColumn]
      const type = refPos === undefined ? cellType : table.nodeAt(refPos)!.type
      const pos = map.positionAt(row, col, table)
      tr.insert(tr.mapping.slice(mapStart).map(start + pos), type.createAndFill()!)
    }
  }
}

const COPIED_CELL_ATTRS = [
  'fill',
  'borders',
  'vAlign',
  'cellMar',
  'textDirection',
  'noWrap',
  'align',
  'bold',
  'color',
] as const

/** new rows/columns take the visual attrs of the neighbouring cell, like Word */
export function copyCellFormatting(
  tr: Transaction,
  tablePos: number,
  isNew: (rect: Rect) => boolean,
  refOf: (rect: Rect) => [number, number],
): void {
  const table = tr.doc.nodeAt(tablePos)!
  const map = TableMap.get(table)
  const start = tablePos + 1
  const seen = new Set<number>()
  for (let i = 0; i < map.map.length; i++) {
    const pos = map.map[i]!
    if (seen.has(pos)) continue
    seen.add(pos)
    const rect = map.findCell(pos)
    if (!isNew(rect)) continue
    const [refRow, refCol] = refOf(rect)
    if (refRow < 0 || refRow >= map.height || refCol < 0 || refCol >= map.width) continue
    const ref = table.nodeAt(map.map[refRow * map.width + refCol]!)!
    const cell = table.nodeAt(pos)!
    // a new cell has no raw tcPr: the copied direction / margins / noWrap only reach the file when written from the attrs
    const patch: Record<string, unknown> = { tcPrEdited: true }
    for (const k of COPIED_CELL_ATTRS) patch[k] = ref.attrs[k]
    tr.setNodeMarkup(start + pos, undefined, { ...cell.attrs, ...patch })
  }
}

/** prosemirror-tables' addRow, but positions are read from the current doc, not mapped through earlier batch steps */
export function insertRowAt(tr: Transaction, tablePos: number, row: number): void {
  const table = tr.doc.nodeAt(tablePos)!
  const map = TableMap.get(table)
  const start = tablePos + 1
  let rowPos = start
  for (let i = 0; i < row; i++) rowPos += table.child(i).nodeSize
  const cells: PmNode[] = []
  let refRow: number | null = row > 0 ? -1 : 0
  if (rowIsHeader(map, table, row + refRow)) refRow = row === 0 || row === map.height ? null : 0
  const types = tableNodeTypes(table.type.schema)
  for (let col = 0; col < map.width; col++) {
    const index = map.width * row + col
    if (row > 0 && row < map.height && map.map[index] === map.map[index - map.width]) {
      const pos = map.map[index]!
      const attrs = table.nodeAt(pos)!.attrs
      tr.setNodeMarkup(start + pos, undefined, { ...attrs, rowspan: Number(attrs.rowspan) + 1 })
      col += Number(attrs.colspan) - 1
    } else {
      const ref = refRow === null ? undefined : map.map[index + refRow * map.width]
      const type = ref === undefined ? types.cell : table.nodeAt(ref)!.type
      cells.push(type.createAndFill()!)
    }
  }
  if (cells.length === 0) {
    throw new Error(`every column at row ${row} is covered by vertically merged cells`)
  }
  tr.insert(rowPos, types.row.create(null, cells))
}
