import type { PDFDocumentProxy } from 'pdfjs-dist'
import type { OutlineEntryInput } from '../shared/ipc'
import { MAX_OUTLINE_DEPTH, type OutlineNode } from './OutlinePanel'

/** Position of an entry: child indices from the top level down */
export type OutlinePath = number[]

type PageRef = Parameters<PDFDocumentProxy['getPageIndex']>[0]

/**
 * Editable copy of the file's outline: every destination resolved to an
 * explicit [pageIndex, {name: fit}, ...args] array, the same shape generated
 * outlines use, so the tree can be saved without the document's references.
 * Unresolvable destinations become null (the entry keeps its title).
 */
export async function normalizeOutline(
  doc: PDFDocumentProxy,
  nodes: readonly OutlineNode[],
  depth = 0,
): Promise<OutlineNode[]> {
  if (depth >= MAX_OUTLINE_DEPTH) return []
  return Promise.all(
    nodes.map(async (node): Promise<OutlineNode> => {
      let dest: unknown = null
      if (!node.url && node.dest != null) {
        try {
          const arr =
            typeof node.dest === 'string' ? await doc.getDestination(node.dest) : node.dest
          if (Array.isArray(arr) && arr.length > 0) {
            const ref = arr[0]
            const page = typeof ref === 'number' ? ref : await doc.getPageIndex(ref as PageRef)
            dest = [page, ...arr.slice(1)]
          }
        } catch {
          dest = null // dangling destination: keep the entry, drop the jump
        }
      }
      return {
        title: node.title,
        ...(node.bold ? { bold: true } : {}),
        ...(node.italic ? { italic: true } : {}),
        ...(node.color ? { color: node.color } : {}),
        ...(node.url ? { url: node.url } : { dest }),
        items: await normalizeOutline(doc, node.items ?? [], depth + 1),
      }
    }),
  )
}

/** A bookmark to the top of a page (original index) */
export function pageBookmark(title: string, pageIndex: number, pageHeight: number): OutlineNode {
  return { title, dest: [pageIndex, { name: 'XYZ' }, null, pageHeight, null], items: [] }
}

/** Save payload for a normalized tree */
export function toOutlineInput(nodes: readonly OutlineNode[]): OutlineEntryInput[] {
  return nodes.map((node) => {
    const dest = Array.isArray(node.dest) ? node.dest : null
    const fit = dest?.[1] as { name?: unknown } | undefined
    const color = node.color ? Array.from(node.color, (c) => c / 255) : null
    return {
      title: node.title,
      pageIndex: dest && typeof dest[0] === 'number' ? dest[0] : null,
      ...(dest && typeof fit?.name === 'string'
        ? {
            fit: fit.name,
            args: dest.slice(2).map((v) => (typeof v === 'number' ? v : null)),
          }
        : {}),
      ...(node.url ? { url: node.url } : {}),
      ...(node.bold ? { bold: true } : {}),
      ...(node.italic ? { italic: true } : {}),
      ...(color && color.length === 3 ? { color: color as [number, number, number] } : {}),
      items: toOutlineInput(node.items ?? []),
    }
  })
}

export function nodeAt(nodes: readonly OutlineNode[], path: OutlinePath): OutlineNode | null {
  let list = nodes
  let node: OutlineNode | null = null
  for (const i of path) {
    node = list[i] ?? null
    if (!node) return null
    list = node.items ?? []
  }
  return node
}

/** Copy of the tree with the sibling list holding `path` replaced by fn(list, index) */
function editSiblings(
  nodes: readonly OutlineNode[],
  path: OutlinePath,
  fn: (list: OutlineNode[], index: number) => OutlineNode[],
): OutlineNode[] {
  const [head, ...rest] = path
  if (head === undefined) return [...nodes]
  if (rest.length === 0) return fn([...nodes], head)
  return nodes.map((n, i) =>
    i === head ? { ...n, items: editSiblings(n.items ?? [], rest, fn) } : n,
  )
}

/** Insert after `path` on the same level (top-level end when path is null); returns the new path */
export function insertAfter(
  nodes: readonly OutlineNode[],
  path: OutlinePath | null,
  node: OutlineNode,
): [OutlineNode[], OutlinePath] {
  if (!path || path.length === 0 || !nodeAt(nodes, path)) return [[...nodes, node], [nodes.length]]
  const at = path[path.length - 1]! + 1
  return [
    editSiblings(nodes, path, (list) => [...list.slice(0, at), node, ...list.slice(at)]),
    [...path.slice(0, -1), at],
  ]
}

export function removeAt(nodes: readonly OutlineNode[], path: OutlinePath): OutlineNode[] {
  return editSiblings(nodes, path, (list, i) => list.filter((_, j) => j !== i))
}

export function renameAt(
  nodes: readonly OutlineNode[],
  path: OutlinePath,
  title: string,
): OutlineNode[] {
  return editSiblings(nodes, path, (list, i) => list.map((n, j) => (j === i ? { ...n, title } : n)))
}

/** Swap with the previous (-1) or next (+1) sibling; null when already at that end */
export function moveAt(
  nodes: readonly OutlineNode[],
  path: OutlinePath,
  delta: -1 | 1,
): [OutlineNode[], OutlinePath] | null {
  const i = path[path.length - 1]!
  const parent = path.length > 1 ? nodeAt(nodes, path.slice(0, -1)) : null
  const siblings = parent ? (parent.items ?? []) : nodes
  const j = i + delta
  if (j < 0 || j >= siblings.length) return null
  return [
    editSiblings(nodes, path, (list) => {
      const next = [...list]
      ;[next[i], next[j]] = [next[j]!, next[i]!]
      return next
    }),
    [...path.slice(0, -1), j],
  ]
}

/** Make the entry the last child of its previous sibling */
export function indentAt(
  nodes: readonly OutlineNode[],
  path: OutlinePath,
): [OutlineNode[], OutlinePath] | null {
  const i = path[path.length - 1]!
  if (i === 0 || path.length >= MAX_OUTLINE_DEPTH) return null
  const node = nodeAt(nodes, path)!
  let newIndex = 0
  const next = editSiblings(nodes, path, (list) => {
    const prev = list[i - 1]!
    newIndex = prev.items?.length ?? 0
    const moved = { ...prev, items: [...(prev.items ?? []), node] }
    return [...list.slice(0, i - 1), moved, ...list.slice(i + 1)]
  })
  return [next, [...path.slice(0, -1), i - 1, newIndex]]
}

/** Move the entry out of its parent, right after it */
export function outdentAt(
  nodes: readonly OutlineNode[],
  path: OutlinePath,
): [OutlineNode[], OutlinePath] | null {
  if (path.length < 2) return null
  const node = nodeAt(nodes, path)!
  const parentPath = path.slice(0, -1)
  const without = removeAt(nodes, path)
  return insertAfter(without, parentPath, node)
}
