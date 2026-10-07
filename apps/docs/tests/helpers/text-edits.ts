import type { Editor } from '@tiptap/core'

/**
 * Replace every occurrence of `find` inside single text nodes, keeping the
 * matched node's marks (the way a typed replacement inherits its run
 * formatting). A replacement equal to the match leaves the node untouched.
 * Returns how many occurrences changed.
 */
export function replaceText(editor: Editor, find: string, replace: string): number {
  const hits: Array<{ from: number; to: number; marks: readonly unknown[] }> = []
  editor.state.doc.descendants((node, pos) => {
    if (!node.isText || !node.text) return true
    let at = node.text.indexOf(find)
    while (at >= 0) {
      hits.push({ from: pos + at, to: pos + at + find.length, marks: node.marks })
      at = node.text.indexOf(find, at + find.length)
    }
    return false
  })
  if (find === replace || hits.length === 0) return 0
  const { schema } = editor
  const tr = editor.state.tr
  for (const hit of hits.reverse()) {
    if (replace) {
      tr.replaceWith(hit.from, hit.to, schema.text(replace, hit.marks as never))
    } else {
      tr.delete(hit.from, hit.to)
    }
  }
  editor.view.dispatch(tr)
  return hits.length
}

/** Move one top-level block so it lands after block `afterIndex` (indexes in the current doc). */
export function moveBlock(editor: Editor, index: number, afterIndex: number): void {
  const doc = editor.state.doc
  const posOf = (i: number) => {
    let pos = 0
    for (let k = 0; k < i; k++) pos += doc.child(k).nodeSize
    return pos
  }
  const node = doc.child(index)
  const from = posOf(index)
  const tr = editor.state.tr.delete(from, from + node.nodeSize)
  tr.insert(tr.mapping.map(posOf(afterIndex + 1)), node)
  editor.view.dispatch(tr)
}
