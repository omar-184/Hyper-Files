import { describe, expect, it } from 'vitest'
import {
  indentAt,
  insertAfter,
  moveAt,
  outdentAt,
  removeAt,
  renameAt,
  toOutlineInput,
} from '../src/renderer/outline-edit'
import type { OutlineNode } from '../src/renderer/OutlinePanel'

const n = (title: string, items: OutlineNode[] = []): OutlineNode => ({
  title,
  dest: [0, { name: 'XYZ' }, null, 700, null],
  items,
})
const titles = (nodes: OutlineNode[]): unknown =>
  nodes.map((x) => (x.items?.length ? [x.title, titles(x.items)] : x.title))

describe('outline tree edits', () => {
  const tree = [n('A', [n('A1'), n('A2')]), n('B'), n('C')]

  it('inserts after a path or at the end', () => {
    expect(insertAfter(tree, [0, 0], n('new'))).toEqual([
      [n('A', [n('A1'), n('new'), n('A2')]), n('B'), n('C')],
      [0, 1],
    ])
    const [end, path] = insertAfter(tree, null, n('Z'))
    expect(titles(end)).toEqual([['A', ['A1', 'A2']], 'B', 'C', 'Z'])
    expect(path).toEqual([3])
  })

  it('removes, renames and moves within a level', () => {
    expect(titles(removeAt(tree, [0, 1]))).toEqual([['A', ['A1']], 'B', 'C'])
    expect(renameAt(tree, [1], 'Bee')[1]!.title).toBe('Bee')
    const [moved, path] = moveAt(tree, [2], -1)!
    expect(titles(moved)).toEqual([['A', ['A1', 'A2']], 'C', 'B'])
    expect(path).toEqual([1])
    expect(moveAt(tree, [0], -1)).toBeNull()
    expect(moveAt(tree, [0, 1], 1)).toBeNull()
  })

  it('indents under the previous sibling and outdents after the parent', () => {
    const [indented, path] = indentAt(tree, [1])!
    expect(titles(indented)).toEqual([['A', ['A1', 'A2', 'B']], 'C'])
    expect(path).toEqual([0, 2])
    expect(indentAt(tree, [0])).toBeNull()
    const [out, outPath] = outdentAt(indented, [0, 0])!
    expect(titles(out)).toEqual([['A', ['A2', 'B']], 'A1', 'C'])
    expect(outPath).toEqual([1])
    expect(outdentAt(tree, [1])).toBeNull()
  })

  it('serializes destinations, links and colors for the save request', () => {
    expect(
      toOutlineInput([
        { ...n('Page'), bold: true, color: new Uint8ClampedArray([255, 0, 0]) },
        { title: 'Web', url: 'https://example.com/', items: [] },
        { title: 'Lost', dest: null, items: [] },
      ]),
    ).toEqual([
      {
        title: 'Page',
        pageIndex: 0,
        fit: 'XYZ',
        args: [null, 700, null],
        bold: true,
        color: [1, 0, 0],
        items: [],
      },
      { title: 'Web', pageIndex: null, url: 'https://example.com/', items: [] },
      { title: 'Lost', pageIndex: null, items: [] },
    ])
  })
})
