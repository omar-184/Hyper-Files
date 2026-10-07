import { describe, expect, it } from 'vitest'
import type { TextBlock } from '../src/renderer/text-block'
import {
  blockMoveInput,
  patchPendingEdits,
  shiftPendingEdit,
} from '../src/renderer/text-edit-preview'
import type { LocalTextEdit } from '../src/renderer/text-edit-preview'

const block: TextBlock = {
  rect: [50, 676, 250, 712],
  fontSize: 10,
  lineHeight: 12,
  align: 'left',
  lines: [
    { text: 'first line', y: 700, rect: [50, 698, 250, 712], fontSize: 10 },
    { text: 'second line', y: 688, rect: [50, 686, 250, 700], fontSize: 10 },
    { text: 'third', y: 676, rect: [50, 674, 250, 688], fontSize: 10 },
  ],
}

describe('blockMoveInput', () => {
  it('builds a pure translate edit anchored at the shifted block corner', () => {
    const input = blockMoveInput(0, block, [5, -30])
    expect(input.translate).toEqual([5, -30])
    expect(input.origin).toEqual([55, 670])
    expect(input.rect).toEqual(block.rect)
    expect(input.newText).toBe('first line\nsecond line\nthird')
    expect(input.oldText).toBe(input.blockSource)
    expect(input.align).toBeUndefined()
  })
})

describe('shiftPendingEdit', () => {
  it('stacks a second move onto a pending pure move', () => {
    const te: LocalTextEdit = {
      id: 'a',
      input: blockMoveInput(0, block, [0, -30]),
      moveBy: [0, -30],
    }
    const shifted = shiftPendingEdit(te, block, [5, -10])
    expect(shifted.input.translate).toEqual([5, -40])
    expect(shifted.input.origin).toEqual([55, 660])
    expect(shifted.moveBy).toEqual([5, -40])
    expect(shifted.input.rect).toEqual(block.rect)
    expect(te.input.translate).toEqual([0, -30])
  })

  it('moves a pending paragraph rebuild by its origin without adding a translate', () => {
    const te: LocalTextEdit = {
      id: 'b',
      input: {
        pageIndex: 0,
        rect: block.rect,
        oldText: 'first line second line third',
        newText: 'rewritten',
        fontSize: 10,
        origin: [50, 700],
        lineLeading: 12,
      },
    }
    const shifted = shiftPendingEdit(te, block, [0, -30])
    expect(shifted.input.origin).toEqual([50, 670])
    expect(shifted.input.translate).toBeUndefined()
    expect(shifted.moveBy).toEqual([0, -30])
  })

  it('anchors a pending line edit at its own shifted row', () => {
    const te: LocalTextEdit = {
      id: 'c',
      input: {
        pageIndex: 0,
        rect: [80, 686, 200, 700],
        oldText: 'second line',
        newText: 'second LINE',
        fontSize: 10,
      },
    }
    const shifted = shiftPendingEdit(te, block, [0, -30])
    expect(shifted.input.origin).toEqual([80, 658])
    expect(shifted.input.lineLeading).toBe(12)
    expect(shifted.input.translate).toBeUndefined()
  })
})

describe('patchPendingEdits', () => {
  const mk = (id: string, newText = id): LocalTextEdit => ({
    id,
    input: { pageIndex: 0, rect: [0, 0, 10, 10], oldText: 'o', newText, fontSize: 10 },
  })

  it('replaces an existing edit in place and keeps everything else', () => {
    const out = patchPendingEdits([mk('a'), mk('b'), mk('c')], mk('b', 'B'))
    expect(out.map((e) => e.id)).toEqual(['a', 'b', 'c'])
    expect(out[1]!.input.newText).toBe('B')
  })

  it('appends a new edit and drops the ids it supersedes', () => {
    const out = patchPendingEdits([mk('a'), mk('b')], mk('n'), new Set(['a']))
    expect(out.map((e) => e.id)).toEqual(['b', 'n'])
  })

  it('never removes the edit being landed even when listed', () => {
    const out = patchPendingEdits([mk('a')], mk('a', 'A'), new Set(['a']))
    expect(out.map((e) => e.input.newText)).toEqual(['A'])
  })

  it('in replace mode leaves the list untouched when the owner is gone', () => {
    const prev = [mk('b')]
    const out = patchPendingEdits(prev, mk('a', 'A'), new Set(['b']), 'replace')
    expect(out).toBe(prev)
    expect(patchPendingEdits([mk('a'), mk('b')], mk('a', 'A'), new Set(['b']), 'replace')).toEqual([
      mk('a', 'A'),
    ])
  })
})
