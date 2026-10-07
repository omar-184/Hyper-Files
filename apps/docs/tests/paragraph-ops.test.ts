import { afterEach, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { runUiOps } from '../src/renderer/editor/paragraph-ops'

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
  marks?: Array<{ type: string; attrs?: Record<string, unknown> }>
}

const text = (t: string, marks?: JsonNode['marks']): JsonNode => ({
  type: 'text',
  text: t,
  ...(marks && marks.length > 0 ? { marks } : {}),
})

const heading = (
  content: JsonNode[],
  level = 1,
  attrs: Record<string, unknown> = {},
): JsonNode => ({
  type: 'docHeading',
  attrs: { docxIndex: null, level, ...attrs },
  content,
})

const para = (content: JsonNode[], attrs: Record<string, unknown> = {}): JsonNode => ({
  type: 'docParagraph',
  attrs: { docxIndex: null, ...attrs },
  content,
})

const listItem = (content: JsonNode[], attrs: Record<string, unknown> = {}): JsonNode => ({
  type: 'docListItem',
  attrs: { docxIndex: null, kind: 'bullet', ...attrs },
  content,
})

const protectedTable = (attrs: Record<string, unknown> = {}): JsonNode => ({
  type: 'docProtected',
  attrs: {
    docxIndex: 90,
    blockType: 'table',
    label: 'Table 2×2',
    previewText: 'City GDP',
    ...attrs,
  },
})

const editors = new Set<Editor>()

afterEach(() => {
  for (const editor of editors) editor.destroy()
  editors.clear()
})

function createEditor(content: JsonNode[]): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
  editors.add(editor)
  return editor
}

/** standard fixture: 0 h1 | 1 p | 2 h2 | 3 p | 4 li | 5 protected table */
function fixtureDoc(): JsonNode[] {
  return [
    heading(
      [
        text('Chapter 1 Overview', [
          { type: 'bold' },
          { type: 'docTextStyle', attrs: { sizeHalfPoints: 32 } },
        ]),
      ],
      1,
      {
        docxIndex: 0,
      },
    ),
    para([text('GenSpark intro,'), text('GenSpark is great', [{ type: 'bold' }])], {
      docxIndex: 1,
    }),
    heading([text('Risk Notes')], 2, { docxIndex: 2 }),
    para([text('Body paragraph')], { docxIndex: 3, align: 'center' }),
    listItem([text('List item')], { docxIndex: 4, numId: '1' }),
    protectedTable({ docxIndex: 5 }),
  ]
}

describe('paragraph ops', () => {
  const cellTable = (): JsonNode => ({
    type: 'docTable',
    attrs: { docxIndex: null },
    content: [
      {
        type: 'docTableRow',
        content: [
          {
            type: 'docTableCell',
            content: [{ type: 'docParagraph', attrs: { docxIndex: null }, content: [text('A1')] }],
          },
          {
            type: 'docTableCell',
            content: [{ type: 'docParagraph', attrs: { docxIndex: null }, content: [text('B1')] }],
          },
        ],
      },
    ],
  })
  const cellPara = (editor: Editor, col: number) =>
    editor.state.doc.child(1).child(0).child(col).child(0)

  it('setParagraphAttrs patches the targeted block', () => {
    const editor = createEditor(fixtureDoc())
    expect(
      runUiOps(editor, [
        { op: 'setParagraphAttrs', target: { blockIndexes: [1] }, attrs: { align: 'right' } },
      ]),
    ).toBe(true)
    expect(editor.state.doc.child(1).attrs.align).toBe('right')
    expect(editor.state.doc.child(0).attrs.align ?? null).toBeNull()
  })

  it('a range target stands in for the selection (blur-committed dialog inputs)', () => {
    const editor = createEditor(fixtureDoc())
    const block3Pos =
      editor.state.doc.child(0).nodeSize +
      editor.state.doc.child(1).nodeSize +
      editor.state.doc.child(2).nodeSize
    editor.commands.setTextSelection(2) // live selection in block 0
    runUiOps(
      editor,
      [
        {
          op: 'setParagraphAttrs',
          target: { range: { from: block3Pos + 2, to: block3Pos + 4 } },
          attrs: { spaceBefore: 240 },
        },
      ],
      { focus: false },
    )
    expect(editor.state.doc.child(3).attrs.spaceBefore).toBe(240)
    // an explicit spacing value turns Word's "Auto" spacing off
    expect(editor.state.doc.child(3).attrs.spaceBeforeAuto).toBe(false)
    expect(editor.state.doc.child(0).attrs.spaceBefore).toBeNull()
  })

  it('a selection target inside a table reaches only the selected cell', () => {
    const ui = createEditor([heading([text('T')]), cellTable()])
    // caret inside cell B1: table(1) + row(1) + cellA(A1 para = 4, cell = 6) + cellB(1) + para(1)
    const b1 = ui.state.doc.child(0).nodeSize + 1 + 1 + 6 + 1 + 1 + 1
    expect(ui.state.doc.textBetween(b1 - 1, b1 + 1)).toBe('B1')
    ui.commands.setTextSelection(b1)
    runUiOps(ui, [
      { op: 'setParagraphAttrs', target: { scope: 'selection' }, attrs: { align: 'right' } },
    ])
    expect(cellPara(ui, 1).attrs.align).toBe('right')
    expect(cellPara(ui, 0).attrs.align).toBeNull()
  })

  it('alignment also lands on selected images as their w:jc', () => {
    const image = (): JsonNode => ({
      type: 'docProtected',
      attrs: {
        docxIndex: 7,
        blockType: 'image',
        label: 'Image',
        imageWidthPx: 400,
        imageHeightPx: 200,
      },
    })
    const ui = createEditor([para([text('a')]), image(), para([text('b')])])
    runUiOps(ui, [
      {
        op: 'setParagraphAttrs',
        target: { range: { from: 0, to: ui.state.doc.content.size } },
        attrs: { align: 'center' },
      },
    ])
    expect(ui.state.doc.child(1).attrs.imageAlign).toBe('center')
    expect(ui.state.doc.child(0).attrs.align).toBe('center')
  })

  it('stepIndent treats each paragraph on its own: list items change level, paragraphs snap to half-inch stops', () => {
    const editor = createEditor([
      listItem([text('item')], { numId: '1', ilvl: 1 }),
      para([text('plain')], { indentLeft: 500 }),
    ])
    runUiOps(editor, [
      {
        op: 'stepIndent',
        target: { range: { from: 0, to: editor.state.doc.content.size } },
        delta: 1,
      },
    ])
    expect(editor.state.doc.child(0).attrs.ilvl).toBe(2)
    expect(editor.state.doc.child(1).attrs.indentLeft).toBe(720)
    runUiOps(editor, [{ op: 'stepHangingIndent', target: { blockIndexes: [1] }, delta: 1 }])
    expect(editor.state.doc.child(1).attrs.indentLeft).toBe(1440)
    expect(editor.state.doc.child(1).attrs.indentFirstLine).toBe(-720)
    const error = vi.spyOn(console, 'error').mockImplementation(() => {})
    expect(runUiOps(editor, [{ op: 'stepIndent', target: { blockIndexes: [1] }, delta: 2 }])).toBe(
      false,
    )
    expect(String(error.mock.calls[0]?.[1])).toContain('delta must be 1 or -1')
    error.mockRestore()
  })
})
