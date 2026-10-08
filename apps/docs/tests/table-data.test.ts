import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection, type Command } from '@tiptap/pm/state'
import { TableMap } from '@tiptap/pm/tables'
import type { Node as PmNode } from '@tiptap/pm/model'
import { parseDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  blocksToPmDoc,
  pmDocToSavePlan,
  type PmNode as JsonNode,
} from '../src/renderer/editor/convert'
import {
  convertTableToText,
  convertTextToTable,
  dateKey,
  guessHasHeader,
  numberKey,
  sortRefusal,
  sortTableRows,
  textToTablePlan,
} from '../src/renderer/editor/table-data'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const editors = new Set<Editor>()
afterEach(() => {
  for (const e of editors) e.destroy()
  editors.clear()
})

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;')
const run = (text: string, rPr = '') =>
  `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${esc(text)}</w:t></w:r>`
const cell = (text: string, rPr = '') => `<w:tc><w:p>${run(text, rPr)}</w:p></w:tc>`

function tableXml(rows: string[][], trPr: Record<number, string> = {}): string {
  const grid = rows[0].map(() => '<w:gridCol w:w="2000"/>').join('')
  const body = rows
    .map((r, i) => `<w:tr>${trPr[i] ?? ''}${r.map((t) => cell(t)).join('')}</w:tr>`)
    .join('')
  return `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr><w:tblGrid>${grid}</w:tblGrid>${body}</w:tbl>`
}

async function open(bodyXml: string) {
  const parsed = await parseDocx(await buildDocx({ bodyXml: '<w:p/>' + bodyXml + '<w:p/>' }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  editors.add(editor)
  return { editor, parsed }
}

function tableAt(editor: Editor): { node: PmNode; pos: number } | null {
  let found: { node: PmNode; pos: number } | null = null
  editor.state.doc.forEach((node, offset) => {
    if (!found && node.type.name === 'docTable') found = { node, pos: offset }
  })
  return found
}

function grid(editor: Editor): string[][] {
  const { node } = tableAt(editor)!
  const map = TableMap.get(node)
  return Array.from({ length: map.height }, (_, r) =>
    Array.from(
      { length: map.width },
      (_, c) => node.nodeAt(map.map[r * map.width + c])!.textContent,
    ),
  )
}

function caretInTable(editor: Editor, row = 0, col = 0) {
  const { node, pos } = tableAt(editor)!
  const map = TableMap.get(node)
  const at = pos + 1 + map.map[row * map.width + col] + 1
  editor.view.dispatch(
    editor.state.tr.setSelection(TextSelection.near(editor.state.doc.resolve(at))),
  )
}

const exec = (editor: Editor, command: Command) => command(editor.state, editor.view.dispatch)

function paragraphs(editor: Editor): string[] {
  const out: string[] = []
  editor.state.doc.forEach((node) =>
    out.push(node.type.name === 'docTable' ? '[table]' : node.textContent),
  )
  return out
}

describe('sort keys', () => {
  it('reads numbers with separators, currency and accounting negatives', () => {
    expect(numberKey('$1,250.50')).toBe(1250.5)
    expect(numberKey('(300)')).toBe(-300)
    expect(numberKey('-4')).toBe(-4)
    expect(numberKey('about 12 kg')).toBe(12)
    expect(numberKey('n/a')).toBeNull()
  })

  it('reads ISO, month-name and numeric dates', () => {
    expect(dateKey('2024-03-05')).toBe(Date.UTC(2024, 2, 5))
    expect(dateKey('3/5/2024')).toBe(Date.UTC(2024, 2, 5))
    expect(dateKey('25/12/2023')).toBe(Date.UTC(2023, 11, 25))
    expect(dateKey('March 5, 2024')).not.toBeNull()
    expect(dateKey('soon')).toBeNull()
  })
})

describe('sortTableRows', () => {
  const people = [
    ['Name', 'Age', 'Joined'],
    ['carol', '41', '2021-06-01'],
    ['Alice', '9', '2023-01-15'],
    ['bob', '120', '2019-11-30'],
  ]

  it('sorts text case-insensitively below the header row', async () => {
    const { editor } = await open(tableXml(people))
    caretInTable(editor, 1, 0)
    expect(guessHasHeader(editor.state)).toBe(false)
    expect(
      exec(
        editor,
        sortTableRows({ hasHeader: true, keys: [{ column: 0, type: 'text', descending: false }] }),
      ),
    ).toBe(true)
    expect(grid(editor).map((r) => r[0])).toEqual(['Name', 'Alice', 'bob', 'carol'])
  })

  it('sorts numbers numerically and descending', async () => {
    const { editor } = await open(tableXml(people))
    caretInTable(editor)
    exec(
      editor,
      sortTableRows({ hasHeader: true, keys: [{ column: 1, type: 'number', descending: true }] }),
    )
    expect(grid(editor).map((r) => r[1])).toEqual(['Age', '120', '41', '9'])
  })

  it('sorts dates and keeps whole rows together', async () => {
    const { editor } = await open(tableXml(people))
    caretInTable(editor)
    exec(
      editor,
      sortTableRows({ hasHeader: true, keys: [{ column: 2, type: 'date', descending: false }] }),
    )
    expect(grid(editor).slice(1)).toEqual([
      ['bob', '120', '2019-11-30'],
      ['carol', '41', '2021-06-01'],
      ['Alice', '9', '2023-01-15'],
    ])
  })

  it('uses a second key for ties and puts empty cells last', async () => {
    const { editor } = await open(
      tableXml([
        ['b', '2'],
        ['', '1'],
        ['a', '3'],
        ['b', '1'],
      ]),
    )
    caretInTable(editor)
    exec(
      editor,
      sortTableRows({
        hasHeader: false,
        keys: [
          { column: 0, type: 'text', descending: false },
          { column: 1, type: 'number', descending: false },
        ],
      }),
    )
    expect(grid(editor)).toEqual([
      ['a', '3'],
      ['b', '1'],
      ['b', '2'],
      ['', '1'],
    ])
  })

  it('keeps every repeat-header row in place', async () => {
    const header = '<w:trPr><w:tblHeader/></w:trPr>'
    const { editor } = await open(
      tableXml(
        [
          ['Title', 'x'],
          ['Sub', 'y'],
          ['z', '1'],
          ['a', '2'],
        ],
        { 0: header, 1: header },
      ),
    )
    caretInTable(editor)
    expect(guessHasHeader(editor.state)).toBe(true)
    exec(
      editor,
      sortTableRows({ hasHeader: true, keys: [{ column: 0, type: 'text', descending: false }] }),
    )
    expect(grid(editor).map((r) => r[0])).toEqual(['Title', 'Sub', 'a', 'z'])
  })

  it('refuses tables with vertically merged cells', async () => {
    const merged =
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
      `<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p>${run('a')}</w:p></w:tc>${cell('1')}</w:tr>` +
      `<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>${cell('2')}</w:tr>` +
      `<w:tr>${cell('b')}${cell('3')}</w:tr></w:tbl>`
    const { editor } = await open(merged)
    caretInTable(editor, 2, 0)
    expect(sortRefusal(editor.state, false)).toBe('mergedRows')
    expect(
      exec(
        editor,
        sortTableRows({ hasHeader: false, keys: [{ column: 0, type: 'text', descending: false }] }),
      ),
    ).toBe(false)
  })

  it('saves the sorted order to the docx and undoes in one step', async () => {
    const { editor, parsed } = await open(tableXml(people))
    caretInTable(editor)
    exec(
      editor,
      sortTableRows({ hasHeader: true, keys: [{ column: 0, type: 'text', descending: false }] }),
    )
    const plan = pmDocToSavePlan(editor.getJSON() as JsonNode, parsed.blocks)
    const xml = plan.saveBlocks
      .map((b) => (b.kind === 'xml' ? b.xml : ''))
      .find((x) => x.includes('<w:tbl>'))!
    const order = ['Name', 'Alice', 'bob', 'carol'].map((s) => xml.indexOf(`>${s}<`))
    expect(order.every((i) => i > 0)).toBe(true)
    expect([...order].sort((a, b) => a - b)).toEqual(order)
    editor.commands.undo()
    expect(grid(editor).map((r) => r[0])).toEqual(['Name', 'carol', 'Alice', 'bob'])
  })
})

describe('convertTableToText', () => {
  it('joins cells with tabs, one paragraph per row, keeping bold runs', async () => {
    const bold =
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
      `<w:tr>${cell('Head', '<w:b/>')}${cell('Two')}</w:tr><w:tr>${cell('a')}${cell('b')}</w:tr></w:tbl>`
    const { editor } = await open(bold)
    caretInTable(editor)
    expect(exec(editor, convertTableToText('tab'))).toBe(true)
    expect(paragraphs(editor)).toEqual(['', 'Head\tTwo', 'a\tb', ''])
    const first = editor.state.doc.child(1).firstChild!
    expect(first.marks.some((m) => m.type.name === 'bold')).toBe(true)
  })

  it('puts each cell in its own paragraph with the paragraph-mark separator', async () => {
    const { editor } = await open(
      tableXml([
        ['a', 'b'],
        ['c', 'd'],
      ]),
    )
    caretInTable(editor)
    exec(editor, convertTableToText('paragraph'))
    expect(paragraphs(editor)).toEqual(['', 'a', 'b', 'c', 'd', ''])
  })

  it('uses a custom separator character', async () => {
    const { editor } = await open(tableXml([['a', 'b', 'c']]))
    caretInTable(editor)
    exec(editor, convertTableToText({ other: ';' }))
    expect(paragraphs(editor)).toEqual(['', 'a;b;c', ''])
  })
})

describe('convertTextToTable', () => {
  async function openParagraphs(lines: string[]) {
    const { editor, parsed } = await open(lines.map((l) => `<w:p>${run(l)}</w:p>`).join(''))
    // select from the start of the first line to the end of the last one (doc starts with an empty paragraph)
    const first = 2 + 1
    let end = 0
    editor.state.doc.forEach((node, offset) => {
      end = node.textContent && offset + node.nodeSize > end ? offset + node.nodeSize - 1 : end
    })
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, first, end)),
    )
    return { editor, parsed }
  }

  it('splits tab-separated lines into rows and columns', async () => {
    const { editor, parsed } = await openParagraphs(['Name\tAge', 'Ann\t30', 'Bo\t4\textra'])
    expect(textToTablePlan(editor.state, 'tab')).toEqual({ rows: 3, cols: 3 })
    expect(exec(editor, convertTextToTable('tab'))).toBe(true)
    expect(grid(editor)).toEqual([
      ['Name', 'Age', ''],
      ['Ann', '30', ''],
      ['Bo', '4', 'extra'],
    ])
    const plan = pmDocToSavePlan(editor.getJSON() as JsonNode, parsed.blocks)
    const xml = plan.saveBlocks.map((b) => (b.kind === 'xml' ? b.xml : '')).join('')
    expect(xml).toContain('<w:tbl>')
    expect(xml).toContain('extra')
  })

  it('splits on commas', async () => {
    const { editor } = await openParagraphs(['a,b', 'c,d'])
    exec(editor, convertTextToTable('comma'))
    expect(grid(editor)).toEqual([
      ['a', 'b'],
      ['c', 'd'],
    ])
  })

  it('fills one paragraph per cell, row by row, for the paragraph separator', async () => {
    const { editor } = await openParagraphs(['1', '2', '3'])
    exec(editor, convertTextToTable('paragraph', 2))
    expect(grid(editor)).toEqual([
      ['1', '2'],
      ['3', ''],
    ])
  })

  it('round-trips with Convert to Text', async () => {
    const { editor } = await openParagraphs(['x\ty', 'z\tw'])
    exec(editor, convertTextToTable('tab'))
    caretInTable(editor)
    exec(editor, convertTableToText('tab'))
    expect(paragraphs(editor)).toEqual(['', 'x\ty', 'z\tw', ''])
  })

  it('is unavailable inside a table', async () => {
    const { editor } = await open(tableXml([['a', 'b']]))
    caretInTable(editor)
    expect(textToTablePlan(editor.state, 'tab')).toBeNull()
  })
})
