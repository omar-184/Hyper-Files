import { test, expect } from '@playwright/test'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, screenshotPath, waitForPageWithUrl } from './helpers'

/**
 * Word's table Data tools: Sort (number column, descending, header row kept),
 * Convert to Text (tabs) and Insert ▸ Table ▸ Convert Text to Table, saved as OOXML.
 */

interface AidocsWindow {
  __aidocs?: { editor?: unknown; save?: () => Promise<unknown> }
}

const ROWS = [
  ['Item', 'Qty'],
  ['Pens', '12'],
  ['Paper', '250'],
  ['Clips', '7'],
]

function tableXml(): string {
  const line = 'w:val="single" w:sz="4" w:space="0" w:color="auto"'
  const borders = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
    .map((side) => `<w:${side} ${line}/>`)
    .join('')
  const body = ROWS.map(
    (row) =>
      `<w:tr>${row.map((text) => `<w:tc><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`).join('')}</w:tr>`,
  ).join('')
  return `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblBorders>${borders}</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>${body}</w:tbl>`
}

async function minimalDocx(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  )
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Before</w:t></w:r></w:p>${tableXml()}<w:p><w:r><w:t>After</w:t></w:r></w:p></w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

async function documentXml(path: string): Promise<string> {
  const zip = await JSZip.loadAsync(readFileSync(path))
  return zip.file('word/document.xml')!.async('string')
}

test.describe('docs table data tools', () => {
  let dir: string
  let docPath: string

  test.beforeEach(async () => {
    dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-table-data-')))
    docPath = join(dir, 'stock.docx')
    writeFileSync(docPath, await minimalDocx())
  })

  test.afterEach(() => {
    rmSync(dir, { recursive: true, force: true })
  })

  test('sort, convert to text and convert text to table save as OOXML', async () => {
    test.setTimeout(180_000)
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-table-data',
      openFile: docPath,
    })
    const { app } = launched
    try {
      const page = await waitForPageWithUrl(app, '://docs/')
      await page.waitForFunction(
        () => Boolean((window as unknown as AidocsWindow).__aidocs?.editor),
        undefined,
        { timeout: 30_000 },
      )
      const table = page.locator('.doc-page .doc-table').first()
      const column = (col: number) =>
        table
          .locator('tr')
          .evaluateAll(
            (rows, c) => rows.map((row) => row.querySelectorAll('td, th')[c]?.textContent ?? ''),
            col,
          )

      // ---- Sort by Qty, number, descending; the header row stays on top
      await table.locator('tr').nth(1).locator('td, th').first().click()
      await page.locator('.ribbon-tab', { hasText: 'Table Layout' }).click()
      await page.getByRole('button', { name: 'Sort', exact: true }).click()
      const sort = page.locator('.modal.table-sort-dialog')
      await expect(sort).toBeVisible()
      await sort.getByLabel('My list has a header row').check()
      await sort.getByLabel('Sort by').selectOption({ label: 'Qty' })
      await sort.getByLabel('Type').first().selectOption('number')
      await sort.getByLabel('Descending').first().check()
      await sort.screenshot({ path: screenshotPath('docs-table-sort-dialog') })
      await sort.getByRole('button', { name: 'OK' }).click()
      await expect(sort).toBeHidden()
      await expect.poll(() => column(0)).toEqual(['Item', 'Paper', 'Pens', 'Clips'])

      // ---- Convert to Text with tabs
      await table.locator('tr').first().locator('td, th').first().click()
      await page.getByRole('button', { name: 'Convert to Text' }).click()
      const toText = page.locator('.modal.table-to-text-dialog')
      await expect(toText).toBeVisible()
      await toText.getByLabel('Tabs').check()
      await toText.getByRole('button', { name: 'OK' }).click()
      await expect(page.locator('.doc-page .doc-table')).toHaveCount(0)
      await expect(page.locator('.doc-page')).toContainText('Paper')

      // ---- select the four lines back and Convert Text to Table
      await page.evaluate(() => {
        const editor = (window as unknown as AidocsWindow).__aidocs!.editor as {
          state: { doc: { content: { size: number } } }
          commands: { setTextSelection: (r: { from: number; to: number }) => boolean }
        }
        // "Before" is the first paragraph (8 positions); stop before "After"
        const end = editor.state.doc.content.size - 'After'.length - 2
        editor.commands.setTextSelection({ from: 9, to: end })
      })
      await page.locator('.ribbon-tab', { hasText: 'Insert' }).click()
      await page.getByRole('button', { name: 'Table', exact: true }).click()
      await page.getByRole('button', { name: 'Convert Text to Table…' }).click()
      const toTable = page.locator('.modal.text-to-table-dialog')
      await expect(toTable).toBeVisible()
      await expect(toTable).toContainText('Number of rows: 4')
      await expect(toTable.getByLabel('Number of columns')).toHaveValue('2')
      await toTable.screenshot({ path: screenshotPath('docs-text-to-table-dialog') })
      await toTable.getByRole('button', { name: 'OK' }).click()
      await expect(page.locator('.doc-page .doc-table')).toHaveCount(1)
      await expect.poll(() => column(1)).toEqual(['Qty', '250', '12', '7'])

      expect(await page.evaluate(() => (window as unknown as AidocsWindow).__aidocs!.save!())).toBe(
        true,
      )
      await expect
        .poll(() => documentXml(docPath), { timeout: 15_000 })
        .toMatch(/<w:tbl>[\s\S]*Paper[\s\S]*Pens[\s\S]*Clips[\s\S]*<\/w:tbl>/)
    } finally {
      await closeAndSaveVideo(launched, 'docs-table-data')
    }
  })
})
