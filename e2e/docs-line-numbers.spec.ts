import { test, expect } from '@playwright/test'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, screenshotPath, waitForPageWithUrl } from './helpers'

/**
 * Layout ▸ Line Numbers: Continuous shows numbers in the margin and saves
 * w:lnNumType; the options dialog changes Count by; None removes it again.
 */

interface AidocsWindow {
  __aidocs?: { editor?: unknown; save?: () => Promise<unknown> }
}

const SECT_PR =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="720"/></w:sectPr>'

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
  const paras = ['First line', 'Second line', 'Third line', 'Fourth line']
    .map((text) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`)
    .join('')
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${paras}${SECT_PR}</w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

async function documentXml(path: string): Promise<string> {
  const zip = await JSZip.loadAsync(readFileSync(path))
  return zip.file('word/document.xml')!.async('string')
}

test.describe('docs line numbers', () => {
  let dir: string
  let docPath: string

  test.beforeEach(async () => {
    dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-line-numbers-')))
    docPath = join(dir, 'lines.docx')
    writeFileSync(docPath, await minimalDocx())
  })

  test.afterEach(() => {
    rmSync(dir, { recursive: true, force: true })
  })

  test('turn line numbers on, change count by, save, then remove them', async () => {
    test.setTimeout(180_000)
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-line-numbers',
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
      const save = async () =>
        expect(
          await page.evaluate(() => (window as unknown as AidocsWindow).__aidocs!.save!()),
        ).toBe(true)
      const numbers = page.locator('.page-line-numbers .page-line-number')

      await page.locator('.doc-page p').first().click()
      await page.locator('.ribbon-tab', { hasText: 'Layout' }).first().click()
      await page.getByRole('button', { name: 'Line Numbers' }).click()
      await page.locator('.layout-menu button', { hasText: 'Continuous' }).click()
      await expect(numbers).toHaveCount(4)
      await expect(numbers.first()).toHaveText('1')
      await save()
      await expect
        .poll(() => documentXml(docPath), { timeout: 15_000 })
        .toContain('<w:lnNumType w:countBy="1" w:restart="continuous"/>')

      // options dialog: count by 2 labels every second line
      await page.getByRole('button', { name: 'Line Numbers' }).click()
      await page.locator('.layout-menu button', { hasText: 'Line Numbering Options' }).click()
      const dialog = page.locator('.modal.line-numbers-dialog')
      await expect(dialog).toBeVisible()
      await dialog.getByLabel('Count by').fill('2')
      await dialog.screenshot({ path: screenshotPath('docs-line-numbers-dialog') })
      await dialog.getByRole('button', { name: 'OK' }).click()
      await expect(numbers).toHaveText(['2', '4'])
      await page.screenshot({ path: screenshotPath('docs-line-numbers-page') })
      await save()
      await expect
        .poll(() => documentXml(docPath), { timeout: 15_000 })
        .toContain('<w:lnNumType w:countBy="2" w:restart="continuous"/>')

      await page.getByRole('button', { name: 'Line Numbers' }).click()
      await page.locator('.layout-menu button', { hasText: 'None' }).click()
      await expect(numbers).toHaveCount(0)
      await save()
      await expect.poll(() => documentXml(docPath), { timeout: 15_000 }).not.toContain('lnNumType')
    } finally {
      await closeAndSaveVideo(launched, 'docs-line-numbers')
    }
  })
})
