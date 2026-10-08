import { test, expect } from '@playwright/test'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, screenshotPath, waitForPageWithUrl } from './helpers'

/**
 * Layout ▸ Hyphenation: Automatic breaks line-end words with a hyphen (soft
 * hyphen widgets, never document text) and saves w:autoHyphenation; None
 * removes both again.
 */

interface AidocsWindow {
  __aidocs?: { editor?: unknown; save?: () => Promise<unknown> }
}

const TEXT = Array.from(
  { length: 6 },
  () =>
    'The international organization documented extraordinary administrative responsibilities, comprehensive environmental considerations and representative recommendations.',
).join(' ')

const SECT_PR =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="720"/></w:sectPr>'

/** one paragraph of TEXT; `justifiedWord2013` makes it a justified
 *  Word 2013+ document that already has automatic hyphenation on */
async function minimalDocx(justifiedWord2013 = false): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
  )
  if (justifiedWord2013) {
    const types = await zip.file('[Content_Types].xml')!.async('string')
    zip.file(
      '[Content_Types].xml',
      types.replace(
        '</Types>',
        '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>',
      ),
    )
    zip.file(
      'word/_rels/document.xml.rels',
      '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>',
    )
    zip.file(
      'word/settings.xml',
      '<?xml version="1.0" encoding="UTF-8"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:autoHyphenation/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>',
    )
  }
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  )
  const pPr = justifiedWord2013 ? '<w:pPr><w:jc w:val="both"/></w:pPr>' : ''
  const paras = `<w:p>${pPr}<w:r><w:t>${TEXT}</w:t></w:r></w:p>`
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${paras}${SECT_PR}</w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

async function part(path: string, name: string): Promise<string> {
  const zip = await JSZip.loadAsync(readFileSync(path))
  return (await zip.file(name)?.async('string')) ?? ''
}

/** soft hyphens Chromium broke a line at render the hyphen glyph */
function visibleHyphens(): number {
  return Array.from(document.querySelectorAll('.doc-shy')).filter(
    (el) => el.getBoundingClientRect().width > 0,
  ).length
}

test.describe('docs hyphenation', () => {
  let dir: string
  let docPath: string

  test.beforeEach(async () => {
    dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-hyphenation-')))
    docPath = join(dir, 'hyphenation.docx')
    writeFileSync(docPath, await minimalDocx())
  })

  test.afterEach(() => {
    rmSync(dir, { recursive: true, force: true })
  })

  test('automatic hyphenation breaks line-end words and saves the setting', async () => {
    test.setTimeout(180_000)
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-hyphenation',
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
      const paragraph = page.locator('.doc-page p').first()
      await paragraph.click()
      const ragged = await paragraph.evaluate((el) => el.getBoundingClientRect().height)
      expect(await page.locator('.doc-shy').count()).toBe(0)

      await page.locator('.ribbon-tab', { hasText: 'Layout' }).first().click()
      await page.getByRole('button', { name: 'Hyphenation' }).click()
      await page.locator('.layout-menu button', { hasText: 'Automatic' }).click()
      await expect.poll(() => page.evaluate(visibleHyphens), { timeout: 10_000 }).toBeGreaterThan(0)
      // hyphenated lines hold more text: the paragraph never grows
      expect(
        await paragraph.evaluate((el) => el.getBoundingClientRect().height),
      ).toBeLessThanOrEqual(ragged)
      await page.screenshot({ path: screenshotPath('docs-hyphenation-on') })

      // the break points are display only: the text and the saved words stay whole
      expect(await paragraph.evaluate((el) => el.textContent?.includes('\u00AD'))).toBe(true)
      expect(
        await page.evaluate(() =>
          (
            (window as unknown as AidocsWindow).__aidocs!.editor as {
              state: { doc: { textContent: string } }
            }
          ).state.doc.textContent.includes('\u00AD'),
        ),
      ).toBe(false)
      await save()
      await expect
        .poll(() => part(docPath, 'word/settings.xml'), { timeout: 15_000 })
        .toContain('<w:autoHyphenation/>')
      const saved = await part(docPath, 'word/document.xml')
      expect(saved).not.toContain('softHyphen')
      expect(saved).toContain('international organization')

      // typing inside a hyphenated paragraph still edits the plain text
      await paragraph.click({ position: { x: 4, y: 4 } })
      await page.keyboard.press('Home')
      await page.keyboard.type('Overall, ')
      await expect(paragraph).toContainText('Overall, The international')
      await expect.poll(() => page.evaluate(visibleHyphens), { timeout: 10_000 }).toBeGreaterThan(0)

      await page.getByRole('button', { name: 'Hyphenation' }).click()
      await page.locator('.layout-menu button', { hasText: 'None' }).click()
      await expect(page.locator('.doc-shy')).toHaveCount(0)
      await save()
      await expect
        .poll(() => part(docPath, 'word/settings.xml'), { timeout: 15_000 })
        .not.toContain('autoHyphenation')
    } finally {
      await closeAndSaveVideo(launched, 'docs-hyphenation')
    }
  })
  test('a justified Word 2013 document with hyphenation on settles with its shrink pass', async () => {
    test.setTimeout(180_000)
    writeFileSync(docPath, await minimalDocx(true))
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-hyphenation-justified',
      openFile: docPath,
    })
    const { app } = launched
    try {
      const page = await waitForPageWithUrl(app, '://docs/')
      const warnings: string[] = []
      page.on('console', (msg) => {
        if (msg.text().includes('did not converge')) warnings.push(msg.text())
      })
      await page.waitForFunction(
        () => Boolean((window as unknown as AidocsWindow).__aidocs?.editor),
        undefined,
        { timeout: 30_000 },
      )
      await expect.poll(() => page.evaluate(visibleHyphens), { timeout: 10_000 }).toBeGreaterThan(0)
      await page.locator('.ribbon-tab', { hasText: 'Layout' }).first().click()
      await expect(page.getByRole('button', { name: 'Hyphenation' })).toHaveClass(/active/)
      await page.screenshot({ path: screenshotPath('docs-hyphenation-justified') })
      await page.waitForTimeout(1_000)
      expect(warnings).toEqual([])
    } finally {
      await closeAndSaveVideo(launched, 'docs-hyphenation-justified')
    }
  })
})
