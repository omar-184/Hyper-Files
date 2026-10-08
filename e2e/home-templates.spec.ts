import { test, expect } from '@playwright/test'
import { execSync } from 'node:child_process'
import { existsSync } from 'node:fs'
import { mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, screenshotPath } from './helpers'

test.describe('home: new from template', () => {
  test('each template lands as a real file and opens in its editor', async () => {
    const root = await mkdtemp(join(tmpdir(), 'genoffice-templates-'))
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'home-templates',
      settings: { defaultSaveDir: root },
    })
    const { page, app } = launched
    try {
      await expect(page.locator('.template-card')).toHaveCount(5)
      await page.screenshot({ path: screenshotPath('home-templates') })

      // Letter → a Word file in the save folder, opened in Docs with its text
      await page.locator('.template-card', { hasText: 'Letter' }).click()
      const docs = await waitForPageWithUrl(app, '://docs/')
      await expect(docs.locator('.ProseMirror')).toContainText('Dear [Recipient Name],', {
        timeout: 30_000,
      })
      expect(existsSync(join(root, 'Letter.docx'))).toBe(true)

      // Budget → a workbook whose formulas already carry their totals
      await page.bringToFront()
      await page.locator('.tab-bar .tab-item.tab-home').click()
      await page.locator('.template-card', { hasText: 'Monthly budget' }).click()
      await waitForPageWithUrl(app, '://sheets/')
      const budget = join(root, 'Monthly Budget.xlsx')
      await expect(() => expect(existsSync(budget)).toBe(true)).toPass({ timeout: 15_000 })
      const sheetXml = execSync(`unzip -p "${budget}" xl/worksheets/sheet1.xml`).toString()
      expect(sheetXml).toContain('<f>B7-B17</f><v>750</v>')

      // picking the same template again never overwrites the first copy
      await page.locator('.tab-bar .tab-item.tab-home').click()
      await page.locator('.template-card', { hasText: 'Letter' }).click()
      await expect(() => expect(existsSync(join(root, 'Letter-2.docx'))).toBe(true)).toPass({
        timeout: 15_000,
      })
    } finally {
      await closeAndSaveVideo(launched, 'home-templates')
    }
  })
})
