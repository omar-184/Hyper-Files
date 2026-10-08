import { test, expect } from '@playwright/test'
import { readFileSync, writeFileSync } from 'node:fs'
import { mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import JSZip from 'jszip'
import type { Page } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, screenshotPath } from './helpers'

const FIXTURE = resolve(__dirname, '../apps/sheets/fixtures/generated/compatibility-basic.xlsx')
const BLOCKED = 'This cell is on a protected sheet'

/** The basic fixture with a protected Sheet1: A1 stays locked (Excel's
 * default), B1 carries an xf with protection/@locked="0". */
async function protectedWorkbook(path: string): Promise<void> {
  const zip = await JSZip.loadAsync(readFileSync(FIXTURE))
  const styles = await zip.file('xl/styles.xml')!.async('string')
  zip.file(
    'xl/styles.xml',
    styles.replace(
      '<cellXfs count="1"><xf/></cellXfs>',
      '<cellXfs count="2"><xf/><xf applyProtection="1"><protection locked="0"/></xf></cellXfs>',
    ),
  )
  const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  zip.file(
    'xl/worksheets/sheet1.xml',
    sheet
      .replace('<c r="B1"><v>10</v></c>', '<c r="B1" s="1"><v>10</v></c>')
      .replace('</sheetData>', '</sheetData><sheetProtection sheet="1"/>'),
  )
  writeFileSync(path, await zip.generateAsync({ type: 'nodebuffer' }))
}

async function sheetXml(path: string): Promise<string> {
  const zip = await JSZip.loadAsync(readFileSync(path))
  return zip.file('xl/worksheets/sheet1.xml')!.async('string')
}

/** center of cell A1: right of the ~46px row header, below the ~24px column header */
async function cellA1(page: Page): Promise<{ x: number; y: number }> {
  const grid = await page.evaluate(() => {
    for (const canvas of document.querySelectorAll('canvas')) {
      const rect = canvas.getBoundingClientRect()
      if (rect.width > 500 && rect.height > 300) return { x: rect.x, y: rect.y }
    }
    return null
  })
  if (!grid) throw new Error('worksheet canvas not found')
  return { x: grid.x + 46 + 43, y: grid.y + 24 + 12 }
}

test('sheets: a protected sheet blocks locked cells and keeps unlocked ones editable', async () => {
  const scratch = await mkdtemp(join(tmpdir(), 'genoffice-sheets-protect-'))
  const workbook = join(scratch, 'protected.xlsx')
  await protectedWorkbook(workbook)

  const session = await launchShell({
    onboardingSeen: true,
    videoDir: 'sheets-protect-sheet',
    openFile: workbook,
  })
  try {
    const sheets = await waitForPageWithUrl(session.app, '://sheets/')
    await sheets.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, {
      timeout: 30_000,
    })
    // protection is known once the sheet finishes indexing
    await sheets.waitForTimeout(2_500)
    const nameBox = sheets.locator('[data-u-comp="defined-name"] input')

    // Locked A1: typing never opens the editor, and Delete cannot clear it.
    const a1 = await cellA1(sheets)
    await sheets.mouse.click(a1.x, a1.y)
    await expect(nameBox).toHaveValue('A1')
    await sheets.keyboard.type('Blocked', { delay: 30 })
    await sheets.keyboard.press('Enter')
    await expect(sheets.getByText(BLOCKED).first()).toBeVisible()
    await sheets.mouse.click(a1.x, a1.y)
    await sheets.keyboard.press('Delete')
    await sheets.screenshot({ path: screenshotPath('sheets-protect-blocked') })

    // Unlocked B1 takes the edit.
    await sheets.keyboard.press('ArrowRight')
    await expect(nameBox).toHaveValue('B1')
    await sheets.keyboard.type('42', { delay: 30 })
    await sheets.keyboard.press('Enter')

    await session.app.evaluate(({ webContents }) => {
      const wc = webContents.getAllWebContents().find((w) => w.getURL().includes('://sheets/'))
      wc?.send('menu:action', 'save')
    })
    await expect(async () => {
      expect(await sheetXml(workbook)).toContain('<v>42</v>')
    }).toPass({ timeout: 15_000 })
    const saved = await sheetXml(workbook)
    expect(saved).toContain('Old')
    expect(saved).not.toContain('Blocked')
    expect(saved).toContain('<sheetProtection')
  } finally {
    await closeAndSaveVideo(session, 'sheets-protect-sheet')
  }
})
