import { test, expect } from '@playwright/test'
import { execSync } from 'node:child_process'
import { copyFile, mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import type { ElectronApplication, Page } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, screenshotPath } from './helpers'

const FIXTURE = resolve(__dirname, '../apps/sheets/fixtures/generated/compatibility-basic.xlsx')
/** the red swatch in Univer's tab color palette */
const RED = 'rgb(240, 82, 82)'

function sheetXml(workbookPath: string): string {
  return execSync(`unzip -p "${workbookPath}" xl/worksheets/sheet1.xml`).toString()
}

async function openSheets(app: ElectronApplication): Promise<Page> {
  const sheets = await waitForPageWithUrl(app, '://sheets/')
  await sheets.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, {
    timeout: 30_000,
  })
  await sheets.waitForTimeout(1_500)
  return sheets
}

test('sheets: tab color survives save and reopen', async () => {
  const scratch = await mkdtemp(join(tmpdir(), 'genoffice-sheets-tab-color-'))
  const workbook = join(scratch, 'tab-color.xlsx')
  await copyFile(FIXTURE, workbook)
  expect(sheetXml(workbook)).not.toContain('<tabColor')

  // ── session 1: sheet tab ▸ Change color ▸ red, then save ──
  const first = await launchShell({
    onboardingSeen: true,
    videoDir: 'sheets-tab-color',
    openFile: workbook,
  })
  try {
    const sheets = await openSheets(first.app)
    await sheets.getByText('Sheet1', { exact: true }).first().click({ button: 'right' })
    await sheets.getByText('Change color').hover()
    const swatch = sheets.locator(`button[style*="background-color: ${RED}"]`).first()
    await swatch.waitFor()
    // hovering straight across can close the submenu; settle on the swatch first
    await swatch.hover()
    await swatch.click()
    await sheets.screenshot({ path: screenshotPath('sheets-tab-colored') })

    await first.app.evaluate(({ webContents }) => {
      const wc = webContents.getAllWebContents().find((w) => w.getURL().includes('://sheets/'))
      wc?.send('menu:action', 'save')
    })
    await expect(() => {
      expect(sheetXml(workbook)).toContain('<sheetPr><tabColor rgb="FFF05252"/></sheetPr>')
    }).toPass({ timeout: 15_000 })
    // the rest of the sheet is untouched
    expect(sheetXml(workbook)).toContain('<v>10</v>')
  } finally {
    await closeAndSaveVideo(first, 'sheets-tab-color')
  }

  // ── session 2: the reopened tab is still red ──
  const second = await launchShell({
    onboardingSeen: true,
    videoDir: 'sheets-tab-color-reopen',
    openFile: workbook,
  })
  try {
    const sheets = await openSheets(second.app)
    const colored = await sheets.evaluate(
      (red) =>
        [...document.querySelectorAll<HTMLElement>('[style]')].some((el) =>
          el.style.cssText.includes(red),
        ),
      RED,
    )
    await sheets.screenshot({ path: screenshotPath('sheets-tab-color-reopened') })
    expect(colored).toBe(true)
  } finally {
    await closeAndSaveVideo(second, 'sheets-tab-color-reopen')
  }
})
