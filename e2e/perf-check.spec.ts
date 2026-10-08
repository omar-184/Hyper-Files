import { test, expect } from '@playwright/test'
import { appendFileSync, mkdirSync, writeFileSync } from 'node:fs'
import { copyFile, mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import type { ElectronApplication, Page } from 'playwright-core'
import { PDFDocument, StandardFonts } from 'pdf-lib'
import {
  ARTIFACTS_DIR,
  closeAndSaveVideo,
  launchShell,
  screenshotPath,
  waitForPageWithUrl,
} from './helpers'

/**
 * Settings → Performance end to end: the self-test must finish every step on
 * a real build (each sample opens in a fresh headless app process), and the
 * run prints real numbers. A second test opens one document of each kind and
 * records the app's memory as tabs pile up. Both write their numbers to
 * e2e/artifacts and, on CI, to the job summary.
 */

const DOCX = resolve(__dirname, 'assets/justify-pagegap-fr.docx')
const XLSX = resolve(__dirname, '../apps/sheets/fixtures/generated/compatibility-basic.xlsx')
const PPTX = resolve(__dirname, '../packages/pptx-engine/tests/fixtures/01_standard_business.pptx')

function publish(name: string, text: string): void {
  mkdirSync(ARTIFACTS_DIR, { recursive: true })
  writeFileSync(join(ARTIFACTS_DIR, name), text + '\n')
  console.log(`\n${text}\n`)
  const summary = process.env.GITHUB_STEP_SUMMARY
  if (summary) appendFileSync(summary, '\n```\n' + text + '\n```\n')
}

async function openSettingsPerformance(page: Page): Promise<void> {
  await page.getByRole('button', { name: 'Settings' }).first().click()
  await page
    .getByRole('dialog', { name: 'Settings' })
    .getByRole('button', { name: 'Performance' })
    .click()
}

test.describe('settings performance check', () => {
  test('runs every step offline and reports numbers', async () => {
    test.setTimeout(600_000)
    const launched = await launchShell({ videoDir: 'perf-check', onboardingSeen: true })
    const { app, page } = launched
    try {
      await openSettingsPerformance(page)
      await page.getByTestId('perf-run').click()
      const result = page.getByTestId('perf-result')
      await expect(result).toBeVisible({ timeout: 540_000 })

      await page.screenshot({ path: screenshotPath('perf-check-light') })
      await page.evaluate(() => document.documentElement.setAttribute('data-theme', 'dark'))
      await page.screenshot({ path: screenshotPath('perf-check-dark') })

      await page.getByTestId('perf-copy').click()
      await expect(page.getByTestId('perf-copy')).toHaveText('Copied')
      const text = await app.evaluate(({ clipboard }) => clipboard.readText())
      publish('perf-report.txt', text)

      // every measurement finished: nothing in the report failed
      expect(text).not.toContain('[failed]')
      for (const row of ['docx', 'xlsx', 'pptx', 'pdf', 'cpu', 'disk']) {
        await expect(result.locator(`[data-perf-row="${row}"] .perf-badge`)).not.toHaveText(
          'Failed',
        )
      }
      // the run left no window behind: the samples rendered in separate processes
      expect(await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows().length)).toBe(
        1,
      )
    } finally {
      await closeAndSaveVideo(launched, 'perf-check')
    }
  })

  test('records memory use as documents open', async () => {
    test.setTimeout(300_000)
    const dir = await mkdtemp(join(tmpdir(), 'genoffice-e2e-perf-mem-'))
    const files = {
      docx: join(dir, 'sample.docx'),
      xlsx: join(dir, 'sample.xlsx'),
      pptx: join(dir, 'sample.pptx'),
      pdf: join(dir, 'sample.pdf'),
    }
    await copyFile(DOCX, files.docx)
    await copyFile(XLSX, files.xlsx)
    await copyFile(PPTX, files.pptx)
    const pdf = await PDFDocument.create()
    const font = await pdf.embedFont(StandardFonts.Helvetica)
    for (let i = 0; i < 20; i++) {
      const p = pdf.addPage([595, 842])
      for (let line = 0; line < 45; line++) {
        p.drawText(
          `Page ${i + 1}, line ${line + 1}: the quick brown fox jumps over the lazy dog.`,
          {
            x: 50,
            y: 790 - line * 16,
            size: 10,
            font,
          },
        )
      }
    }
    await writeFile(files.pdf, await pdf.save())

    const launched = await launchShell({ videoDir: 'perf-memory', onboardingSeen: true })
    const { app, page } = launched
    const settle = () => new Promise((r) => setTimeout(r, 3_000))
    const sample = (app: ElectronApplication) =>
      app.evaluate(({ app: electronApp }) => {
        const metrics = electronApp.getAppMetrics()
        const byType: Record<string, number> = {}
        let total = 0
        for (const m of metrics) {
          total += m.memory.workingSetSize
          byType[m.type] = (byType[m.type] ?? 0) + m.memory.workingSetSize
        }
        return { totalKB: total, processes: metrics.length, byType }
      })
    try {
      const rows: string[] = []
      const record = async (label: string) => {
        await settle()
        const s = await sample(app)
        const types = Object.entries(s.byType)
          .sort((a, b) => b[1] - a[1])
          .map(([type, kb]) => `${type} ${Math.round(kb / 1024)}`)
          .join(', ')
        rows.push(
          `${label.padEnd(16)} ${String(Math.round(s.totalKB / 1024)).padStart(5)} MB  ${String(s.processes).padStart(2)} processes  (${types})`,
        )
        return s.totalKB / 1024
      }
      await record('Home only')
      const open = async (kind: keyof typeof files, urlPart: string) => {
        await page.evaluate(
          (p) =>
            (
              window as unknown as { hyperFiles: { openPath(p: string): Promise<void> } }
            ).hyperFiles.openPath(p),
          files[kind],
        )
        await waitForPageWithUrl(app, urlPart, 60_000)
        return record(`+ ${kind}`)
      }
      await open('docx', '://docs/')
      await open('xlsx', '://sheets/')
      await open('pptx', '://slides/')
      const total = await open('pdf', '://pdf/')
      publish(
        'perf-memory.txt',
        ['Hypercube Office memory (working set, MB) as tabs open', ...rows].join('\n'),
      )
      // sanity bound, not a target: all four editors together must stay well
      // under what a 4 GB machine can give one app
      expect(total).toBeLessThan(2_500)
    } finally {
      await closeAndSaveVideo(launched, 'perf-memory')
    }
  })
})
