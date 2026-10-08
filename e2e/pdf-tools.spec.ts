import { test, expect } from '@playwright/test'
import { mkdtemp, readFile, readdir, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, screenshotPath } from './helpers'

/**
 * A one-page PDF reading `label`. The cross-reference table is left out on
 * purpose: MuPDF rebuilds it, which also exercises the repair path.
 */
function tinyPdf(label: string): string {
  const content = `BT /F1 24 Tf 72 720 Td (${label}) Tj ET`
  return [
    '%PDF-1.4',
    '1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj',
    '2 0 obj << /Type /Pages /Kids [3 0 R] /Count 1 >> endobj',
    '3 0 obj << /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R',
    '  /Resources << /Font << /F1 5 0 R >> >> >> endobj',
    `4 0 obj << /Length ${content.length} >> stream`,
    content,
    'endstream endobj',
    '5 0 obj << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> endobj',
    'trailer << /Root 1 0 R >>',
    '%%EOF',
    '',
  ].join('\n')
}

/** Page objects in a file MuPDF wrote without object streams. */
function countPages(pdf: string): number {
  return (pdf.match(/\/Type\s*\/Page(?![s\w])/g) ?? []).length
}

test.describe('PDF tools', () => {
  test('lists the tools and merges two files into one', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'pdf-tools-e2e-'))
    const a = join(dir, 'first.pdf')
    const b = join(dir, 'second.pdf')
    await writeFile(a, tinyPdf('First'))
    await writeFile(b, tinyPdf('Second'))

    const launched = await launchShell({ onboardingSeen: true, videoDir: 'pdf-tools-merge' })
    const { app, page } = launched
    try {
      await page.locator('.nav-item', { hasText: 'PDF Tools' }).click()
      await expect(page.locator('.pt-card')).toHaveCount(17)
      await page.screenshot({ path: screenshotPath('pdf-tools-grid') })

      // the native file picker cannot be driven; answer it from the main process
      await app.evaluate(
        ({ dialog }, paths) => {
          dialog.showOpenDialog = (async () => ({
            canceled: false,
            filePaths: paths,
          })) as typeof dialog.showOpenDialog
        },
        [a, b],
      )

      await page.locator('.pt-card', { hasText: 'Merge PDF' }).click()
      await page.getByRole('button', { name: 'Choose PDF files' }).click()
      await expect(page.locator('.pt-file')).toHaveCount(2)
      await expect(page.locator('.pt-file').first()).toContainText('1 page')
      await expect(page.locator('.pt-file').nth(1)).toContainText('1 page')
      await page.screenshot({ path: screenshotPath('pdf-tools-merge-files') })

      await page.locator('.pt-start').click()
      await expect(page.locator('.pt-results-title')).toHaveText('Done')
      await page.screenshot({ path: screenshotPath('pdf-tools-merge-done') })

      expect((await readdir(dir)).sort()).toEqual(['first-merged.pdf', 'first.pdf', 'second.pdf'])
      const merged = await readFile(join(dir, 'first-merged.pdf'), 'latin1')
      expect(countPages(merged)).toBe(2)
    } finally {
      await closeAndSaveVideo(launched, 'pdf-tools-merge')
    }
  })
})
