import { test, expect } from '@playwright/test'
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { PDFArray, PDFDict, PDFDocument, PDFName } from 'pdf-lib'
import { launchShell, waitForPageWithUrl, closeAndSaveVideo } from './helpers'

async function freeTexts(path: string): Promise<string[]> {
  const pdf = await PDFDocument.load(await readFile(path))
  const annots = pdf.getPage(0).node.lookupMaybe(PDFName.of('Annots'), PDFArray)
  if (!annots) return []
  return Array.from({ length: annots.size() }, (_, i) => annots.lookup(i, PDFDict))
    .filter((a) => a.lookup(PDFName.of('Subtype'), PDFName).decodeText() === 'FreeText')
    .map((a) => (a.lookup(PDFName.of('Contents')) as { decodeText(): string }).decodeText())
}

test('a text-box comment is typed on the page, edited, saved and listed', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'hyperfiles-text-box-'))
  const source = join(dir, 'sample.pdf')
  const pdf = await PDFDocument.create()
  pdf.addPage([400, 300])
  await writeFile(source, await pdf.save())
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'pdf-text-box',
    openFile: source,
  })
  try {
    const editor = await waitForPageWithUrl(launched.app, '://pdf/')
    await expect(editor.locator('.pdf-page').first()).toBeVisible()
    await editor.getByRole('button', { name: 'Annotate', exact: true }).click()
    await editor.getByRole('button', { name: 'Text box', exact: true }).click()

    const layer = editor.locator('.pdf-page .pdf-draw-layer').first()
    const box = await layer.boundingBox()
    if (!box) throw new Error('Missing draw layer')
    await editor.mouse.click(box.x + box.width * 0.2, box.y + box.height * 0.2)
    const input = editor.getByRole('textbox', { name: 'Text box' })
    await expect(input).toBeFocused()
    await input.fill('Check the total')
    await input.press('ControlOrMeta+Enter')
    await expect(input).toHaveCount(0)
    await expect(layer.locator('text')).toHaveText('Check the total')

    // Double-click reopens the editor with the text
    await editor.getByRole('button', { name: 'Text box', exact: true }).click()
    await layer.locator('text').dblclick()
    await expect(input).toHaveValue('Check the total')
    await input.fill('Check the grand total')
    await input.press('ControlOrMeta+Enter')
    await expect(layer.locator('text')).toHaveText('Check the grand total')

    await editor.keyboard.press('ControlOrMeta+s')
    await expect.poll(() => freeTexts(source)).toEqual(['Check the grand total'])
    await expect(editor.getByText('Unsaved', { exact: true })).toHaveCount(0)

    // After the reload it is a saved annotation: listed with its text
    await editor.getByRole('button', { name: 'View', exact: true }).click()
    await editor.getByRole('button', { name: 'Comments', exact: true }).click()
    const row = editor.locator('.pdf-comment-row')
    await expect(row).toHaveCount(1)
    await expect(row).toContainText('Text box')
    await expect(row).toContainText('Check the grand total')
  } finally {
    await closeAndSaveVideo(launched, 'pdf-text-box')
    await rm(dir, { recursive: true, force: true })
  }
})
