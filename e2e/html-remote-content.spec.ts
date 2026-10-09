/**
 * Opening an HTML file must not reach the network: the preview blocks web content (a
 * tracking pixel would otherwise tell the sender the file was opened) until the user
 * clicks "Load web content" for that tab. Local pictures beside the file still load.
 */
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { test, expect, type ElectronApplication } from '@playwright/test'
import { PNG } from 'pngjs'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

// .invalid never resolves, and the hook below cancels every request anyway
const REMOTE = [
  'https://fonts.example.invalid/css2?family=Roboto',
  'https://cdn.example.invalid/tracker.js',
  'https://images.example.invalid/logo.png',
  'https://tracker.example.invalid/pixel.gif?doc=letter123',
]

const PAGE = `<!doctype html>
<html>
<head>
<link rel="stylesheet" href="${REMOTE[0]}">
<script src="${REMOTE[1]}"></script>
</head>
<body>
<h1>Letter</h1>
<img id="remote" src="${REMOTE[2]}">
<img id="pixel" src="${REMOTE[3]}" width="1" height="1">
<img id="local" src="local.png">
</body>
</html>
`

/** Record (and cancel) every http(s) request the app makes from now on */
async function recordRemoteRequests(app: ElectronApplication): Promise<void> {
  await app.evaluate(({ session }) => {
    const g = globalThis as unknown as { __remote: string[] }
    g.__remote = []
    session.defaultSession.webRequest.onBeforeRequest((details, callback) => {
      if (
        /^https?:/i.test(details.url) &&
        !/^https?:\/\/(localhost|127\.0\.0\.1)[:/]/.test(details.url)
      ) {
        g.__remote.push(details.url)
        callback({ cancel: true })
      } else callback({})
    })
  })
}

const remoteRequests = (app: ElectronApplication) =>
  app.evaluate(() => (globalThis as unknown as { __remote: string[] }).__remote)

test('an HTML file loads no web content until the user asks for it', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'hyperfiles-html-remote-'))
  const htmlPath = join(dir, 'letter.html')
  await writeFile(htmlPath, PAGE)
  const png = new PNG({ width: 4, height: 4 })
  png.data.fill(200)
  await writeFile(join(dir, 'local.png'), PNG.sync.write(png))

  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'html-remote-content',
    openFile: htmlPath,
  })
  const { app } = launched
  try {
    let editor = await waitForPageWithUrl(app, '://html/')
    await expect(editor.locator('.preview-frame')).toBeVisible()
    // The hook can only be installed once the app runs: reload the editor so its first
    // preview load happens under it
    await recordRemoteRequests(app)
    await editor.reload()
    editor = await waitForPageWithUrl(app, '://html/')
    const frame = editor.frameLocator('.preview-frame')
    await expect(frame.locator('h1')).toHaveText('Letter')

    const bar = editor.locator('.remote-bar')
    await expect(bar).toContainText('blocked in this preview')
    // the local picture beside the file still loads
    await expect
      .poll(() => frame.locator('#local').evaluate((img: HTMLImageElement) => img.naturalWidth))
      .toBe(4)
    await editor.waitForTimeout(1000)
    expect(await remoteRequests(app)).toEqual([])

    // One click loads it, for this tab
    await bar.getByRole('button', { name: 'Load web content' }).click()
    await expect(bar).toHaveCount(0)
    await expect.poll(async () => new Set(await remoteRequests(app))).toEqual(new Set(REMOTE))
  } finally {
    await closeAndSaveVideo(launched, 'html-remote-content')
    await rm(dir, { recursive: true, force: true })
  }
})
