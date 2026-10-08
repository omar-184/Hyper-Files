import { existsSync } from 'node:fs'
import { createRequire } from 'node:module'
import { join } from 'node:path'
import { BrowserWindow, app, dialog, ipcMain } from 'electron'
import {
  IMAGE_EXTENSIONS,
  PDF_TOOLS_CHANNELS,
  TOOL_IDS,
  type ToolProgress,
  type ToolRequest,
  type ToolResult,
} from '../../shared/pdf-tools-api'
import { PdfToolsRunner } from './runner'

/** mupdf.js: shipped under Resources/mupdf in a packaged app (electron-builder.cjs). */
export function mupdfModulePath(): string {
  const packaged = join(process.resourcesPath ?? '', 'mupdf', 'mupdf.js')
  if (app.isPackaged && existsSync(packaged)) return packaged
  return createRequire(import.meta.url).resolve('mupdf')
}

function badRequest(message: string): ToolResult {
  return { ok: false, error: { code: 'bad-input', message } }
}

/** The renderer is ours, but a request still crosses a process boundary: check its shape. */
function validRequest(raw: unknown): raw is ToolRequest {
  if (!raw || typeof raw !== 'object') return false
  const r = raw as Partial<ToolRequest>
  if (!TOOL_IDS.includes(r.tool as ToolRequest['tool'])) return false
  if (!Array.isArray(r.files) || r.files.length > 500) return false
  if (!r.files.every((f) => f && typeof f.path === 'string' && f.path.length > 0)) return false
  if (r.outputDir !== undefined && typeof r.outputDir !== 'string') return false
  return typeof r.options === 'object' && r.options !== null
}

export function registerPdfToolsIpc(workerPath: string): PdfToolsRunner {
  const runner = new PdfToolsRunner(workerPath, () => ({ mupdfPath: mupdfModulePath() }))

  ipcMain.handle(PDF_TOOLS_CHANNELS.pickFiles, async (event, kind: unknown, multiple: unknown) => {
    const win = BrowserWindow.fromWebContents(event.sender) ?? undefined
    const filters =
      kind === 'image'
        ? [{ name: 'Images', extensions: [...IMAGE_EXTENSIONS] }]
        : [{ name: 'PDF', extensions: ['pdf'] }]
    const options: Electron.OpenDialogOptions = {
      properties: multiple === true ? ['openFile', 'multiSelections'] : ['openFile'],
      filters,
    }
    const res = win
      ? await dialog.showOpenDialog(win, options)
      : await dialog.showOpenDialog(options)
    return res.canceled ? [] : res.filePaths
  })

  ipcMain.handle(PDF_TOOLS_CHANNELS.pickFolder, async (event) => {
    const win = BrowserWindow.fromWebContents(event.sender) ?? undefined
    const options: Electron.OpenDialogOptions = { properties: ['openDirectory', 'createDirectory'] }
    const res = win
      ? await dialog.showOpenDialog(win, options)
      : await dialog.showOpenDialog(options)
    return res.canceled || res.filePaths.length === 0 ? null : res.filePaths[0]
  })

  ipcMain.handle(PDF_TOOLS_CHANNELS.info, (_event, path: unknown, password: unknown) => {
    if (typeof path !== 'string' || path === '') {
      return { ok: false, error: { code: 'bad-input', message: 'no file' } }
    }
    return runner.info(path, typeof password === 'string' && password ? password : undefined)
  })

  ipcMain.handle(PDF_TOOLS_CHANNELS.run, (event, raw: unknown) => {
    if (!validRequest(raw)) return badRequest('the request could not be read')
    const sender = event.sender
    return runner.run(raw, (done, total) => {
      if (!sender.isDestroyed()) {
        sender.send(PDF_TOOLS_CHANNELS.progress, { done, total } satisfies ToolProgress)
      }
    })
  })

  app.on('will-quit', () => runner.stop())
  return runner
}
