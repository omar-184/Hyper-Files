/**
 * Worker-thread entry for the PDF tools: MuPDF runs here so a long compress
 * or a 500-page render never freezes the window. The runner terminates the
 * worker when it goes idle, which returns the wasm heap to the system.
 */
import { pathToFileURL } from 'node:url'
import { parentPort, workerData } from 'node:worker_threads'
import type { Mupdf, OcrRecognizer } from '@genoffice/pdf-tools'
import {
  createVisionOcrEngine,
  createWindowsOcrEngine,
} from '../../../../../packages/pdf2docx/src/ocr-vision'
import type { FileInfo, ToolRequest, ToolResult } from '../../shared/pdf-tools-api'
import { fileInfo, runRequest } from './jobs'

export interface WorkerInit {
  /** absolute path of mupdf.js (dev: node_modules; packaged: Resources/mupdf) */
  mupdfPath: string
  /** system OCR helper binary; null when this computer has none */
  ocrHelperPath: string | null
}

export type WorkerRequest =
  | { id: number; type: 'run'; request: ToolRequest }
  | { id: number; type: 'info'; path: string; password?: string }

export type WorkerMessage =
  | { id: number; type: 'run'; result: ToolResult }
  | { id: number; type: 'info'; info: FileInfo }
  | { id: number; type: 'progress'; done: number; total: number }

const { mupdfPath, ocrHelperPath } = workerData as WorkerInit

/** The OCR engine is only started by the OCR tool, never on open. */
let ocr: OcrRecognizer | null | undefined
function ocrEngine(): OcrRecognizer | null {
  if (ocr === undefined) {
    const create = process.platform === 'darwin' ? createVisionOcrEngine : createWindowsOcrEngine
    ocr = ocrHelperPath ? create(ocrHelperPath) : null
  }
  return ocr
}

// mupdf.js is an ES module with top-level await; importing it by URL at run
// time keeps it out of the CommonJS bundle this file is compiled into
const load = new Function('url', 'return import(url)') as (url: string) => Promise<Mupdf>
const mupdf: Promise<Mupdf> = load(pathToFileURL(mupdfPath).href).then((m) => {
  // MuPDF reports recoverable damage as warnings; results carry what matters
  m.setLog(null)
  return m
})

parentPort?.on('message', async (req: WorkerRequest) => {
  const m = await mupdf
  if (req.type === 'info') {
    const info = fileInfo(m, req.path, req.password)
    parentPort?.postMessage({ id: req.id, type: 'info', info } satisfies WorkerMessage)
    return
  }
  const result = runRequest(
    m,
    req.request,
    (done, total) =>
      parentPort?.postMessage({
        id: req.id,
        type: 'progress',
        done,
        total,
      } satisfies WorkerMessage),
    req.request.tool === 'ocr' ? { ocr: ocrEngine() } : {},
  )
  parentPort?.postMessage({ id: req.id, type: 'run', result } satisfies WorkerMessage)
})
