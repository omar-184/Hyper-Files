import { Worker } from 'node:worker_threads'
import type { FileInfo, ToolRequest, ToolResult } from '../../shared/pdf-tools-api'
import type { WorkerInit, WorkerMessage, WorkerRequest } from './worker'

/** Idle time before the worker (and its wasm heap) is released. */
const IDLE_MS = 30_000

/** Omit that keeps a union a union */
type WithoutId<T> = T extends unknown ? Omit<T, 'id'> : never

type Pending = {
  /** null when the worker died before answering */
  resolve: (msg: WorkerMessage | null) => void
  onProgress?: (done: number, total: number) => void
}

/**
 * One worker thread, started on first use and stopped when idle. Requests run
 * one at a time in the worker (MuPDF is single-threaded); a crash fails the
 * requests in flight and the next request starts a fresh worker.
 */
export class PdfToolsRunner {
  private worker: Worker | null = null
  private nextId = 1
  private readonly pending = new Map<number, Pending>()
  private idleTimer: NodeJS.Timeout | null = null

  constructor(
    private readonly workerPath: string,
    private readonly init: () => WorkerInit,
  ) {}

  run(request: ToolRequest, onProgress?: (done: number, total: number) => void) {
    return this.send({ type: 'run', request }, onProgress).then((msg) =>
      msg?.type === 'run' ? msg.result : crashed<ToolResult>(),
    )
  }

  info(path: string, password?: string): Promise<FileInfo> {
    return this.send({ type: 'info', path, ...(password ? { password } : {}) }).then((msg) =>
      msg?.type === 'info' ? msg.info : crashed<FileInfo>(),
    )
  }

  stop(): void {
    if (this.idleTimer) clearTimeout(this.idleTimer)
    this.idleTimer = null
    const w = this.worker
    this.worker = null
    void w?.terminate()
  }

  private send(
    req: WithoutId<WorkerRequest>,
    onProgress?: (done: number, total: number) => void,
  ): Promise<WorkerMessage | null> {
    if (this.idleTimer) clearTimeout(this.idleTimer)
    this.idleTimer = null
    const id = this.nextId++
    const worker = this.ensureWorker()
    return new Promise((resolve) => {
      this.pending.set(id, { resolve, onProgress })
      worker.postMessage({ ...req, id } as WorkerRequest)
    })
  }

  private ensureWorker(): Worker {
    if (this.worker) return this.worker
    const w = new Worker(this.workerPath, { workerData: this.init() })
    w.on('message', (msg: WorkerMessage) => {
      const p = this.pending.get(msg.id)
      if (!p) return
      if (msg.type === 'progress') {
        p.onProgress?.(msg.done, msg.total)
        return
      }
      this.pending.delete(msg.id)
      p.resolve(msg)
      if (this.pending.size === 0) this.scheduleIdleStop()
    })
    const drop = (err?: unknown) => {
      if (this.worker !== w) return
      this.worker = null
      if (err) console.error('pdf tools worker failed', err)
      for (const p of this.pending.values()) p.resolve(null)
      this.pending.clear()
    }
    w.on('error', drop)
    w.on('exit', () => drop())
    this.worker = w
    return w
  }

  private scheduleIdleStop() {
    if (this.idleTimer) clearTimeout(this.idleTimer)
    this.idleTimer = setTimeout(() => {
      this.idleTimer = null
      if (this.pending.size === 0) this.stop()
    }, IDLE_MS)
    this.idleTimer.unref?.()
  }
}

/** What a request resolves to when the worker died under it (out of memory, wasm abort). */
function crashed<T extends ToolResult | FileInfo>(): T {
  return {
    ok: false,
    error: {
      code: 'failed',
      message: 'the PDF engine stopped unexpectedly; the file may be too large for this computer',
    },
  } as T
}
