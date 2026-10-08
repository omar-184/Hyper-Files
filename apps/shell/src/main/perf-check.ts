/**
 * Settings → Performance self-test (shapes and ratings: shared/perf-check.ts).
 *
 * Everything runs on this computer and nothing is sent anywhere:
 *   - system facts from `os` and the free space of the userData drive;
 *   - a fixed JavaScript workload in a worker thread (the UI stays responsive);
 *   - a sequential write of a temporary file, flushed to disk;
 *   - for each sample document (the bundled Home templates), a fresh app
 *     process running the headless export entry: cold start + open + render
 *     to PDF, i.e. roughly what a double-click on a file costs;
 *   - every page of those PDFs rendered to images through the PDF Tools worker;
 *   - the working set of the running app's own processes.
 * Temporary files live in one folder under the OS temp dir, removed at the end.
 */

import { spawn } from 'node:child_process'
import { existsSync, statfsSync } from 'node:fs'
import { mkdtemp, open, rm, writeFile } from 'node:fs/promises'
import os from 'node:os'
import { join } from 'node:path'
import { Worker } from 'node:worker_threads'

import {
  PERF_OPEN_KINDS,
  PERF_STEPS,
  type PerfOpenKind,
  type PerfProgress,
  type PerfReport,
  type PerfStepId,
  type PerfTiming,
} from '../shared/perf-check'

export interface PerfCheckDeps {
  appVersion: string
  /** folder whose drive is measured for free space (userData) */
  dataDir: string
  /** parent of the scratch folder (os.tmpdir() in the app) */
  tempRoot: string
  sampleBytes(kind: PerfOpenKind): Promise<Buffer>
  /** cold start + open + export to PDF in a separate app process; profileDir is scratch userData */
  exportToPdf(input: string, outPath: string, profileDir: string): Promise<void>
  /** render every page of a PDF to images; resolves to the page count */
  renderPdfPages(pdfPath: string, outDir: string): Promise<number>
  appMemory(): { totalMB: number; processCount: number }
  cpuBenchmark(): Promise<number>
  freeMemMB(): number
  totalMemMB(): number
  diskFreeMB(dir: string): number | null
  cpuInfo(): { model: string; cores: number }
  osLabel(): string
  now(): number
  onProgress?(progress: PerfProgress): void
}

/** a sample that takes this long is reported as failed rather than waited for */
export const OPEN_TIMEOUT_MS = 120_000
const DISK_TEST_MB = 48
const FREE_MEM_SAMPLE_MS = 250

function message(err: unknown): string {
  const text = err instanceof Error ? err.message : String(err)
  // first line only: the report is a summary, not a log
  return text.split('\n')[0].slice(0, 160)
}

async function timed(fn: () => Promise<unknown>, now: () => number): Promise<PerfTiming> {
  const start = now()
  try {
    await fn()
    return { ms: now() - start }
  } catch (err) {
    return { ms: null, error: message(err) }
  }
}

export async function measureDiskWrite(dir: string, now: () => number): Promise<number> {
  const path = join(dir, 'disk-test.bin')
  const chunk = Buffer.alloc(4 * 1024 * 1024, 0x5a)
  const start = now()
  const fh = await open(path, 'w')
  try {
    for (let i = 0; i < DISK_TEST_MB / 4; i++) await fh.write(chunk)
    await fh.sync()
  } finally {
    await fh.close()
  }
  const seconds = Math.max(now() - start, 1) / 1000
  await rm(path, { force: true })
  return DISK_TEST_MB / seconds
}

export async function runPerfCheck(deps: PerfCheckDeps): Promise<PerfReport> {
  const { now } = deps
  const startedAt = now()
  let index = 0
  const step = (id: PerfStepId) =>
    deps.onProgress?.({ step: id, index: index++, total: PERF_STEPS.length })

  step('system')
  const cpuInfo = deps.cpuInfo()
  const system = {
    os: deps.osLabel(),
    arch: process.arch,
    cpuModel: cpuInfo.model,
    cpuCores: cpuInfo.cores,
    totalMemMB: Math.round(deps.totalMemMB()),
    freeMemMB: Math.round(deps.freeMemMB()),
    diskFreeMB: deps.diskFreeMB(deps.dataDir),
  }

  const work = await mkdtemp(join(deps.tempRoot, 'hyper-files-perf-'))
  try {
    step('cpu')
    let cpu: PerfTiming
    try {
      cpu = { ms: Math.round(await deps.cpuBenchmark()) }
    } catch (err) {
      cpu = { ms: null, error: message(err) }
    }

    step('disk')
    let disk: PerfReport['disk']
    if (system.diskFreeMB !== null && system.diskFreeMB < DISK_TEST_MB * 4) {
      disk = { writeMBps: null, error: 'not enough free space' }
    } else {
      try {
        disk = { writeMBps: await measureDiskWrite(work, now) }
      } catch (err) {
        disk = { writeMBps: null, error: message(err) }
      }
    }

    // the opens are where a 4 GB machine runs short; watch free RAM meanwhile
    let lowestFreeMB = deps.freeMemMB()
    const sampler = setInterval(() => {
      lowestFreeMB = Math.min(lowestFreeMB, deps.freeMemMB())
    }, FREE_MEM_SAMPLE_MS)
    const pdfs: string[] = []
    const opens = {} as Record<PerfOpenKind, PerfTiming>
    let pdf: PerfReport['pdf']
    try {
      for (const kind of PERF_OPEN_KINDS) {
        step(kind)
        const input = join(work, `sample.${kind}`)
        const out = join(work, `sample-${kind}.pdf`)
        try {
          await writeFile(input, await deps.sampleBytes(kind))
        } catch (err) {
          opens[kind] = { ms: null, error: message(err) }
          continue
        }
        opens[kind] = await timed(() => deps.exportToPdf(input, out, join(work, 'profile')), now)
        if (opens[kind].ms !== null && existsSync(out)) pdfs.push(out)
      }

      step('pdf')
      if (pdfs.length === 0) {
        pdf = { ms: null, error: 'no sample PDF was produced' }
      } else {
        let pages = 0
        const imagesDir = await mkdtemp(join(work, 'pages-'))
        const t = await timed(async () => {
          for (const p of pdfs) pages += await deps.renderPdfPages(p, imagesDir)
        }, now)
        pdf = { ...t, pages }
      }
    } finally {
      clearInterval(sampler)
    }
    lowestFreeMB = Math.min(lowestFreeMB, deps.freeMemMB())

    step('memory')
    const mem = deps.appMemory()
    return {
      appVersion: deps.appVersion,
      startedAt,
      durationMs: now() - startedAt,
      system,
      cpu,
      disk,
      opens,
      pdf,
      memory: {
        totalMB: Math.round(mem.totalMB),
        processCount: mem.processCount,
        lowestFreeMB: Math.round(lowestFreeMB),
      },
    }
  } finally {
    await rm(work, { recursive: true, force: true }).catch(() => {})
  }
}

// ── real-machine pieces (wired in main/index.ts) ────────────────────

/**
 * Fixed JavaScript workload: string building, hashing into a Map, sorting.
 * Runs in a worker so the main process keeps serving the UI. Best of three
 * runs, so a background hiccup does not decide the result.
 */
const CPU_WORKER_SOURCE = `
const { parentPort } = require('node:worker_threads')
function work() {
  const words = []
  for (let i = 0; i < 150000; i++) words.push('w' + (Math.imul(i, 2654435761) >>> 0).toString(36))
  const counts = new Map()
  for (const w of words) {
    const k = w.slice(0, 4)
    counts.set(k, (counts.get(k) || 0) + 1)
  }
  const sorted = words.slice().sort()
  let h = counts.size
  for (const w of sorted) h = (h * 31 + w.charCodeAt(w.length - 1)) | 0
  const objs = []
  for (let i = 0; i < 200000; i++) objs.push({ a: i, b: String(i), c: [i, i + 1] })
  objs.sort((x, y) => (y.a % 997) - (x.a % 997))
  return h + objs[0].a
}
let best = Infinity
let sink = 0
for (let r = 0; r < 3; r++) {
  const t = performance.now()
  sink += work()
  best = Math.min(best, performance.now() - t)
}
parentPort.postMessage({ ms: best, sink })
`

export function cpuBenchmark(): Promise<number> {
  return new Promise((resolve, reject) => {
    const worker = new Worker(CPU_WORKER_SOURCE, { eval: true })
    worker.once('message', (msg: { ms: number }) => {
      resolve(msg.ms)
      void worker.terminate()
    })
    worker.once('error', reject)
    worker.once('exit', (code) => {
      if (code !== 0) reject(new Error(`CPU test stopped (exit ${code})`))
    })
  })
}

export function diskFreeMB(dir: string): number | null {
  try {
    const s = statfsSync(dir)
    return Math.round((s.bavail * s.bsize) / (1024 * 1024))
  } catch {
    return null
  }
}

export function osLabel(): string {
  if (process.platform === 'win32') {
    // Windows 11 still reports 10.0; build 22000+ is 11
    const build = Number(os.release().split('.')[2] ?? 0)
    return `Windows ${build >= 22000 ? '11' : '10'} (build ${build})`
  }
  return `${os.type()} ${os.release()}`
}

export interface HeadlessSpawnOptions {
  execPath: string
  /** argv before the headless flags: Chromium switches, then the app path when unpacked */
  leadingArgs: string[]
  env: NodeJS.ProcessEnv
  timeoutMs?: number
}

/** Runs `<app> --headless-export <input> --to pdf --out <out> --json` and waits for it. */
export function spawnHeadlessExport(
  input: string,
  outPath: string,
  opts: HeadlessSpawnOptions,
): Promise<void> {
  return new Promise((resolve, reject) => {
    const child = spawn(
      opts.execPath,
      [...opts.leadingArgs, '--headless-export', input, '--to', 'pdf', '--out', outPath, '--json'],
      { env: opts.env, stdio: ['ignore', 'pipe', 'pipe'], windowsHide: true },
    )
    let stdout = ''
    child.stdout.on('data', (d: Buffer) => {
      if (stdout.length < 8192) stdout += d.toString()
    })
    child.stderr.resume()
    const timer = setTimeout(() => {
      child.kill()
      reject(
        new Error(`timed out after ${Math.round((opts.timeoutMs ?? OPEN_TIMEOUT_MS) / 1000)} s`),
      )
    }, opts.timeoutMs ?? OPEN_TIMEOUT_MS)
    child.once('error', (err) => {
      clearTimeout(timer)
      reject(err)
    })
    child.once('exit', (code) => {
      clearTimeout(timer)
      if (code === 0 && existsSync(outPath)) {
        resolve()
        return
      }
      // the --json envelope's error says what went wrong
      let reason = `exit ${code}`
      try {
        const line = stdout.trim().split('\n').pop() ?? ''
        const parsed = JSON.parse(line) as { error?: unknown }
        if (typeof parsed.error === 'string' && parsed.error) reason = parsed.error
      } catch {
        // not JSON: keep the exit code
      }
      reject(new Error(reason))
    })
  })
}
