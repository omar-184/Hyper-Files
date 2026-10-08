/**
 * Settings → Performance: a short, fully offline self-test that tells the user
 * how Hyper-Files runs on this computer. The main process measures
 * (main/perf-check.ts); this module holds the shared shapes and the pure
 * rating/report logic so both sides and the unit tests agree on them.
 *
 * Thresholds target the project's floor: Windows 10/11 on a 4 GB machine.
 * Nothing measured here leaves the computer; the report is only copied to the
 * clipboard when the user asks for it.
 */

export type PerfRating = 'good' | 'ok' | 'slow' | 'failed'

/** sample documents the check opens, each in a fresh headless app process */
export const PERF_OPEN_KINDS = ['docx', 'xlsx', 'pptx'] as const
export type PerfOpenKind = (typeof PERF_OPEN_KINDS)[number]

/** steps in run order; the renderer shows the one in progress */
export const PERF_STEPS = [
  'system',
  'cpu',
  'disk',
  'docx',
  'xlsx',
  'pptx',
  'pdf',
  'memory',
] as const
export type PerfStepId = (typeof PERF_STEPS)[number]

export interface PerfSystemInfo {
  os: string
  arch: string
  cpuModel: string
  cpuCores: number
  totalMemMB: number
  /** free RAM when the check started */
  freeMemMB: number
  /** free space on the drive that holds the app's data; null when unknown */
  diskFreeMB: number | null
}

export interface PerfTiming {
  /** wall-clock time; null when the step failed */
  ms: number | null
  error?: string
}

export interface PerfMemory {
  /** working set of every Hyper-Files process (main, windows, helpers) */
  totalMB: number
  processCount: number
  /** lowest free system RAM seen while the sample documents were open */
  lowestFreeMB: number | null
}

export interface PerfReport {
  appVersion: string
  startedAt: number
  durationMs: number
  system: PerfSystemInfo
  /** fixed JavaScript workload; lower is faster */
  cpu: PerfTiming
  /** sequential write of a temporary file, flushed to disk, in MB/s */
  disk: { writeMBps: number | null; error?: string }
  /** start a fresh app process, open the sample and render it to PDF */
  opens: Record<PerfOpenKind, PerfTiming>
  /** render every page of a sample PDF to images */
  pdf: PerfTiming & { pages?: number }
  memory: PerfMemory
}

export type PerfRatings = {
  ram: PerfRating
  freeRam: PerfRating
  cpu: PerfRating
  disk: PerfRating
  pdf: PerfRating
  memory: PerfRating
} & Record<PerfOpenKind, PerfRating>

export type PerfTip = 'closeApps' | 'ramLow' | 'diskSlow' | 'diskFull' | 'cpuSlow' | 'openSlow'

/** progress pushed to the renderer while the check runs */
export interface PerfProgress {
  step: PerfStepId
  index: number
  total: number
}

// ── thresholds ──────────────────────────────────────────────────────

/** Windows reports ~3.8 GB usable on a 4 GB machine; that still counts as the 4 GB floor */
export const RAM_GOOD_MB = 7_500
export const RAM_OK_MB = 3_500
export const FREE_RAM_GOOD_MB = 1_500
export const FREE_RAM_OK_MB = 700
/** CPU workload: about 150 ms on a CI server core, roughly 3-4x that on a budget Celeron */
export const CPU_GOOD_MS = 350
export const CPU_OK_MS = 800
/** SSDs write hundreds of MB/s; laptop HDDs and cheap eMMC sit around 30-100 */
export const DISK_GOOD_MBPS = 150
export const DISK_OK_MBPS = 40
export const DISK_FREE_LOW_MB = 2_000
/** cold start + open + render, in a fresh process: this is the "double-click a file" time */
export const OPEN_GOOD_MS = 6_000
export const OPEN_OK_MS = 15_000
export const PDF_GOOD_MS = 2_500
export const PDF_OK_MS = 6_000
/** share of physical RAM the app itself holds with only Home open */
export const MEMORY_GOOD_SHARE = 0.15
export const MEMORY_OK_SHARE = 0.3

function lowerIsBetter(value: number | null, good: number, ok: number): PerfRating {
  if (value === null || !Number.isFinite(value)) return 'failed'
  if (value <= good) return 'good'
  if (value <= ok) return 'ok'
  return 'slow'
}

function higherIsBetter(value: number | null, good: number, ok: number): PerfRating {
  if (value === null || !Number.isFinite(value)) return 'failed'
  if (value >= good) return 'good'
  if (value >= ok) return 'ok'
  return 'slow'
}

export function ratePerfReport(r: PerfReport): PerfRatings {
  const opens = Object.fromEntries(
    PERF_OPEN_KINDS.map((k) => [k, lowerIsBetter(r.opens[k].ms, OPEN_GOOD_MS, OPEN_OK_MS)]),
  ) as Record<PerfOpenKind, PerfRating>
  const total = r.system.totalMemMB
  return {
    ram: higherIsBetter(total, RAM_GOOD_MB, RAM_OK_MB),
    freeRam: higherIsBetter(r.system.freeMemMB, FREE_RAM_GOOD_MB, FREE_RAM_OK_MB),
    cpu: lowerIsBetter(r.cpu.ms, CPU_GOOD_MS, CPU_OK_MS),
    disk: higherIsBetter(r.disk.writeMBps, DISK_GOOD_MBPS, DISK_OK_MBPS),
    pdf: lowerIsBetter(r.pdf.ms, PDF_GOOD_MS, PDF_OK_MS),
    memory:
      total > 0
        ? lowerIsBetter(r.memory.totalMB / total, MEMORY_GOOD_SHARE, MEMORY_OK_SHARE)
        : 'failed',
    ...opens,
  }
}

export type PerfOverall = 'good' | 'ok' | 'slow' | 'incomplete'

/**
 * Worst rating wins. A failed measurement says nothing about speed, so it only
 * turns the verdict into 'incomplete' when nothing else came out slow.
 */
export function overallRating(ratings: PerfRatings): PerfOverall {
  const values = Object.values(ratings)
  if (values.includes('slow')) return 'slow'
  if (values.includes('failed')) return 'incomplete'
  if (values.includes('ok')) return 'ok'
  return 'good'
}

/** plain-language advice for what came out slow, most useful first */
export function perfTips(r: PerfReport, ratings: PerfRatings): PerfTip[] {
  const tips: PerfTip[] = []
  if (ratings.freeRam !== 'good') tips.push('closeApps')
  if (ratings.ram === 'slow') tips.push('ramLow')
  if (ratings.disk === 'slow') tips.push('diskSlow')
  if (r.system.diskFreeMB !== null && r.system.diskFreeMB < DISK_FREE_LOW_MB) tips.push('diskFull')
  if (ratings.cpu === 'slow') tips.push('cpuSlow')
  if (PERF_OPEN_KINDS.some((k) => ratings[k] === 'slow') || ratings.pdf === 'slow')
    tips.push('openSlow')
  return tips
}

function fmtMs(t: PerfTiming): string {
  return t.ms === null ? `failed${t.error ? ` (${t.error})` : ''}` : `${(t.ms / 1000).toFixed(2)} s`
}

function fmtMB(mb: number | null): string {
  if (mb === null) return 'unknown'
  return mb >= 1024 ? `${(mb / 1024).toFixed(1)} GB` : `${Math.round(mb)} MB`
}

const OPEN_LABELS: Record<PerfOpenKind, string> = {
  docx: 'Open Word document',
  xlsx: 'Open Excel workbook',
  pptx: 'Open PowerPoint deck',
}

/**
 * The plain-text report behind "Copy report". English on purpose: it is meant
 * for bug reports and forum posts, where the reader may not share the UI language.
 */
export function formatPerfReport(r: PerfReport): string {
  const ratings = ratePerfReport(r)
  const s = r.system
  const rows: [string, string, PerfRating][] = [
    ['Installed RAM', fmtMB(s.totalMemMB), ratings.ram],
    ['Free RAM at start', fmtMB(s.freeMemMB), ratings.freeRam],
    ['CPU test', fmtMs(r.cpu), ratings.cpu],
    [
      'Disk write speed',
      r.disk.writeMBps === null
        ? `failed${r.disk.error ? ` (${r.disk.error})` : ''}`
        : `${Math.round(r.disk.writeMBps)} MB/s`,
      ratings.disk,
    ],
    ...PERF_OPEN_KINDS.map((k): [string, string, PerfRating] => [
      OPEN_LABELS[k],
      fmtMs(r.opens[k]),
      ratings[k],
    ]),
    [
      'Render PDF pages',
      r.pdf.ms !== null && r.pdf.pages ? `${fmtMs(r.pdf)} (${r.pdf.pages} pages)` : fmtMs(r.pdf),
      ratings.pdf,
    ],
    [
      'App memory now',
      `${fmtMB(r.memory.totalMB)} in ${r.memory.processCount} processes`,
      ratings.memory,
    ],
  ]
  const width = Math.max(...rows.map(([label]) => label.length))
  const lines = [
    `Hyper-Files ${r.appVersion} performance check`,
    `Date: ${new Date(r.startedAt).toISOString()}`,
    `System: ${s.os} ${s.arch}, ${s.cpuModel} (${s.cpuCores} threads)`,
    `Free disk space: ${fmtMB(s.diskFreeMB)}`,
    `Overall: ${overallRating(ratings).toUpperCase()}`,
    '',
    ...rows.map(([label, value, rating]) => `${label.padEnd(width)}  ${value}  [${rating}]`),
  ]
  if (r.memory.lowestFreeMB !== null)
    lines.push(`${'Lowest free RAM'.padEnd(width)}  ${fmtMB(r.memory.lowestFreeMB)}`)
  lines.push(
    '',
    `Check took ${(r.durationMs / 1000).toFixed(1)} s. Measured on this computer only.`,
  )
  return lines.join('\n')
}
