import { existsSync, readdirSync } from 'node:fs'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'

import { runPerfCheck, type PerfCheckDeps } from '../src/main/perf-check'
import {
  formatPerfReport,
  overallRating,
  perfTips,
  ratePerfReport,
  type PerfProgress,
  type PerfReport,
} from '../src/shared/perf-check'

function report(overrides: Partial<PerfReport> = {}): PerfReport {
  return {
    appVersion: '0.1.0',
    startedAt: Date.UTC(2026, 9, 8),
    durationMs: 42_000,
    system: {
      os: 'Windows 10 (build 19045)',
      arch: 'x64',
      cpuModel: 'Test CPU',
      cpuCores: 4,
      totalMemMB: 16_000,
      freeMemMB: 8_000,
      diskFreeMB: 100_000,
    },
    cpu: { ms: 300 },
    disk: { writeMBps: 400 },
    opens: { docx: { ms: 3_000 }, xlsx: { ms: 4_000 }, pptx: { ms: 3_500 } },
    pdf: { ms: 900, pages: 3 },
    memory: { totalMB: 600, processCount: 5, lowestFreeMB: 7_000 },
    ...overrides,
  }
}

/** a 4 GB budget laptop with a hard disk */
const lowSpec = report({
  system: {
    os: 'Windows 10 (build 19045)',
    arch: 'x64',
    cpuModel: 'Celeron N4020',
    cpuCores: 2,
    totalMemMB: 3_800,
    freeMemMB: 600,
    diskFreeMB: 1_500,
  },
  cpu: { ms: 2_000 },
  disk: { writeMBps: 25 },
  opens: { docx: { ms: 18_000 }, xlsx: { ms: 12_000 }, pptx: { ms: 9_000 } },
  pdf: { ms: 7_000, pages: 3 },
  memory: { totalMB: 700, processCount: 5, lowestFreeMB: 200 },
})

describe('perf ratings', () => {
  it('rates a fast machine good everywhere', () => {
    const ratings = ratePerfReport(report())
    expect(new Set(Object.values(ratings))).toEqual(new Set(['good']))
    expect(overallRating(ratings)).toBe('good')
    expect(perfTips(report(), ratings)).toEqual([])
  })

  it('counts a 4 GB machine (3.8 GB usable) as the supported floor, not below it', () => {
    expect(ratePerfReport(lowSpec).ram).toBe('ok')
  })

  it('calls a slow budget machine slow and explains why', () => {
    const ratings = ratePerfReport(lowSpec)
    expect(ratings).toMatchObject({
      freeRam: 'slow',
      cpu: 'slow',
      disk: 'slow',
      docx: 'slow',
      xlsx: 'ok',
      pptx: 'ok',
      pdf: 'slow',
    })
    expect(overallRating(ratings)).toBe('slow')
    expect(perfTips(lowSpec, ratings)).toEqual([
      'closeApps',
      'diskSlow',
      'diskFull',
      'cpuSlow',
      'openSlow',
    ])
  })

  it('reports an unfinished check as incomplete instead of good', () => {
    const r = report({
      opens: { docx: { ms: null, error: 'boom' }, xlsx: { ms: 1 }, pptx: { ms: 1 } },
    })
    const ratings = ratePerfReport(r)
    expect(ratings.docx).toBe('failed')
    expect(overallRating(ratings)).toBe('incomplete')
  })

  it('formats a plain-text report with every measurement', () => {
    const text = formatPerfReport(lowSpec)
    expect(text).toContain('Hypercube Office 0.1.0 performance check')
    expect(text).toContain('Overall: SLOW')
    expect(text).toContain('Celeron N4020 (2 threads)')
    expect(text).toMatch(/Open Word document\s+18\.00 s\s+\[slow\]/)
    expect(text).toMatch(/Render PDF pages\s+7\.00 s \(3 pages\)\s+\[slow\]/)
    expect(text).toMatch(/Lowest free RAM\s+200 MB/)
  })
})

describe('runPerfCheck', () => {
  let root: string
  beforeEach(async () => {
    root = await mkdtemp(join(tmpdir(), 'perf-test-'))
  })
  afterEach(async () => {
    await rm(root, { recursive: true, force: true })
  })

  function deps(overrides: Partial<PerfCheckDeps> = {}): PerfCheckDeps {
    let clock = 1_000
    return {
      appVersion: '0.1.0',
      dataDir: root,
      tempRoot: root,
      sampleBytes: async (kind) => Buffer.from(`sample ${kind}`),
      exportToPdf: async (_input, outPath) => {
        clock += 2_000
        await writeFile(outPath, '%PDF-1.7')
      },
      renderPdfPages: async () => {
        clock += 100
        return 2
      },
      appMemory: () => ({ totalMB: 512.4, processCount: 4 }),
      cpuBenchmark: async () => 321.6,
      freeMemMB: () => 4_000,
      totalMemMB: () => 8_192,
      diskFreeMB: () => 50_000,
      cpuInfo: () => ({ model: 'Fake CPU', cores: 8 }),
      osLabel: () => 'Windows 11 (build 22631)',
      now: () => clock,
      ...overrides,
    }
  }

  it('measures every step, reports progress in order and cleans up', async () => {
    const steps: PerfProgress[] = []
    const r = await runPerfCheck(deps({ onProgress: (p) => steps.push(p) }))
    expect(steps.map((s) => s.step)).toEqual([
      'system',
      'cpu',
      'disk',
      'docx',
      'xlsx',
      'pptx',
      'pdf',
      'memory',
    ])
    expect(steps.every((s) => s.total === 8)).toBe(true)
    expect(r.cpu.ms).toBe(322)
    expect(r.disk.writeMBps).toBeGreaterThan(0)
    expect(r.opens).toEqual({ docx: { ms: 2_000 }, xlsx: { ms: 2_000 }, pptx: { ms: 2_000 } })
    expect(r.pdf).toEqual({ ms: 300, pages: 6 })
    expect(r.memory).toEqual({ totalMB: 512, processCount: 4, lowestFreeMB: 4_000 })
    expect(r.system).toMatchObject({ cpuModel: 'Fake CPU', totalMemMB: 8_192, diskFreeMB: 50_000 })
    // the scratch folder (samples, PDFs, page images) is gone
    expect(readdirSync(root)).toEqual([])
  })

  it('keeps going when one sample fails and says why', async () => {
    const r = await runPerfCheck(
      deps({
        exportToPdf: async (input, outPath) => {
          if (input.endsWith('.xlsx')) throw new Error('sidecar missing\nstack…')
          await writeFile(outPath, '%PDF-1.7')
        },
      }),
    )
    expect(r.opens.xlsx).toEqual({ ms: null, error: 'sidecar missing' })
    expect(r.opens.docx.ms).not.toBeNull()
    expect(r.pdf.pages).toBe(4)
  })

  it('skips the disk test when the drive is nearly full', async () => {
    const r = await runPerfCheck(deps({ diskFreeMB: () => 100 }))
    expect(r.disk).toEqual({ writeMBps: null, error: 'not enough free space' })
  })

  it('passes a scratch profile folder to the headless child', async () => {
    const profiles: string[] = []
    await runPerfCheck(
      deps({
        exportToPdf: async (_input, outPath, profileDir) => {
          profiles.push(profileDir)
          await writeFile(outPath, '%PDF-1.7')
        },
      }),
    )
    expect(new Set(profiles).size).toBe(1)
    expect(profiles[0].startsWith(root)).toBe(true)
    expect(existsSync(profiles[0])).toBe(false)
  })
})
