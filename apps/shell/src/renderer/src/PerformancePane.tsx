import { useEffect, useRef, useState } from 'react'
import {
  PERF_OPEN_KINDS,
  overallRating,
  perfTips,
  ratePerfReport,
  type PerfOpenKind,
  type PerfOverall,
  type PerfProgress,
  type PerfRating,
  type PerfReport,
  type PerfStepId,
  type PerfTip,
  type PerfTiming,
} from '../../shared/perf-check'
import { useI18n } from './locale'
import type { StringKey } from './locale'

// ── Settings → Performance: runs the offline self-test (main/perf-check.ts)
// and explains the result in plain words. The report lives in the main
// process, so reopening Settings shows the last run of this session.

const STEP_KEYS: Record<PerfStepId, StringKey> = {
  system: 'perfStepSystem',
  cpu: 'perfStepCpu',
  disk: 'perfStepDisk',
  docx: 'perfStepDocx',
  xlsx: 'perfStepXlsx',
  pptx: 'perfStepPptx',
  pdf: 'perfStepPdf',
  memory: 'perfStepMemory',
}

const OPEN_ROW_KEYS: Record<PerfOpenKind, StringKey> = {
  docx: 'perfRowDocx',
  xlsx: 'perfRowXlsx',
  pptx: 'perfRowPptx',
}

const RATING_KEYS: Record<PerfRating, StringKey> = {
  good: 'perfRateGood',
  ok: 'perfRateOk',
  slow: 'perfRateSlow',
  failed: 'perfRateFailed',
}

const OVERALL_KEYS: Record<PerfOverall, StringKey> = {
  good: 'perfOverallGood',
  incomplete: 'perfOverallIncomplete',
  ok: 'perfOverallOk',
  slow: 'perfOverallSlow',
}

const TIP_KEYS: Record<PerfTip, StringKey> = {
  closeApps: 'perfTipCloseApps',
  ramLow: 'perfTipRamLow',
  diskSlow: 'perfTipDiskSlow',
  diskFull: 'perfTipDiskFull',
  cpuSlow: 'perfTipCpuSlow',
  openSlow: 'perfTipOpenSlow',
}

interface Row {
  id: string
  labelKey: StringKey
  value: string
  desc?: string
  rating: PerfRating
}

export function PerformancePane() {
  const { t, dateLocale } = useI18n()
  const [report, setReport] = useState<PerfReport | null>(null)
  const [running, setRunning] = useState(false)
  const [progress, setProgress] = useState<PerfProgress | null>(null)
  const [failed, setFailed] = useState(false)
  const [copied, setCopied] = useState(false)
  const copiedTimer = useRef<number | undefined>(undefined)

  useEffect(() => {
    let alive = true
    void window.hyperFiles.getPerfCheck?.().then((r) => {
      if (alive && r) setReport(r)
    })
    const off = window.hyperFiles.onPerfCheckProgress?.((p) => {
      if (alive) setProgress(p)
    })
    return () => {
      alive = false
      off?.()
      window.clearTimeout(copiedTimer.current)
    }
  }, [])

  const run = () => {
    setRunning(true)
    setFailed(false)
    setProgress(null)
    void window.hyperFiles
      .runPerfCheck()
      .then(setReport)
      .catch(() => setFailed(true))
      .finally(() => {
        setRunning(false)
        setProgress(null)
      })
  }

  const copy = () => {
    void window.hyperFiles.copyPerfReport().then((ok) => {
      if (!ok) return
      setCopied(true)
      window.clearTimeout(copiedTimer.current)
      copiedTimer.current = window.setTimeout(() => setCopied(false), 2000)
    })
  }

  const num = (value: number, digits = 0) =>
    new Intl.NumberFormat(dateLocale, {
      minimumFractionDigits: digits,
      maximumFractionDigits: digits,
    }).format(value)
  const mb = (value: number | null) => {
    if (value === null) return '—'
    return value >= 1024 ? `${num(value / 1024, 1)} GB` : `${num(value)} MB`
  }
  const secs = (timing: PerfTiming, digits = 1) =>
    timing.ms === null ? '—' : `${num(timing.ms / 1000, digits)} s`

  const rows: Row[] = []
  let tips: PerfTip[] = []
  let overall: PerfOverall | null = null
  if (report) {
    const ratings = ratePerfReport(report)
    overall = overallRating(ratings)
    tips = perfTips(report, ratings)
    const s = report.system
    rows.push(
      { id: 'ram', labelKey: 'perfRowRam', value: mb(s.totalMemMB), rating: ratings.ram },
      {
        id: 'freeRam',
        labelKey: 'perfRowFreeRam',
        value: mb(s.freeMemMB),
        desc:
          report.memory.lowestFreeMB !== null
            ? t('perfLowestFree', { value: mb(report.memory.lowestFreeMB) })
            : undefined,
        rating: ratings.freeRam,
      },
      {
        id: 'cpu',
        labelKey: 'perfRowCpu',
        value: secs(report.cpu, 2),
        desc: `${s.cpuModel} · ${t('perfThreads', { n: s.cpuCores })}`,
        rating: ratings.cpu,
      },
      {
        id: 'disk',
        labelKey: 'perfRowDisk',
        value: report.disk.writeMBps === null ? '—' : `${num(report.disk.writeMBps)} MB/s`,
        rating: ratings.disk,
      },
      ...PERF_OPEN_KINDS.map((k): Row => ({
        id: k,
        labelKey: OPEN_ROW_KEYS[k],
        value: secs(report.opens[k]),
        desc: report.opens[k].error,
        rating: ratings[k],
      })),
      {
        id: 'pdf',
        labelKey: 'perfRowPdf',
        value: secs(report.pdf),
        desc: report.pdf.pages ? t('perfPages', { n: report.pdf.pages }) : report.pdf.error,
        rating: ratings.pdf,
      },
      {
        id: 'memory',
        labelKey: 'perfRowMemory',
        value: mb(report.memory.totalMB),
        desc: t('perfProcesses', { n: report.memory.processCount }),
        rating: ratings.memory,
      },
    )
  }

  return (
    <>
      <h3 className="set-pane-title">{t('setSecPerformance')}</h3>
      <div className="set-field">
        <div className="set-field-text">
          <div className="set-field-stack">
            <div className="set-field-label">{t('perfTitle')}</div>
            <div className="set-field-desc" aria-live="polite">
              {running
                ? progress
                  ? `${t('perfRunning', { step: t(STEP_KEYS[progress.step]) })} · ${t(
                      'perfProgress',
                      { n: progress.index + 1, total: progress.total },
                    )}`
                  : t('perfRunning', { step: t('perfStepSystem') })
                : failed
                  ? t('perfFailed')
                  : t('perfIntro')}
            </div>
          </div>
        </div>
        {report && !running && (
          <button className="set-btn" onClick={copy} data-testid="perf-copy">
            {copied ? t('perfCopied') : t('perfCopy')}
          </button>
        )}
        <button className="set-btn primary" disabled={running} onClick={run} data-testid="perf-run">
          {report ? t('perfRunAgain') : t('perfRun')}
        </button>
      </div>
      {report && overall && (
        <div className="perf-result" data-testid="perf-result" data-overall={overall}>
          <div className={`perf-verdict perf-${overall}`}>{t(OVERALL_KEYS[overall])}</div>
          {rows.map((row) => (
            <div className="set-field" key={row.id} data-perf-row={row.id}>
              <div className="set-field-text">
                <div className="set-field-stack">
                  <div className="set-field-label">{t(row.labelKey)}</div>
                  {row.desc && <div className="set-field-desc">{row.desc}</div>}
                </div>
                <div className="set-field-value">{row.value}</div>
              </div>
              <span className={`perf-badge perf-${row.rating}`}>{t(RATING_KEYS[row.rating])}</span>
            </div>
          ))}
          <div className="set-field-desc perf-note">{t('perfOpenNote')}</div>
          {tips.length > 0 && (
            <div className="perf-tips">
              <div className="set-field-label">{t('perfTipsTitle')}</div>
              <ul>
                {tips.map((tip) => (
                  <li key={tip}>{t(TIP_KEYS[tip])}</li>
                ))}
              </ul>
            </div>
          )}
        </div>
      )}
    </>
  )
}
