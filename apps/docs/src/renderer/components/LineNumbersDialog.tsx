import { useState } from 'react'
import type { LineNumbering, SectionSettings } from '@genoffice/docx-engine'
import { useI18n, type StringKey } from '../i18n/locale'
import { LengthInput } from './LengthInput'
import { Frame } from './TableDialogs'

/** Layout ▸ Line Numbers: Word's menu modes and the Line Numbers options dialog */

export type LineNumberMode = 'none' | LineNumbering['restart']

export const LINE_NUMBER_MODES: Array<[LineNumberMode, StringKey]> = [
  ['none', 'ribbonLineNumbersNone'],
  ['continuous', 'ribbonLineNumbersContinuous'],
  ['newPage', 'ribbonLineNumbersEachPage'],
  ['newSection', 'ribbonLineNumbersEachSection'],
]

export function lineNumberModeOf(section: SectionSettings): LineNumberMode {
  return section.lineNumbers?.restart ?? 'none'
}

/** the section's numbering with only the restart rule changed (Word keeps start / count by) */
export function lineNumbersFor(
  section: SectionSettings,
  mode: LineNumberMode,
): LineNumbering | undefined {
  if (mode === 'none') return undefined
  return { ...(section.lineNumbers ?? { countBy: 1, start: 1 }), restart: mode }
}

const clampInt = (raw: string, lo: number, hi: number) =>
  Math.min(hi, Math.max(lo, Math.round(Number(raw) || lo)))

export function LineNumbersDialog({
  section,
  onApply,
  onClose,
}: {
  section: SectionSettings
  onApply: (next: LineNumbering | undefined) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const current = section.lineNumbers
  const [on, setOn] = useState(current !== undefined)
  const [start, setStart] = useState(current?.start ?? 1)
  const [countBy, setCountBy] = useState(current?.countBy ?? 1)
  const [distance, setDistance] = useState<number | null>(current?.distance ?? null)
  const [restart, setRestart] = useState<LineNumbering['restart']>(current?.restart ?? 'newPage')
  const restarts = LINE_NUMBER_MODES.filter(
    (m): m is [LineNumbering['restart'], StringKey] => m[0] !== 'none',
  )
  return (
    <Frame
      className="line-numbers-dialog"
      title={t('ribbonLineNumbers')}
      onClose={onClose}
      onOk={() => {
        onApply(
          on ? { countBy, start, restart, ...(distance !== null ? { distance } : {}) } : undefined,
        )
        onClose()
      }}
    >
      <label className="table-dialog-check">
        <input type="checkbox" checked={on} onChange={(e) => setOn(e.target.checked)} />
        {t('ribbonLineNumbers')}
      </label>
      <fieldset className="line-numbers-fields" disabled={!on}>
        <div className="table-dialog-grid">
          <label>
            {t('ribbonLineNumbersStartAt')}
            <input
              type="number"
              min={1}
              max={32767}
              value={start}
              onChange={(e) => setStart(clampInt(e.target.value, 1, 32767))}
            />
          </label>
          <label>
            {t('ribbonLineNumbersCountBy')}
            <input
              type="number"
              min={1}
              max={100}
              value={countBy}
              onChange={(e) => setCountBy(clampInt(e.target.value, 1, 100))}
            />
          </label>
          <label>
            {t('ribbonLineNumbersFromText')}
            <LengthInput
              value={distance}
              allowEmpty
              live
              min={0}
              max={1440 * 2}
              placeholder={t('ribbonAuto')}
              ariaLabel={t('ribbonLineNumbersFromText')}
              onCommit={setDistance}
            />
          </label>
        </div>
        <div
          className="table-dialog-radios"
          role="radiogroup"
          aria-label={t('ribbonLineNumbersNumbering')}
        >
          {restarts.map(([mode, labelKey]) => (
            <label key={mode}>
              <input
                type="radio"
                name="line-numbers-restart"
                checked={restart === mode}
                onChange={() => setRestart(mode)}
              />
              {t(labelKey)}
            </label>
          ))}
        </div>
      </fieldset>
    </Frame>
  )
}
