import { useMemo, useState } from 'react'
import type { Editor } from '@tiptap/core'
import { useI18n } from '../i18n/locale'
import {
  convertTableToText,
  convertTextToTable,
  guessHasHeader,
  sortColumnLabels,
  sortRefusal,
  sortTableRows,
  tableHasUnflattenable,
  textToTablePlan,
  type CellSeparator,
  type SortKey,
  type SortKeyType,
} from '../editor/table-data'
import { Frame, runAndClose } from './TableDialogs'

/** Word's Table Layout ▸ Data dialogs: Sort, Convert to Text, and Insert ▸
 *  Table ▸ Convert Text to Table. */

interface DialogProps {
  editor: Editor
  onClose: () => void
}

interface KeyRow {
  /** -1 = unused "Then by" row */
  column: number
  type: SortKeyType
  descending: boolean
}

export function SortDialog({ editor, onClose }: DialogProps) {
  const { t } = useI18n()
  const [hasHeader, setHasHeader] = useState(() => guessHasHeader(editor.state))
  const [caseSensitive, setCaseSensitive] = useState(false)
  const [keys, setKeys] = useState<KeyRow[]>([
    { column: 0, type: 'text', descending: false },
    { column: -1, type: 'text', descending: false },
    { column: -1, type: 'text', descending: false },
  ])
  const labels = useMemo(() => sortColumnLabels(editor.state, hasHeader), [editor, hasHeader])
  const refusal = sortRefusal(editor.state, hasHeader)
  const patch = (i: number, next: Partial<KeyRow>) =>
    setKeys((prev) => prev.map((k, j) => (j === i ? { ...k, ...next } : k)))
  const active: SortKey[] = keys.filter((k) => k.column >= 0)
  const types: Array<[SortKeyType, string]> = [
    ['text', t('ribbonSortTypeText')],
    ['number', t('ribbonSortTypeNumber')],
    ['date', t('ribbonSortTypeDate')],
  ]
  return (
    <Frame
      className="table-sort-dialog"
      title={t('ribbonSort')}
      onClose={onClose}
      okDisabled={refusal !== null || active.length === 0}
      onOk={() =>
        runAndClose(editor, sortTableRows({ keys: active, hasHeader, caseSensitive }), onClose)
      }
    >
      {keys.map((key, i) => (
        <fieldset key={i} className="table-sort-key" disabled={i > 0 && keys[i - 1].column < 0}>
          <legend>{i === 0 ? t('ribbonSortBy') : t('ribbonSortThenBy')}</legend>
          <select
            aria-label={i === 0 ? t('ribbonSortBy') : t('ribbonSortThenBy')}
            value={key.column}
            onChange={(e) => {
              const column = Number(e.target.value)
              // clearing a key clears the ones after it, as Word does
              setKeys((prev) =>
                prev.map((k, j) =>
                  j === i ? { ...k, column } : j > i && column < 0 ? { ...k, column: -1 } : k,
                ),
              )
            }}
          >
            {i > 0 && <option value={-1}>{t('ribbonSortKeyNone')}</option>}
            {labels.map((label, c) => (
              <option key={c} value={c}>
                {label ?? t('ribbonSortColumnN', { n: String(c + 1) })}
              </option>
            ))}
          </select>
          <select
            aria-label={t('ribbonSortType')}
            value={key.type}
            disabled={key.column < 0}
            onChange={(e) => patch(i, { type: e.target.value as SortKeyType })}
          >
            {types.map(([value, label]) => (
              <option key={value} value={value}>
                {label}
              </option>
            ))}
          </select>
          <span className="table-sort-order" role="radiogroup">
            <label>
              <input
                type="radio"
                name={`sort-order-${i}`}
                checked={!key.descending}
                disabled={key.column < 0}
                onChange={() => patch(i, { descending: false })}
              />
              {t('ribbonSortAscending')}
            </label>
            <label>
              <input
                type="radio"
                name={`sort-order-${i}`}
                checked={key.descending}
                disabled={key.column < 0}
                onChange={() => patch(i, { descending: true })}
              />
              {t('ribbonSortDescending')}
            </label>
          </span>
        </fieldset>
      ))}
      <label className="table-dialog-check">
        <input
          type="checkbox"
          checked={hasHeader}
          onChange={(e) => setHasHeader(e.target.checked)}
        />
        {t('ribbonSortHasHeader')}
      </label>
      <label className="table-dialog-check">
        <input
          type="checkbox"
          checked={caseSensitive}
          onChange={(e) => setCaseSensitive(e.target.checked)}
        />
        {t('ribbonSortCaseSensitive')}
      </label>
      {refusal === 'mergedRows' && <p className="modal-desc">{t('ribbonSortMerged')}</p>}
      {refusal === 'nothingToSort' && <p className="modal-desc">{t('ribbonSortNothing')}</p>}
    </Frame>
  )
}

type SeparatorChoice = 'paragraph' | 'tab' | 'comma' | 'other'

function SeparatorPicker({
  name,
  choice,
  other,
  onChoice,
  onOther,
}: {
  name: string
  choice: SeparatorChoice
  other: string
  onChoice: (c: SeparatorChoice) => void
  onOther: (s: string) => void
}) {
  const { t } = useI18n()
  const options: Array<[SeparatorChoice, string]> = [
    ['paragraph', t('ribbonSeparatorParagraphs')],
    ['tab', t('ribbonSeparatorTabs')],
    ['comma', t('ribbonSeparatorCommas')],
    ['other', t('ribbonSeparatorOther')],
  ]
  return (
    <fieldset className="table-separator-picker">
      <legend>{t('ribbonSeparateTextWith')}</legend>
      <div className="table-dialog-radios" role="radiogroup">
        {options.map(([value, label]) => (
          <label key={value}>
            <input
              type="radio"
              name={name}
              checked={choice === value}
              onChange={() => onChoice(value)}
            />
            {label}
            {value === 'other' && (
              <input
                type="text"
                className="table-separator-other"
                maxLength={1}
                value={other}
                aria-label={t('ribbonSeparatorOther')}
                onFocus={() => onChoice('other')}
                onChange={(e) => onOther(e.target.value)}
              />
            )}
          </label>
        ))}
      </div>
    </fieldset>
  )
}

function separatorOf(choice: SeparatorChoice, other: string): CellSeparator {
  return choice === 'other' ? { other: other || '-' } : choice
}

export function ConvertToTextDialog({ editor, onClose }: DialogProps) {
  const { t } = useI18n()
  const [choice, setChoice] = useState<SeparatorChoice>('tab')
  const [other, setOther] = useState('-')
  const lossy = tableHasUnflattenable(editor.state)
  return (
    <Frame
      className="table-to-text-dialog"
      title={t('ribbonConvertTableToText')}
      onClose={onClose}
      okDisabled={choice === 'other' && !other}
      onOk={() => runAndClose(editor, convertTableToText(separatorOf(choice, other)), onClose)}
    >
      <SeparatorPicker
        name="table-to-text-sep"
        choice={choice}
        other={other}
        onChoice={setChoice}
        onOther={setOther}
      />
      {lossy && <p className="modal-desc">{t('ribbonConvertDropsNested')}</p>}
    </Frame>
  )
}

export function TextToTableDialog({ editor, onClose }: DialogProps) {
  const { t } = useI18n()
  const [choice, setChoice] = useState<SeparatorChoice>(() =>
    // Word's guess: tabs when the selection holds any, else paragraph marks
    /\t/.test(
      editor.state.doc.textBetween(editor.state.selection.from, editor.state.selection.to, '\n'),
    )
      ? 'tab'
      : 'paragraph',
  )
  const [other, setOther] = useState('-')
  const [columns, setColumns] = useState<number | null>(null)
  const sep = separatorOf(choice, other)
  const natural = textToTablePlan(editor.state, sep)
  const plan = textToTablePlan(editor.state, sep, columns ?? natural?.cols)
  return (
    <Frame
      className="text-to-table-dialog"
      title={t('ribbonConvertTextToTableTitle')}
      onClose={onClose}
      okDisabled={!plan}
      onOk={() => runAndClose(editor, convertTextToTable(sep, plan?.cols), onClose)}
    >
      {plan ? (
        <>
          <div className="table-dialog-grid">
            <label>
              {t('ribbonTableColsLabel')}
              <input
                type="number"
                min={1}
                max={63}
                value={plan.cols}
                onChange={(e) =>
                  setColumns(Math.min(63, Math.max(1, Math.round(Number(e.target.value) || 1))))
                }
              />
            </label>
            <p className="modal-desc">{t('ribbonTextToTableRowsN', { n: String(plan.rows) })}</p>
          </div>
          <SeparatorPicker
            name="text-to-table-sep"
            choice={choice}
            other={other}
            onChoice={(c) => {
              setChoice(c)
              setColumns(null)
            }}
            onOther={(s) => {
              setOther(s)
              setColumns(null)
            }}
          />
        </>
      ) : (
        <p className="modal-desc">{t('ribbonTextToTableNeedsText')}</p>
      )}
    </Frame>
  )
}
