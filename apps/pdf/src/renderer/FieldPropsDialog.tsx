/**
 * Properties of a form field placed with the form designer: its name, plus
 * the type's own settings (text: multiline; dropdown: options; radio: the
 * button's value) and whether filling it is required.
 */
import { useState } from 'react'
import type { ReactElement } from 'react'
import type { DrawingInput } from '../shared/ipc'
import { useI18n } from './i18n/locale'
import { useModalDialog } from './modal-dialog'

export type FieldDraft = Extract<DrawingInput, { kind: 'field' }>

export interface FieldProps {
  name: string
  options?: string[]
  exportValue?: string
  multiline: boolean
  required: boolean
}

export function FieldPropsDialog({
  field,
  nameTaken,
  onApply,
  onCancel,
}: {
  field: FieldDraft
  /** True when another field (in the file or pending) already uses this name */
  nameTaken: (name: string) => boolean
  onApply: (props: FieldProps) => void
  onCancel: () => void
}): ReactElement {
  const { t } = useI18n()
  const dialogRef = useModalDialog(onCancel)
  const [name, setName] = useState(field.name)
  const [options, setOptions] = useState((field.options ?? []).join('\n'))
  const [exportValue, setExportValue] = useState(field.exportValue ?? '')
  const [multiline, setMultiline] = useState(field.multiline === true)
  const [required, setRequired] = useState(field.required === true)

  const clean = name.trim().replace(/\./g, '_')
  const error = !clean
    ? t('fieldNameEmpty')
    : field.fieldType !== 'radio' && nameTaken(clean)
      ? t('fieldNameTaken')
      : null
  const apply = () => {
    if (error) return
    onApply({
      name: clean,
      ...(field.fieldType === 'dropdown'
        ? {
            options: options
              .split('\n')
              .map((o) => o.trim())
              .filter((o) => o.length > 0),
          }
        : {}),
      ...(field.fieldType === 'radio' ? { exportValue: exportValue.trim() || 'Choice' } : {}),
      multiline: field.fieldType === 'text' && multiline,
      required,
    })
  }

  return (
    <div className="pdf-modal-mask" onClick={onCancel}>
      <div
        ref={dialogRef}
        className="pdf-modal pdf-field-props"
        role="dialog"
        aria-modal="true"
        aria-label={t('fieldPropsTitle')}
        onClick={(e) => e.stopPropagation()}
        onKeyDown={(e) => {
          if (e.key === 'Enter' && !(e.target instanceof HTMLTextAreaElement)) apply()
        }}
      >
        <div className="pdf-modal-title">{t('fieldPropsTitle')}</div>
        <label className="pdf-field">
          <span>{field.fieldType === 'radio' ? t('fieldGroupName') : t('fieldName')}</span>
          <input
            className={`pdf-modal-input${error ? ' invalid' : ''}`}
            value={name}
            onChange={(e) => setName(e.target.value)}
          />
        </label>
        {error && <div className="pdf-field-props-error">{error}</div>}
        {field.fieldType === 'radio' && (
          <label className="pdf-field">
            <span>{t('fieldButtonValue')}</span>
            <input
              className="pdf-modal-input"
              value={exportValue}
              onChange={(e) => setExportValue(e.target.value)}
            />
          </label>
        )}
        {field.fieldType === 'dropdown' && (
          <label className="pdf-field">
            <span>{t('fieldOptions')}</span>
            <textarea
              className="pdf-modal-input pdf-field-props-options"
              value={options}
              placeholder={t('fieldOptionsHint')}
              onChange={(e) => setOptions(e.target.value)}
            />
          </label>
        )}
        {field.fieldType === 'text' && (
          <label className="pdf-modal-check">
            <input
              type="checkbox"
              checked={multiline}
              onChange={(e) => setMultiline(e.target.checked)}
            />
            {t('fieldMultiline')}
          </label>
        )}
        <label className="pdf-modal-check">
          <input
            type="checkbox"
            checked={required}
            onChange={(e) => setRequired(e.target.checked)}
          />
          {t('fieldRequired')}
        </label>
        <div className="pdf-modal-actions">
          <button className="pdf-modal-btn" onClick={onCancel}>
            {t('cancel')}
          </button>
          <button className="pdf-modal-btn primary" disabled={!!error} onClick={apply}>
            {t('ok')}
          </button>
        </div>
      </div>
    </div>
  )
}
