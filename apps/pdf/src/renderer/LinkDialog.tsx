/**
 * Target of a link area: a page of this document or a web address.
 */
import { useState } from 'react'
import type { ReactElement } from 'react'
import { useI18n } from './i18n/locale'
import { useModalDialog } from './modal-dialog'

export type LinkTarget = { url: string } | { page: number }

/** Accept what people type: a bare domain gets https://, a bare address mailto:.
    Only web and mail links are allowed. Null = not a usable address. */
export function normalizeLinkUrl(raw: string): string | null {
  const text = raw.trim()
  if (!text || /\s/.test(text)) return null
  const scheme = /^([a-z][a-z0-9+.-]*):/i.exec(text)?.[1]?.toLowerCase()
  if (scheme) {
    if (scheme === 'mailto') return text.length > 7 ? text : null
    if (scheme !== 'http' && scheme !== 'https') return null
    try {
      return new URL(text).href
    } catch {
      return null
    }
  }
  if (/^[^@/]+@[^@/]+\.[^@/]+$/.test(text)) return `mailto:${text}`
  try {
    return new URL(`https://${text}`).href
  } catch {
    return null
  }
}

export function LinkDialog({
  initial,
  pageCount,
  onApply,
  onCancel,
}: {
  initial: LinkTarget
  pageCount: number
  onApply: (target: LinkTarget) => void
  onCancel: () => void
}): ReactElement {
  const { t } = useI18n()
  const dialogRef = useModalDialog(onCancel)
  const [mode, setMode] = useState<'page' | 'url'>('url' in initial ? 'url' : 'page')
  const [page, setPage] = useState(String('page' in initial ? initial.page : 1))
  const [url, setUrl] = useState('url' in initial ? initial.url : '')

  const pageNo = Number(page)
  const pageOk = Number.isInteger(pageNo) && pageNo >= 1 && pageNo <= pageCount
  const cleanUrl = normalizeLinkUrl(url)
  const ok = mode === 'page' ? pageOk : cleanUrl !== null
  const apply = () => {
    if (!ok) return
    onApply(mode === 'page' ? { page: pageNo } : { url: cleanUrl! })
  }

  return (
    <div className="pdf-modal-mask" onClick={onCancel}>
      <div
        ref={dialogRef}
        className="pdf-modal pdf-link-dialog"
        role="dialog"
        aria-modal="true"
        aria-label={t('linkTitle')}
        onClick={(e) => e.stopPropagation()}
        onKeyDown={(e) => {
          if (e.key === 'Enter') apply()
        }}
      >
        <div className="pdf-modal-title">{t('linkTitle')}</div>
        <label className="pdf-modal-check">
          <input type="radio" checked={mode === 'url'} onChange={() => setMode('url')} />
          {t('linkToWeb')}
        </label>
        <input
          className={`pdf-modal-input${mode === 'url' && url && !cleanUrl ? ' invalid' : ''}`}
          aria-label={t('linkToWeb')}
          placeholder="https://"
          value={url}
          autoFocus={mode === 'url'}
          disabled={mode !== 'url'}
          onChange={(e) => setUrl(e.target.value)}
        />
        <label className="pdf-modal-check">
          <input type="radio" checked={mode === 'page'} onChange={() => setMode('page')} />
          {t('linkToPage')}
        </label>
        <input
          className={`pdf-modal-input${mode === 'page' && !pageOk ? ' invalid' : ''}`}
          aria-label={t('linkToPage')}
          type="number"
          min={1}
          max={pageCount}
          value={page}
          autoFocus={mode === 'page'}
          disabled={mode !== 'page'}
          onChange={(e) => setPage(e.target.value)}
        />
        <div className="pdf-modal-actions">
          <button className="pdf-modal-btn" onClick={onCancel}>
            {t('cancel')}
          </button>
          <button className="pdf-modal-btn primary" disabled={!ok} onClick={apply}>
            {t('ok')}
          </button>
        </div>
      </div>
    </div>
  )
}
