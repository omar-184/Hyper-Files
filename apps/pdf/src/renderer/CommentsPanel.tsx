import { useMemo, useState } from 'react'
import { foldCase } from '@genoffice/ui'
import type { CommentEntry, CommentKind } from './comments-list'

export interface CommentsPanelLabels {
  title: string
  filter: string
  empty: string
  noMatch: string
  loading: string
  remove: string
  unsaved: string
  kind: Record<CommentKind, string>
  /** "Page {n}" */
  page: (n: number) => string
  /** "{n} replies" */
  replies: (n: number) => string
}

/**
 * Sidebar list of every comment and markup in the document (Acrobat's Comments
 * list): grouped by page, filterable, click to jump, delete in place.
 */
export function CommentsPanel({
  entries,
  loading,
  readOnly,
  labels,
  pageNumber,
  formatTime,
  onSelect,
  onDelete,
}: {
  entries: CommentEntry[]
  /** Saved annotations are still being read from some pages */
  loading: boolean
  readOnly: boolean
  labels: CommentsPanelLabels
  /** 1-based on-screen page number of an original page index */
  pageNumber: (pageIndex: number) => number
  formatTime: (ms: number) => string
  onSelect: (entry: CommentEntry) => void
  onDelete: (entry: CommentEntry) => void
}) {
  const [query, setQuery] = useState('')
  const shown = useMemo(() => {
    const q = foldCase(query.trim())
    if (!q) return entries
    return entries.filter((e) =>
      [e.text, e.author, labels.kind[e.kind]].some((s) => foldCase(s).includes(q)),
    )
  }, [entries, query, labels])

  const groups: { page: number; items: CommentEntry[] }[] = []
  for (const e of shown) {
    const page = pageNumber(e.pageIndex)
    const last = groups[groups.length - 1]
    if (last?.page === page) last.items.push(e)
    else groups.push({ page, items: [e] })
  }

  return (
    <div className="pdf-comments">
      <input
        className="pdf-comments-filter"
        type="search"
        placeholder={labels.filter}
        aria-label={labels.filter}
        value={query}
        onChange={(e) => setQuery(e.target.value)}
      />
      <div className="pdf-comments-list" role="list" aria-label={labels.title}>
        {groups.map((g) => (
          <div key={g.page} className="pdf-comments-group">
            <div className="pdf-comments-page">{labels.page(g.page)}</div>
            {g.items.map((e) => (
              <div key={e.key} className="pdf-comment-row" role="listitem">
                <button type="button" className="pdf-comment-main" onClick={() => onSelect(e)}>
                  <span className="pdf-comment-head">
                    <span className={`pdf-comment-kind pdf-comment-kind-${e.kind}`}>
                      {labels.kind[e.kind]}
                    </span>
                    {e.author && <span className="pdf-comment-author">{e.author}</span>}
                    {e.timeMs !== null && (
                      <span className="pdf-comment-time">{formatTime(e.timeMs)}</span>
                    )}
                    {e.pending && <span className="pdf-comment-unsaved">{labels.unsaved}</span>}
                  </span>
                  {e.text && <span className="pdf-comment-text">{e.text}</span>}
                  {e.replies > 0 && (
                    <span className="pdf-comment-replies">{labels.replies(e.replies)}</span>
                  )}
                </button>
                {!readOnly && (
                  <button
                    type="button"
                    className="pdf-comment-delete"
                    aria-label={labels.remove}
                    data-tip={labels.remove}
                    onClick={() => onDelete(e)}
                  >
                    ×
                  </button>
                )}
              </div>
            ))}
          </div>
        ))}
        {shown.length === 0 && !loading && (
          <div className="pdf-comments-empty">
            {entries.length === 0 ? labels.empty : labels.noMatch}
          </div>
        )}
        {loading && <div className="pdf-comments-empty">{labels.loading}</div>}
      </div>
    </div>
  )
}
