import { useEffect, useState } from 'react'
import type { ReactElement } from 'react'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import { pdfRectToCss } from './annotations'
import type { PageGeom } from './annotations'
import type { SavedLinkAnnot } from './edit-state'

interface LinkItem {
  rect: [number, number, number, number]
  url?: string
  dest?: unknown
  /** PDF object number, when the link is an indirect object (deletable) */
  objNum?: number
}

interface RawLinkAnnotation {
  id?: string
  subtype?: string
  rect?: number[]
  url?: string
  dest?: unknown
  action?: string
}

/** Max link overlays per page: a hostile PDF can carry thousands of annots. */
export const MAX_PAGE_LINKS = 500

function isFiniteRect(rect: unknown): rect is [number, number, number, number] {
  return (
    Array.isArray(rect) &&
    rect.length === 4 &&
    rect.every((n) => typeof n === 'number' && Number.isFinite(n))
  )
}

/**
 * Filter raw pdfjs annotations down to renderable links: Link subtype with a
 * finite rect and a target, capped per page. Exported for tests.
 */
export function collectPageLinks(annots: RawLinkAnnotation[]): LinkItem[] {
  const out: LinkItem[] = []
  for (const a of annots) {
    if (out.length >= MAX_PAGE_LINKS) break
    if (a.subtype !== 'Link' || (!a.url && !a.dest)) continue
    if (!isFiniteRect(a.rect)) continue
    const objNum = /^(\d+)R$/.exec(a.id ?? '')
    out.push({
      rect: a.rect,
      url: a.url,
      dest: a.dest,
      ...(objNum ? { objNum: Number(objNum[1]) } : {}),
    })
  }
  return out
}

/** Link annot hit areas: external links open a new window (main process routes to shell.openExternal); internal dests jump pages */
export function LinkLayer({
  doc,
  pageNo,
  geom,
  scale,
  onGoToDest,
  hidden,
  editing,
}: {
  doc: PDFDocumentProxy
  pageNo: number
  geom: PageGeom
  scale: number
  onGoToDest: (dest: unknown) => void
  /** Object numbers of links pending deletion: not shown */
  hidden?: ReadonlySet<number>
  /** Link tool on: links show as boxes with a remove button instead of navigating */
  editing?: { removeLabel: string; onRemove: (link: SavedLinkAnnot) => void }
}): ReactElement | null {
  const [links, setLinks] = useState<LinkItem[] | null>(null)

  useEffect(() => {
    let cancelled = false
    void (async () => {
      const page = await doc.getPage(pageNo)
      const annots = (await page.getAnnotations()) as RawLinkAnnotation[]
      if (cancelled) return
      setLinks(collectPageLinks(annots))
    })()
    return () => {
      cancelled = true
    }
  }, [doc, pageNo])

  const shown = links?.filter((l) => l.objNum === undefined || !hidden?.has(l.objNum))
  if (!shown || shown.length === 0) return null

  if (editing) {
    return (
      <div className="pdf-link-layer is-editing">
        {shown.map((l, i) => (
          <div
            key={i}
            className="pdf-link-edit"
            style={pdfRectToCss(geom, l.rect, scale)}
            data-tip={l.url}
          >
            {l.objNum !== undefined && (
              <button
                type="button"
                className="pdf-link-remove"
                aria-label={editing.removeLabel}
                data-tip={editing.removeLabel}
                onPointerDown={(e) => e.stopPropagation()}
                onClick={() =>
                  editing.onRemove({
                    pageIndex: pageNo - 1,
                    objNum: l.objNum!,
                    type: 'link',
                    rect: l.rect,
                  })
                }
              >
                ×
              </button>
            )}
          </div>
        ))}
      </div>
    )
  }

  return (
    <div className="pdf-link-layer">
      {shown.map((l, i) => (
        <a
          key={i}
          className="pdf-link"
          style={pdfRectToCss(geom, l.rect, scale)}
          href={l.url ?? '#'}
          data-tip={l.url}
          target={l.url ? '_blank' : undefined}
          rel="noreferrer"
          onClick={(e) => {
            if (!l.url) {
              e.preventDefault()
              onGoToDest(l.dest)
            }
          }}
        />
      ))}
    </div>
  )
}
