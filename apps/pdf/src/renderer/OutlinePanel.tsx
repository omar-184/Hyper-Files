import { useState } from 'react'
import type { ReactElement, ReactNode } from 'react'
import type { OutlinePath } from './outline-edit'

export interface OutlineNode {
  title: string
  bold?: boolean
  italic?: boolean
  dest?: unknown
  url?: string
  /** pdf.js text color, 0-255 per channel */
  color?: ArrayLike<number> | null
  items?: OutlineNode[]
}

/** Max outline nesting rendered: a hostile PDF can nest bookmarks thousands
 *  deep, and unbounded recursion would overflow the render stack. Deeper
 *  levels are dropped (their ancestors still render). */
export const MAX_OUTLINE_DEPTH = 32

/** Bookmark editing: the host applies each change to its pending outline */
export interface OutlineEditing {
  labels: {
    add: string
    rename: string
    remove: string
    up: string
    down: string
    indent: string
    outdent: string
  }
  /** Add a bookmark to the current page after `path` (end when null); resolves to its path */
  onAdd: (after: OutlinePath | null) => Promise<OutlinePath | null>
  onRename: (path: OutlinePath, title: string) => void
  onRemove: (path: OutlinePath) => void
  /** Structural moves resolve to the entry's new path (null = not possible) */
  onMove: (
    path: OutlinePath,
    how: 'up' | 'down' | 'indent' | 'outdent',
  ) => Promise<OutlinePath | null>
}

const key = (path: OutlinePath) => path.join('.')

function Item({
  node,
  path,
  currentDest,
  selectedKey,
  renaming,
  onGo,
  onSelect,
  onStartRename,
  onRename,
}: {
  node: OutlineNode
  path: OutlinePath
  currentDest?: unknown
  selectedKey: string | null
  renaming: string | null
  onGo: (n: OutlineNode) => void
  onSelect?: (path: OutlinePath) => void
  onStartRename?: (path: OutlinePath) => void
  onRename: (path: OutlinePath, title: string | null) => void
}): ReactElement {
  const depth = path.length - 1
  const hasChildren = (node.items?.length ?? 0) > 0
  const isCurrent = currentDest != null && node.dest != null && node.dest === currentDest
  const k = key(path)
  const style = {
    paddingLeft: 10 + depth * 14,
    fontWeight: node.bold ? 600 : 400,
    fontStyle: node.italic ? 'italic' : undefined,
  } as const
  return (
    <>
      {renaming === k ? (
        <input
          className="pdf-outline-rename"
          style={{ marginLeft: 4 + depth * 14 }}
          defaultValue={node.title}
          autoFocus
          onFocus={(e) => e.currentTarget.select()}
          onKeyDown={(e) => {
            if (e.key === 'Enter') onRename(path, e.currentTarget.value)
            else if (e.key === 'Escape') {
              e.stopPropagation()
              onRename(path, null)
            }
          }}
          onBlur={(e) => onRename(path, e.currentTarget.value)}
        />
      ) : (
        <button
          type="button"
          role="treeitem"
          aria-level={depth + 1}
          aria-expanded={hasChildren ? true : undefined}
          aria-current={isCurrent ? true : undefined}
          aria-selected={onSelect ? selectedKey === k : undefined}
          className={`pdf-outline-item${selectedKey === k ? ' is-selected' : ''}`}
          style={style}
          data-tip={node.title}
          onClick={() => {
            onSelect?.(path)
            onGo(node)
          }}
          onDoubleClick={() => onStartRename?.(path)}
        >
          {node.title}
        </button>
      )}
      {depth + 1 < MAX_OUTLINE_DEPTH &&
        node.items?.map((c, i) => (
          <Item
            key={i}
            node={c}
            path={[...path, i]}
            currentDest={currentDest}
            selectedKey={selectedKey}
            renaming={renaming}
            onGo={onGo}
            onSelect={onSelect}
            onStartRename={onStartRename}
            onRename={onRename}
          />
        ))}
    </>
  )
}

function ToolButton({
  label,
  disabled,
  onClick,
  children,
}: {
  label: string
  disabled?: boolean
  onClick: () => void
  children: ReactNode
}): ReactElement {
  return (
    <button
      type="button"
      className="rb-icon"
      aria-label={label}
      data-tip={label}
      disabled={disabled}
      onClick={onClick}
    >
      <svg
        width="16"
        height="16"
        viewBox="0 0 16 16"
        fill="none"
        stroke="currentColor"
        strokeWidth="1.4"
        strokeLinecap="round"
        strokeLinejoin="round"
        aria-hidden="true"
      >
        {children}
      </svg>
    </button>
  )
}

/** Outline (bookmark) tree: click jumps to internal destinations; url entries open external links.
    With `editing`, a toolbar adds, renames, removes and rearranges bookmarks. */
export function OutlinePanel({
  outline,
  note,
  label,
  emptyLabel,
  currentDest,
  onGoToDest,
  editing,
}: {
  outline: OutlineNode[]
  /** Caption above the tree, e.g. when the tree was derived from headings */
  note?: string
  /** Accessible name for the tree; caller passes t('outline') (English fallback kept local) */
  label?: string
  /** Empty-state copy; caller passes t('searchNoResults') (English fallback kept local) */
  emptyLabel?: string
  /** Destination of the current location, when known; matching item gets aria-current */
  currentDest?: unknown
  onGoToDest: (dest: unknown) => void
  editing?: OutlineEditing
}): ReactElement {
  const [selected, setSelected] = useState<OutlinePath | null>(null)
  const [renaming, setRenaming] = useState<string | null>(null)
  const onGo = (n: OutlineNode) => {
    if (n.url) window.open(n.url, '_blank')
    else if (n.dest != null) onGoToDest(n.dest)
  }
  // A selection the tree no longer has (undo, reload) is dropped
  let node: OutlineNode | undefined
  let list = outline
  for (const i of selected ?? []) {
    node = list[i]
    list = node?.items ?? []
    if (!node) break
  }
  const sel = node && selected ? selected : null
  const selKey = sel ? key(sel) : null
  const siblings = sel
    ? sel.length > 1
      ? (sel.slice(0, -1).reduce<OutlineNode[]>((l, i) => l[i]?.items ?? [], outline) ?? [])
      : outline
    : []
  const last = sel ? sel[sel.length - 1]! : 0

  const move = async (how: 'up' | 'down' | 'indent' | 'outdent') => {
    if (!editing || !sel) return
    const next = await editing.onMove(sel, how)
    if (next) setSelected(next)
  }
  const add = async () => {
    if (!editing) return
    const path = await editing.onAdd(sel)
    if (!path) return
    setSelected(path)
    setRenaming(key(path))
  }
  const finishRename = (path: OutlinePath, title: string | null) => {
    setRenaming(null)
    const clean = title?.trim()
    if (editing && clean) editing.onRename(path, clean)
  }

  return (
    <div className="pdf-outline-wrap">
      {editing && (
        <div className="pdf-outline-tools" role="toolbar" aria-label={label}>
          <ToolButton label={editing.labels.add} onClick={() => void add()}>
            <path d="M8 3.5v9M3.5 8h9" />
          </ToolButton>
          <ToolButton
            label={editing.labels.rename}
            disabled={!sel}
            onClick={() => selKey && setRenaming(selKey)}
          >
            <path d="M10.5 3.5l2 2-6.5 6.5H4v-2z" />
          </ToolButton>
          <ToolButton
            label={editing.labels.remove}
            disabled={!sel}
            onClick={() => {
              if (!sel) return
              editing.onRemove(sel)
              setSelected(null)
            }}
          >
            <path d="M3.5 5h9M6.5 5V3.5h3V5M5 5l.6 7.5h4.8L11 5" />
          </ToolButton>
          <span className="pdf-outline-tools-sep" />
          <ToolButton
            label={editing.labels.up}
            disabled={!sel || last === 0}
            onClick={() => void move('up')}
          >
            <path d="M8 12.5v-9M4.5 7L8 3.5 11.5 7" />
          </ToolButton>
          <ToolButton
            label={editing.labels.down}
            disabled={!sel || last >= siblings.length - 1}
            onClick={() => void move('down')}
          >
            <path d="M8 3.5v9M4.5 9L8 12.5 11.5 9" />
          </ToolButton>
          <ToolButton
            label={editing.labels.outdent}
            disabled={!sel || sel.length < 2}
            onClick={() => void move('outdent')}
          >
            <path d="M12.5 8h-9M6.5 5L3.5 8l3 3" />
          </ToolButton>
          <ToolButton
            label={editing.labels.indent}
            disabled={!sel || last === 0}
            onClick={() => void move('indent')}
          >
            <path d="M3.5 8h9M9.5 5l3 3-3 3" />
          </ToolButton>
        </div>
      )}
      <div
        className="pdf-outline"
        role="tree"
        aria-label={label ?? 'Outline'}
        onKeyDown={(e) => {
          if (!editing || !sel || renaming) return
          if (e.key === 'F2') {
            e.preventDefault()
            setRenaming(key(sel))
          } else if (e.key === 'Delete') {
            e.preventDefault()
            editing.onRemove(sel)
            setSelected(null)
          }
        }}
      >
        {note && <div className="pdf-outline-note">{note}</div>}
        {outline.length === 0 ? (
          <div className="pdf-outline-empty" role="status">
            {emptyLabel ?? 'No results'}
          </div>
        ) : (
          outline.map((n, i) => (
            <Item
              key={i}
              node={n}
              path={[i]}
              currentDest={currentDest}
              selectedKey={selKey}
              renaming={renaming}
              onGo={onGo}
              onSelect={editing ? setSelected : undefined}
              onStartRename={editing ? (path) => setRenaming(key(path)) : undefined}
              onRename={finishRename}
            />
          ))
        )}
      </div>
    </div>
  )
}
