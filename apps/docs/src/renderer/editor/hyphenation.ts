import { Extension } from '@tiptap/core'
import { Plugin, type EditorState } from '@tiptap/pm/state'
import { Decoration, DecorationSet, type EditorView } from '@tiptap/pm/view'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { rangeSlot } from '../dom-range'
import { SettledParagraphCache, noteFloatTransaction } from './settled-measure'
import { PHASED_CONTENT_SETTLED_EVENT, isPhasedContentPending } from '../phased-content'
import { DOC_CSS_COMMITTED_EVENT } from './cjk-punct-shrink'
import { hyphenBreaks, hyphenationPluginKey } from './hyphenate'
import { justifyShrinkPluginKey, sameLine } from './justify-shrink'

/**
 * Word's automatic hyphenation (settings.xml w:autoHyphenation, Layout >
 * Hyphenation > Automatic). The document stylesheet already marks those
 * paragraphs hyphens:auto (w:suppressAutoHyphens styles opt out), but Electron
 * ships no hyphenation dictionaries, so Chromium never splits a word on its
 * own. This extension supplies the break points as display-only soft hyphen
 * (U+00AD) widgets, and Chromium's line breaker then takes the last one that
 * fits like any typed soft hyphen. Nothing enters the document: saves, copy
 * and find see the plain words.
 *
 * Only words near a line boundary get widgets (a long document would
 * otherwise carry one per syllable): a line's first two hyphenatable words
 * and any word already split across lines. After a pass re-wraps the lines,
 * the next pass covers the words that moved onto a boundary.
 *
 * Hyphenation zone: on ragged (left-aligned) lines Word leaves a word whole
 * when the gap it would leave at the line end is narrower than the zone, so
 * such a word gets no break points. Justified lines are stretched to the
 * margin and hyphenate whenever a fragment fits.
 */

export interface HyphenationStorage {
  /** settings.xml w:autoHyphenation (or the Layout > Hyphenation toggle) */
  enabled: boolean
  /** w:hyphenationZone in layout px (Word's default 0.25in) */
  zonePx: number
  /** w:doNotHyphenateCaps: words in all capitals never break */
  noCaps: boolean
}

declare module '@tiptap/core' {
  interface Storage {
    hyphenation: HyphenationStorage
  }
}

export { hyphenationPluginKey }

/** Word's default hyphenation zone: 0.25in (360 twips) */
export const DEFAULT_HYPHENATION_ZONE_PX = 24

/** w:hyphenationZone twips → layout px */
export function hyphenationZonePx(twips: number | undefined): number {
  return twips === undefined ? DEFAULT_HYPHENATION_ZONE_PX : twips / 15
}

const MEASURE_RETRY_MAX = 10
const MEASURE_SIGS_MAX = 12
/** a paragraph re-wraps every time a boundary word splits; a few passes cover the lines that moved */
const PARA_ROUNDS = 6
/** hyphenatable words per line that get break points (the line's first ones) */
const WORDS_PER_LINE = 2

/** a candidate word inside a paragraph's flattened text: the letters core
 *  of a whitespace-delimited run, with punctuation either side allowed */
export interface HyphenWord {
  /** offset of the run in the paragraph text */
  runStart: number
  /** offset of the letters core in the paragraph text */
  start: number
  word: string
  breaks: number[]
}

const CORE_RE = /^([^A-Za-z￼­-]*)([A-Za-z]+)([^A-Za-z￼­-]*)$/

/** hyphenatable words of a paragraph's text (inline nodes as U+FFFC, hard
 *  breaks as \n) in order */
export function hyphenWords(text: string, noCaps: boolean): HyphenWord[] {
  const out: HyphenWord[] = []
  const re = /\S+/g
  let m: RegExpExecArray | null
  while ((m = re.exec(text))) {
    const core = CORE_RE.exec(m[0])
    if (!core) continue
    const word = core[2]
    if (noCaps && word === word.toUpperCase()) continue
    const breaks = hyphenBreaks(word)
    if (breaks.length === 0) continue
    out.push({ runStart: m.index, start: m.index + core[1].length, word, breaks })
  }
  return out
}

/** the paragraph's inline content as one string, one char per position */
function paragraphText(node: ProseMirrorNode): string {
  const parts: string[] = []
  node.forEach((child) => {
    if (child.isText) parts.push(child.text ?? '')
    else parts.push((child.type.name === 'hardBreak' ? '\n' : '￼').repeat(child.nodeSize))
  })
  return parts.join('')
}

const probe = rangeSlot()

class HyphenationView {
  private lastSig = ''
  private seenSigs = new Set<string>()
  private frozen = false
  private retryRaf = 0
  private retries = 0
  private resizeObserver?: ResizeObserver
  private lastDomWidth = -1
  private results = new SettledParagraphCache<number[]>(
    (r, d) => r.map((p) => p + d),
    (r) => r.length === 0,
    PARA_ROUNDS,
  )
  private relayout = () => {
    this.invalidate()
    this.measure()
  }

  constructor(
    private view: EditorView,
    private storage: HyphenationStorage,
  ) {
    this.measure()
    document.fonts?.addEventListener('loadingdone', this.relayout)
    // the hyphens:auto rules arrive with the doc stylesheet
    document.addEventListener(DOC_CSS_COMMITTED_EVENT, this.relayout)
    document.addEventListener(PHASED_CONTENT_SETTLED_EVENT, this.relayout)
    if (typeof ResizeObserver !== 'undefined') {
      this.resizeObserver = new ResizeObserver(() => {
        const w = this.view.dom.offsetWidth
        if (w === this.lastDomWidth) return
        this.lastDomWidth = w
        this.relayout()
      })
      this.resizeObserver.observe(view.dom)
    }
  }

  private invalidate() {
    this.results.clear()
    this.restartConvergence()
  }

  private restartConvergence() {
    this.seenSigs.clear()
    this.frozen = false
    this.lastSig = ''
  }

  update(view: EditorView, prevState: EditorState) {
    if (view.state.doc !== prevState.doc) {
      this.restartConvergence()
    } else if (
      hyphenationPluginKey.getState(view.state) === hyphenationPluginKey.getState(prevState) &&
      // justify-shrink moves wrap points too
      justifyShrinkPluginKey.getState(view.state) === justifyShrinkPluginKey.getState(prevState)
    ) {
      return
    }
    this.measure()
  }

  destroy() {
    document.fonts?.removeEventListener('loadingdone', this.relayout)
    document.removeEventListener(DOC_CSS_COMMITTED_EVENT, this.relayout)
    document.removeEventListener(PHASED_CONTENT_SETTLED_EVENT, this.relayout)
    this.resizeObserver?.disconnect()
    if (this.retryRaf) cancelAnimationFrame(this.retryRaf)
  }

  /** settings changed (Layout > Hyphenation, document load) */
  refresh() {
    this.invalidate()
    this.measure()
  }

  private scheduleRetry() {
    if (this.retryRaf || this.retries >= MEASURE_RETRY_MAX) return
    this.retries++
    this.retryRaf = requestAnimationFrame(() => {
      this.retryRaf = 0
      this.measure()
    })
  }

  private measure() {
    if (this.retryRaf) {
      cancelAnimationFrame(this.retryRaf)
      this.retryRaf = 0
    }
    const { view } = this
    if (isPhasedContentPending()) return
    // PDF export parks the editor subtree; layout reads would re-lay it out
    if (view.dom.closest('.app.pv-exporting')) {
      this.retries = 0
      this.scheduleRetry()
      return
    }
    const old = hyphenationPluginKey.getState(view.state)
    if (!this.storage.enabled) {
      this.lastSig = ''
      if (old && old !== DecorationSet.empty)
        view.dispatch(
          view.state.tr.setMeta(hyphenationPluginKey, []).setMeta('addToHistory', false),
        )
      return
    }
    if (!view.dom.isConnected) {
      this.scheduleRetry()
      return
    }

    const paras: Array<{ node: ProseMirrorNode; pos: number }> = []
    view.state.doc.descendants((node, pos) => {
      if (!node.isTextblock) return true
      if (/[A-Za-z]{5}/.test(node.textContent)) paras.push({ node, pos })
      return false
    })

    const points: number[] = []
    let measurable = paras.length === 0
    this.results.beginPass(view)
    const topLevel = SettledParagraphCache.topLevelDom(view)
    const shrinks = justifyShrinkPluginKey.getState(view.state)
    for (const para of paras) {
      const end = para.pos + para.node.nodeSize
      const measured = this.results.measure(
        view,
        para.node,
        para.pos,
        (el) => this.measureParagraph(para.node, para.pos, el),
        topLevel.get(para.node),
        // justify-shrink's compressed lines wrap elsewhere
        shrinks
          ?.find(para.pos, end)
          .map((d) => `${d.from - para.pos}-${d.to - para.pos}:${shrinkStyle(d)}`)
          .join() ?? '',
      )
      if (!measured) continue
      measurable = true
      points.push(...measured)
    }
    if (!measurable) {
      this.scheduleRetry()
      return
    }
    this.retries = 0

    const sig = points.join(',')
    if (sig === this.lastSig) return
    if (this.frozen) return
    if (this.seenSigs.has(sig) || this.seenSigs.size >= MEASURE_SIGS_MAX) {
      this.frozen = true
      console.warn('[docs] hyphenation layout did not converge; keeping current break points')
      return
    }
    this.seenSigs.add(sig)
    this.lastSig = sig

    if (points.length === 0 && (!old || old === DecorationSet.empty)) return
    const { doc } = view.state
    const decos = points.map((p) =>
      Decoration.widget(p, softHyphen, {
        key: 'shy',
        marks: doc.resolve(p).marks(),
        ignoreSelection: true,
      }),
    )
    view.dispatch(view.state.tr.setMeta(hyphenationPluginKey, decos).setMeta('addToHistory', false))
  }

  /** soft hyphen positions for one paragraph; null = not measurable now */
  private measureParagraph(node: ProseMirrorNode, pos: number, el: HTMLElement): number[] | null {
    const { view } = this
    if (el.offsetWidth === 0) return null
    const rect = el.getBoundingClientRect()
    if (rect.width === 0) return null
    const cs = window.getComputedStyle(el)
    if (cs.hyphens !== 'auto' || cs.direction === 'rtl') return []
    const text = paragraphText(node)
    const words = hyphenWords(text, this.storage.noCaps)
    if (words.length === 0) return []
    // rects are screen px (page zoom transform); the zone is layout px
    const zoom = rect.width / el.offsetWidth
    const ragged = cs.textAlign === 'left' || cs.textAlign === 'start'
    const contentRight =
      rect.right -
      ((parseFloat(cs.paddingRight) || 0) + (parseFloat(cs.borderRightWidth) || 0)) * zoom
    const zone = this.storage.zonePx * zoom

    const base = pos + 1
    const rects = (from: number, to: number): DOMRect[] => {
      const a = view.domAtPos(base + from)
      const b = view.domAtPos(base + to)
      const range = probe()
      try {
        range.setStart(a.node, a.offset)
        range.setEnd(b.node, b.offset)
      } catch {
        return []
      }
      return Array.from(range.getClientRects()).filter((r) => r.width > 0.01)
    }
    /** the last visible box before a run: the previous run's last char */
    const prevBox = (runStart: number): DOMRect | null | 'para-start' | 'hard-break' => {
      let j = runStart - 1
      while (j >= 0 && /\s/.test(text[j]) && text[j] !== '\n') j--
      if (j < 0) return 'para-start'
      if (text[j] === '\n') return 'hard-break'
      if (text[j] === '￼') {
        const dom = view.nodeDOM(base + j)
        return dom instanceof HTMLElement ? dom.getBoundingClientRect() : null
      }
      const r = rects(j, j + 1)
      return r.length > 0 ? r[r.length - 1] : null
    }

    const out: number[] = []
    /** hyphenatable words so far on the current line */
    let onLine = 0
    /** the current line follows a wrap (not the paragraph's or a hard break's first line) */
    let wrapped = false
    let lineBox: DOMRect | null = null
    for (const w of words) {
      const boxes = rects(w.start, w.start + w.word.length)
      if (boxes.length === 0) continue
      const first = boxes[0]
      const split = boxes.some((r) => !sameLine(r, first))
      const newLine = lineBox === null || !sameLine(first, lineBox)
      lineBox = boxes[boxes.length - 1]
      // still on the line of the previous candidate: not a boundary word, but
      // one of a line's first words takes the boundary when the lines above
      // pull text back
      let boundary = false
      let prev: DOMRect | null = null
      if (newLine || split) {
        const p = prevBox(w.runStart)
        const placed = p !== 'para-start' && p !== 'hard-break' && p !== null
        if (newLine) {
          onLine = 0
          wrapped = placed
        }
        if (placed) {
          prev = p
          boundary = split || !sameLine(first, p)
        }
      }
      onLine++
      if (boundary && prev) {
        // Word keeps the word whole when the line would end short of the zone anyway
        if (!ragged || contentRight - prev.right >= zone) push(out, base + w.start, w)
      } else if (wrapped && onLine <= WORDS_PER_LINE) {
        push(out, base + w.start, w)
      }
      if (split) {
        // the tail opens the next line
        onLine = 1
        wrapped = true
      }
    }
    return out
  }
}

function shrinkStyle(d: Decoration): string {
  return (d as unknown as { type?: { attrs?: { style?: string } } }).type?.attrs?.style ?? ''
}

function push(out: number[], at: number, w: HyphenWord): void {
  for (const b of w.breaks) out.push(at + b)
}

function softHyphen(): HTMLElement {
  const span = document.createElement('span')
  span.className = 'doc-shy'
  span.textContent = '­'
  return span
}

export const HyphenationExtension = Extension.create({
  name: 'hyphenation',
  addStorage(): HyphenationStorage {
    return { enabled: false, zonePx: DEFAULT_HYPHENATION_ZONE_PX, noCaps: false }
  },
  addProseMirrorPlugins() {
    const storage = this.storage as HyphenationStorage
    return [
      new Plugin({
        key: hyphenationPluginKey,
        state: {
          init: () => DecorationSet.empty,
          apply(tr, old) {
            noteFloatTransaction(tr)
            const meta = tr.getMeta(hyphenationPluginKey) as Decoration[] | undefined
            if (meta)
              // create() consumes entries of the array it is given; App's
              // pagination listener reads the meta array afterwards
              return meta.length > 0 ? DecorationSet.create(tr.doc, [...meta]) : DecorationSet.empty
            return old.map(tr.mapping, tr.doc)
          },
        },
        props: {
          decorations(state) {
            return this.getState(state)
          },
        },
        view: (editorView) => {
          const v = new HyphenationView(editorView, storage)
          hyphenationViews.set(editorView, v)
          return v
        },
      }),
    ]
  },
})

const hyphenationViews = new WeakMap<EditorView, HyphenationView>()

/** re-run hyphenation after its settings changed */
export function refreshHyphenation(view: EditorView): void {
  hyphenationViews.get(view)?.refresh()
}
