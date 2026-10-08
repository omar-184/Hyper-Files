/**
 * Op documentation, parsed from `docs/*.md`: the reference for every edit op
 * the executor registers. A test asserts the docs cover the registry exactly,
 * and the executor's guided errors append the failing op's one-line
 * signature via `opUsage(name)`.
 *
 * Markdown block format (see docs/_format.md):
 *
 *   ### opName [(internal[, pending])]
 *   `signature`
 *   ...free prose, tables, ```json examples...
 *
 * The files are inlined at build time (`?raw`); nothing is read from disk at
 * runtime, in the main process or the renderer.
 */
import textMd from './docs/text.md?raw'
import elementMd from './docs/element.md?raw'
import insertMd from './docs/insert.md?raw'
import tableMd from './docs/table.md?raw'
import slideMd from './docs/slide.md?raw'
import deckMd from './docs/deck.md?raw'

export type OpGroup = 'text' | 'element' | 'insert' | 'table' | 'slide' | 'deck'

export interface OpDoc {
  /** Compact usage signature: the fields beside target (? marks optional). */
  sig: string
  /** Doc group (the file it lives in). */
  group: OpGroup
  /**
   * true: the payload is bytes/clipboard/part-path data that only the app's
   * own UI produces (file pickers, paste, media parts).
   */
  internal?: true
  /**
   * Registered by an in-flight branch: hidden from usage lines until it
   * actually lands (a usage line on the unknown-op error would read as a
   * field problem). Drop the flag when the branch merges.
   */
  pending?: true
  /** ```json examples inside the block, verbatim (tests dry-run each one). */
  examples: string[]
  /** The whole block below the signature line (prose, tables, examples). */
  body: string
}

const GROUP_ORDER: OpGroup[] = ['text', 'element', 'insert', 'table', 'slide', 'deck']

const SOURCES: Record<OpGroup, string> = {
  text: textMd,
  element: elementMd,
  insert: insertMd,
  table: tableMd,
  slide: slideMd,
  deck: deckMd,
}

const HEADING_RE = /^### ([A-Za-z][A-Za-z0-9]*)(?:\s+\(([^)]*)\))?\s*$/
const SIG_RE = /^`([^`]+)`\s*$/
const FENCE_JSON_RE = /^```json\s*$/
const FENCE_END_RE = /^```\s*$/

function parseGroup(group: OpGroup, md: string): Record<string, OpDoc> {
  const lines = md.split(/\r?\n/)
  const docs: Record<string, OpDoc> = {}
  let i = 0
  while (i < lines.length) {
    const heading = HEADING_RE.exec(lines[i]!)
    if (!heading) {
      i++
      continue
    }
    const name = heading[1]!
    const flags = (heading[2] ?? '')
      .split(',')
      .map((f) => f.trim())
      .filter(Boolean)
    for (const flag of flags) {
      if (flag !== 'internal' && flag !== 'pending') {
        throw new Error(`op docs (${group}): unknown flag "${flag}" on ### ${name}`)
      }
    }
    if (docs[name]) throw new Error(`op docs (${group}): duplicate block ### ${name}`)
    // Signature: first non-empty line after the heading, in single backticks
    i++
    while (i < lines.length && lines[i]!.trim() === '') i++
    const sigMatch = i < lines.length ? SIG_RE.exec(lines[i]!) : null
    if (!sigMatch) {
      throw new Error(
        `op docs (${group}): ### ${name} must be followed by its signature in backticks`,
      )
    }
    const sig = sigMatch[1]!
    i++
    // Body: until the next heading
    const bodyLines: string[] = []
    const examples: string[] = []
    let fence: string[] | null = null
    while (i < lines.length && !HEADING_RE.test(lines[i]!)) {
      const line = lines[i]!
      if (fence) {
        if (FENCE_END_RE.test(line)) {
          examples.push(fence.join('\n'))
          fence = null
        } else fence.push(line)
      } else if (FENCE_JSON_RE.test(line)) {
        fence = []
      }
      bodyLines.push(line)
      i++
    }
    if (fence) throw new Error(`op docs (${group}): unterminated json fence in ### ${name}`)
    docs[name] = {
      sig,
      group,
      ...(flags.includes('internal') ? { internal: true as const } : {}),
      ...(flags.includes('pending') ? { pending: true as const } : {}),
      examples,
      body: bodyLines.join('\n').trim(),
    }
  }
  return docs
}

export const OP_DOCS: Record<string, OpDoc> = Object.assign(
  {},
  ...GROUP_ORDER.map((g) => parseGroup(g, SOURCES[g])),
)

/** One-line usage for a failing op, appended to its guided error. */
export function opUsage(name: string): string | undefined {
  const doc = OP_DOCS[name]
  if (!doc || doc.pending) return undefined
  return `Usage: ${name} ${doc.sig}`
}

export const OP_GROUPS: readonly OpGroup[] = GROUP_ORDER
