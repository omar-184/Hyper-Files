import type { ParseMap } from './parse-map'
import type { Patch } from './patch'

/**
 * The edit vocabulary of the manual UI. Every op compiles to text patches
 * against one base version; the batch is validated as a whole and applied
 * atomically.
 */
export type HtmlOp =
  | { op: 'set_inner_html'; sid: number; html: string }
  | {
      op: 'insert_html'
      sid: number
      position: 'before' | 'after' | 'prepend' | 'append'
      html: string
    }
  | { op: 'remove'; sid: number }
  | { op: 'move'; sid: number; position: 'before' | 'after'; ref_sid: number }
  | { op: 'set_attr'; sid: number; name: string; value: string | null }
  | { op: 'set_style'; sid: number; styles: Record<string, string | null> }
  /** replace one direct child text node (index into the element's text nodes) */
  | { op: 'set_text_node'; sid: number; index: number; text: string }
  /** wrap a character range of one text node in a new inline element */
  | {
      op: 'wrap_text'
      sid: number
      index: number
      start: number
      end: number
      tag: string
      attrs?: Record<string, string>
    }

export type OpErrorKind = 'unknown_op' | 'bad_args' | 'no_such_sid' | 'void_element' | 'overlap'

export interface OpError {
  index: number
  kind: OpErrorKind
  message: string
}

export interface CompiledOps {
  patches: Patch[]
  errors: OpError[]
}

const VOID_TAGS = new Set([
  'area',
  'base',
  'br',
  'col',
  'embed',
  'hr',
  'img',
  'input',
  'link',
  'meta',
  'source',
  'track',
  'wbr',
])

function escapeText(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

function escapeAttr(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/"/g, '&quot;')
}

/** locate `name="…"` / `name='…'` / bare `name` inside a start tag. `from` includes the
 * whitespace before the attribute so removing it leaves no gap. */
function findAttribute(
  startTag: string,
  name: string,
): {
  from: number
  to: number
  valueFrom: number
  valueTo: number
  quote: string
  lead: string
} | null {
  const tagName = /^<[^\s/>]+/.exec(startTag)
  if (!tagName) return null
  const attrs = /(\s+)([^\s"'<>/=]+)(?:\s*=\s*("([^"]*)"|'([^']*)'|([^\s"'=<>`]+)))?/gy
  attrs.lastIndex = tagName[0].length
  let m: RegExpExecArray | null
  while ((m = attrs.exec(startTag))) {
    if (m[2]!.toLowerCase() !== name.toLowerCase()) continue
    const from = m.index
    const to = m.index + m[0].length
    const lead = m[1]!
    if (m[3] === undefined) return { from, to, valueFrom: to, valueTo: to, quote: '', lead }
    const raw = m[3]
    const quote = raw.startsWith('"') ? '"' : raw.startsWith("'") ? "'" : ''
    const valueTo = to - (quote ? 1 : 0)
    const valueFrom = valueTo - (m[4] ?? m[5] ?? m[6] ?? '').length
    return { from, to, valueFrom, valueTo, quote, lead }
  }
  return null
}

const ENTITY_RE = /&(?:#(\d+)|#x([0-9a-fA-F]+)|[a-zA-Z][a-zA-Z0-9]*);/y

/** UTF-16 units the entity decodes to: an astral numeric reference is a surrogate pair */
function entityUnits(entity: RegExpExecArray): number {
  const code = entity[1] ? parseInt(entity[1], 10) : entity[2] ? parseInt(entity[2], 16) : 0
  return code > 0xffff ? 2 : 1
}

/** raw index of the character at `decoded` in a text node, counting entities by their decoded width */
export function decodedToRaw(raw: string, decoded: number): number {
  const target = Math.max(0, Math.floor(Number.isFinite(decoded) ? decoded : 0))
  let i = 0
  let seen = 0
  while (i < raw.length && seen < target) {
    ENTITY_RE.lastIndex = i
    const entity = ENTITY_RE.exec(raw)
    if (entity) {
      i += entity[0].length
      seen += entityUnits(entity)
      continue
    }
    const astral = raw.codePointAt(i)! > 0xffff
    i += astral ? 2 : 1
    seen += astral ? 2 : 1
  }
  return Math.min(i, raw.length)
}

/** the element's source, widened to its whole line when nothing but indentation shares that line */
function lineChunk(
  text: string,
  range: [number, number],
): { from: number; to: number; text: string; fullLine: boolean } {
  const lineStart = text.lastIndexOf('\n', range[0] - 1) + 1
  const onlyIndentBefore = /^[ \t]*$/.test(text.slice(lineStart, range[0]))
  const newlineAfter = text[range[1]] === '\n'
  if (onlyIndentBefore && newlineAfter) {
    return {
      from: lineStart,
      to: range[1] + 1,
      text: text.slice(lineStart, range[1] + 1),
      fullLine: true,
    }
  }
  return { from: range[0], to: range[1], text: text.slice(range[0], range[1]), fullLine: false }
}

function parseStyle(value: string): Array<[string, string]> {
  return value
    .split(';')
    .map((d) => d.trim())
    .filter(Boolean)
    .map((d) => {
      const i = d.indexOf(':')
      return i < 0 ? [d, ''] : [d.slice(0, i).trim(), d.slice(i + 1).trim()]
    }) as Array<[string, string]>
}

function serializeStyle(decls: Array<[string, string]>): string {
  return decls.map(([k, v]) => `${k}: ${v}`).join('; ')
}

function startTagInsertPoint(startTag: string): number {
  // before the closing `>` or `/>`
  const m = /\s*\/?>$/.exec(startTag)
  return m ? m.index : startTag.length
}

/** Compile a batch against one text + map; never applies anything. */
export function compileOps(text: string, map: ParseMap, ops: readonly HtmlOp[]): CompiledOps {
  const patches: Array<Patch & { index: number }> = []
  const errors: OpError[] = []
  const fail = (index: number, kind: OpErrorKind, message: string) =>
    errors.push({ index, kind, message })
  const element = (index: number, sid: unknown) => {
    const e = typeof sid === 'number' ? map.bySid.get(sid) : undefined
    if (!e) fail(index, 'no_such_sid', `sid ${String(sid)} does not exist in the current document.`)
    return e
  }

  ops.forEach((raw, index) => {
    const op = raw as HtmlOp
    switch (op.op) {
      case 'set_inner_html': {
        const e = element(index, op.sid)
        if (!e) return
        if (VOID_TAGS.has(e.tag)) return fail(index, 'void_element', `<${e.tag}> has no content`)
        patches.push({ from: e.inner[0], to: e.inner[1], text: String(op.html ?? ''), index })
        return
      }
      case 'insert_html': {
        const e = element(index, op.sid)
        if (!e) return
        if (typeof op.html !== 'string') return fail(index, 'bad_args', 'insert_html needs html')
        const at =
          op.position === 'before'
            ? e.range[0]
            : op.position === 'after'
              ? e.range[1]
              : op.position === 'prepend'
                ? e.inner[0]
                : op.position === 'append'
                  ? e.inner[1]
                  : -1
        if (at < 0)
          return fail(index, 'bad_args', 'position must be before | after | prepend | append')
        if ((op.position === 'prepend' || op.position === 'append') && VOID_TAGS.has(e.tag)) {
          return fail(index, 'void_element', `<${e.tag}> has no content`)
        }
        patches.push({ from: at, to: at, text: op.html, index })
        return
      }
      case 'remove': {
        const e = element(index, op.sid)
        if (!e) return
        const chunk = lineChunk(text, e.range)
        patches.push({ from: chunk.from, to: chunk.to, text: '', index })
        return
      }
      case 'move': {
        const e = element(index, op.sid)
        const ref = element(index, op.ref_sid)
        if (!e || !ref) return
        if (e.sid === ref.sid || (e.range[0] <= ref.range[0] && e.range[1] >= ref.range[1])) {
          return fail(
            index,
            'bad_args',
            'cannot move an element relative to itself or its descendant',
          )
        }
        if (op.position !== 'before' && op.position !== 'after')
          return fail(index, 'bad_args', 'position must be before | after')
        // an element alone on its line travels with its line so indentation stays tidy
        const chunk = lineChunk(text, e.range)
        patches.push({ from: chunk.from, to: chunk.to, text: '', index })
        if (chunk.fullLine) {
          const refLineStart = text.lastIndexOf('\n', ref.range[0] - 1) + 1
          if (op.position === 'before') {
            const at = /^[ \t]*$/.test(text.slice(refLineStart, ref.range[0]))
              ? refLineStart
              : ref.range[0]
            patches.push({ from: at, to: at, text: chunk.text, index })
          } else if (text[ref.range[1]] === '\n') {
            const at = ref.range[1] + 1
            patches.push({ from: at, to: at, text: chunk.text, index })
          } else {
            const at = ref.range[1]
            patches.push({ from: at, to: at, text: `\n${chunk.text.replace(/\n$/, '')}`, index })
          }
        } else {
          const at = op.position === 'before' ? ref.range[0] : ref.range[1]
          patches.push({ from: at, to: at, text: chunk.text, index })
        }
        return
      }
      case 'set_attr': {
        const e = element(index, op.sid)
        if (!e) return
        if (typeof op.name !== 'string' || !/^[^\s"'>/=]+$/.test(op.name))
          return fail(index, 'bad_args', 'invalid attribute name')
        const tag = text.slice(e.startTag[0], e.startTag[1])
        const found = findAttribute(tag, op.name)
        if (op.value === null) {
          if (found)
            patches.push({
              from: e.startTag[0] + found.from,
              to: e.startTag[0] + found.to,
              text: '',
              index,
            })
          return
        }
        const value = String(op.value)
        if (found && found.quote) {
          patches.push({
            from: e.startTag[0] + found.valueFrom,
            to: e.startTag[0] + found.valueTo,
            text:
              found.quote === '"'
                ? escapeAttr(value)
                : value.replace(/&/g, '&amp;').replace(/'/g, '&#39;'),
            index,
          })
        } else if (found) {
          patches.push({
            from: e.startTag[0] + found.from,
            to: e.startTag[0] + found.to,
            text: `${found.lead}${op.name}="${escapeAttr(value)}"`,
            index,
          })
        } else {
          const at = e.startTag[0] + startTagInsertPoint(tag)
          patches.push({ from: at, to: at, text: ` ${op.name}="${escapeAttr(value)}"`, index })
        }
        return
      }
      case 'set_style': {
        const e = element(index, op.sid)
        if (!e) return
        if (!op.styles || typeof op.styles !== 'object')
          return fail(index, 'bad_args', 'set_style needs a styles object')
        const tag = text.slice(e.startTag[0], e.startTag[1])
        const found = findAttribute(tag, 'style')
        const decls = found ? parseStyle(tag.slice(found.valueFrom, found.valueTo)) : []
        for (const [prop, val] of Object.entries(op.styles)) {
          const i = decls.findIndex(([k]) => k.toLowerCase() === prop.toLowerCase())
          if (val === null) {
            if (i >= 0) decls.splice(i, 1)
          } else if (i >= 0) decls[i] = [decls[i]![0], String(val)]
          else decls.push([prop, String(val)])
        }
        const serialized = serializeStyle(decls)
        if (found && found.quote) {
          if (!serialized)
            patches.push({
              from: e.startTag[0] + found.from,
              to: e.startTag[0] + found.to,
              text: '',
              index,
            })
          else
            patches.push({
              from: e.startTag[0] + found.valueFrom,
              to: e.startTag[0] + found.valueTo,
              // Mirror set_attr above: a single-quoted style attribute
              // must escape apostrophes, not double quotes, or the value
              // breaks out of the attribute.
              text:
                found.quote === '"'
                  ? escapeAttr(serialized)
                  : serialized.replace(/&/g, '&amp;').replace(/'/g, '&#39;'),
              index,
            })
        } else if (serialized) {
          const at = found ? e.startTag[0] + found.from : e.startTag[0] + startTagInsertPoint(tag)
          const to = found ? e.startTag[0] + found.to : at
          patches.push({
            from: at,
            to,
            text: `${found ? found.lead : ' '}style="${escapeAttr(serialized)}"`,
            index,
          })
        }
        return
      }
      case 'set_text_node': {
        const e = element(index, op.sid)
        if (!e) return
        const node = Number.isInteger(op.index) ? e.textNodes[op.index] : undefined
        if (!node) return fail(index, 'bad_args', `text node ${String(op.index)} does not exist`)
        patches.push({ from: node[0], to: node[1], text: escapeText(String(op.text ?? '')), index })
        return
      }
      case 'wrap_text': {
        const e = element(index, op.sid)
        if (!e) return
        const node = Number.isInteger(op.index) ? e.textNodes[op.index] : undefined
        if (!node) return fail(index, 'bad_args', `text node ${String(op.index)} does not exist`)
        if (typeof op.tag !== 'string' || !/^[a-zA-Z][a-zA-Z0-9-]*$/.test(op.tag))
          return fail(index, 'bad_args', 'invalid tag name')
        const raw = text.slice(node[0], node[1])
        // start/end come from the live DOM (decoded text); entities in the source widen the offsets
        const start = decodedToRaw(raw, Number(op.start))
        const end = Math.max(start, decodedToRaw(raw, Number(op.end)))
        if (end === start) return fail(index, 'bad_args', 'empty range')
        const attrs = Object.entries(op.attrs ?? {})
          .filter(([k]) => /^[^\s"'>/=]+$/.test(k))
          .map(([k, v]) => ` ${k}="${escapeAttr(String(v))}"`)
          .join('')
        patches.push({
          from: node[0] + start,
          to: node[0] + start,
          text: `<${op.tag}${attrs}>`,
          index,
        })
        patches.push({ from: node[0] + end, to: node[0] + end, text: `</${op.tag}>`, index })
        return
      }
      default:
        fail(index, 'unknown_op', `unknown op "${String((raw as { op?: unknown }).op)}"`)
    }
  })

  // overlap check across the whole batch (adjacent insertions at the same offset are allowed)
  const sorted = [...patches].sort((a, b) => a.from - b.from || a.to - b.to)
  for (let i = 1; i < sorted.length; i++) {
    const prev = sorted[i - 1]!
    const cur = sorted[i]!
    if (cur.from < prev.to && !(cur.from === cur.to && cur.from === prev.to)) {
      errors.push({
        index: cur.index,
        kind: 'overlap',
        message: `op ${cur.index + 1} overlaps op ${prev.index + 1}; edit the outer element once instead.`,
      })
      break
    }
  }
  return { patches: patches.map(({ from, to, text: t }) => ({ from, to, text: t })), errors }
}
