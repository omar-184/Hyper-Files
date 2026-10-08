import { describe, expect, it } from 'vitest'
import { buildParseMap } from '../src/renderer/document/parse-map'
import { compileOps, type HtmlOp } from '../src/renderer/document/ops'
import { applyPatches } from '../src/renderer/document/patch'

const DOC = `<!doctype html>
<html>
<head><title>T</title><style>.card { color: red; }</style></head>
<body>
  <section class="hero">
    <h1 id="title">Hello &amp; welcome</h1>
    <p class="lead">First paragraph.</p>
  </section>
  <ul>
    <li>one</li>
    <li>two</li>
    <li>three</li>
  </ul>
  <img src="a.png">
</body>
</html>
`

function run(ops: HtmlOp[], text = DOC) {
  const map = buildParseMap(text, 1)
  const compiled = compileOps(text, map, ops)
  return {
    compiled,
    next: compiled.errors.length ? text : applyPatches(text, compiled.patches),
    map,
  }
}

function sidOf(text: string, tag: string, nth = 0): number {
  const map = buildParseMap(text, 1)
  return map.elements.filter((e) => e.tag === tag)[nth]!.sid
}

describe('compileOps', () => {
  it('set_inner_html / insert_html / remove keep the rest byte-identical', () => {
    const p = sidOf(DOC, 'p')
    const ul = sidOf(DOC, 'ul')
    const img = sidOf(DOC, 'img')
    const { next, compiled } = run([
      { op: 'set_inner_html', sid: p, html: 'Rewritten.' },
      { op: 'insert_html', sid: ul, position: 'append', html: '<li>four</li>' },
      { op: 'remove', sid: img },
    ])
    expect(compiled.errors).toEqual([])
    expect(next).toContain('<p class="lead">Rewritten.</p>')
    expect(next).toContain('<li>three</li>\n  <li>four</li></ul>')
    expect(next).not.toContain('<img')
    expect(next).toContain('</ul>\n</body>')
  })
  it('set_attr and set_style edit the start tag in place, preserving other attributes', () => {
    const h1 = sidOf(DOC, 'h1')
    const { next } = run([
      { op: 'set_attr', sid: h1, name: 'id', value: 'main-title' },
      { op: 'set_attr', sid: h1, name: 'data-x', value: 'y "q"' },
    ])
    expect(next).toContain('<h1 id="main-title" data-x="y &quot;q&quot;">')
    const sec = sidOf(next, 'section')
    const styled = run(
      [{ op: 'set_style', sid: sec, styles: { color: 'blue', padding: '4px' } }],
      next,
    ).next
    expect(styled).toContain('<section class="hero" style="color: blue; padding: 4px">')
    const sec2 = sidOf(styled, 'section')
    const unstyled = run(
      [{ op: 'set_style', sid: sec2, styles: { color: null, padding: null } }],
      styled,
    ).next
    expect(unstyled).toContain('<section class="hero">')
    const removed = run(
      [{ op: 'set_attr', sid: sidOf(unstyled, 'section'), name: 'class', value: null }],
      unstyled,
    ).next
    expect(removed).toContain('<section>')
  })
  it('set_attr matches the complete attribute name rather than a prefix', () => {
    const doc = '<a hreflang=en href=/old>x</a>'
    const { next, compiled } = run(
      [{ op: 'set_attr', sid: sidOf(doc, 'a'), name: 'href', value: '/new' }],
      doc,
    )
    expect(compiled.errors).toEqual([])
    expect(next).toBe('<a hreflang=en href="/new">x</a>')
  })
  it('set_attr ignores attribute-like text inside another quoted value', () => {
    const doc = '<a title="keep href=/wrong" HREF="/old">x</a>'
    const { next, compiled } = run(
      [{ op: 'set_attr', sid: sidOf(doc, 'a'), name: 'href', value: '/new' }],
      doc,
    )
    expect(compiled.errors).toEqual([])
    expect(next).toBe('<a title="keep href=/wrong" HREF="/new">x</a>')
  })
  it('set_style ignores style-like text inside another quoted value', () => {
    const doc = '<div title="keep style=color:red" style="color: red">x</div>'
    const { next, compiled } = run(
      [{ op: 'set_style', sid: sidOf(doc, 'div'), styles: { color: 'blue' } }],
      doc,
    )
    expect(compiled.errors).toEqual([])
    expect(next).toBe('<div title="keep style=color:red" style="color: blue">x</div>')
  })
  it('set_attr escapes ampersands in single-quoted attributes', () => {
    const doc = `<html><body><div title='x'>hi</div></body></html>`
    const sid = sidOf(doc, 'div')
    const { next } = run([{ op: 'set_attr', sid, name: 'title', value: 'a&b' }], doc)
    expect(next).toContain(`title='a&amp;b'`)
    const { next: next2 } = run(
      [{ op: 'set_attr', sid: sidOf(next, 'div'), name: 'title', value: '&lt;' }],
      next,
    )
    expect(next2).toContain(`title='&amp;lt;'`)
  })
  it('set_style escapes apostrophes in single-quoted style attributes', () => {
    const doc = `<html><body><div style='color: red'>hi</div></body></html>`
    const sid = sidOf(doc, 'div')
    const { next } = run(
      [{ op: 'set_style', sid, styles: { 'font-family': "a'b", color: 'a&b' } }],
      doc,
    )
    expect(next).toContain(`style='color: a&amp;b; font-family: a&#39;b'`)
  })
  it('move relocates an element and rejects moving into itself', () => {
    const one = sidOf(DOC, 'li', 0)
    const three = sidOf(DOC, 'li', 2)
    const { next, compiled } = run([{ op: 'move', sid: one, position: 'after', ref_sid: three }])
    expect(compiled.errors).toEqual([])
    expect(next.indexOf('<li>one</li>')).toBeGreaterThan(next.indexOf('<li>three</li>'))
    // whole lines travel: indentation and line breaks stay tidy
    expect(next).toContain('<ul>\n    <li>two</li>\n    <li>three</li>\n    <li>one</li>\n  </ul>')
    const two = sidOf(DOC, 'li', 1)
    const up = run([{ op: 'move', sid: two, position: 'before', ref_sid: one }]).next
    expect(up).toContain('<ul>\n    <li>two</li>\n    <li>one</li>\n    <li>three</li>\n  </ul>')
    const ul = sidOf(DOC, 'ul')
    expect(
      run([{ op: 'move', sid: ul, position: 'after', ref_sid: one }]).compiled.errors[0]?.kind,
    ).toBe('bad_args')
  })
  it('rejects void-element content ops, unknown sids and overlapping batches as a whole', () => {
    const img = sidOf(DOC, 'img')
    const p = sidOf(DOC, 'p')
    const sec = sidOf(DOC, 'section')
    const { compiled } = run([
      { op: 'set_inner_html', sid: img, html: 'x' },
      { op: 'set_inner_html', sid: 9999, html: 'x' },
      { op: 'set_inner_html', sid: sec, html: '' },
      { op: 'set_inner_html', sid: p, html: 'inner' },
    ])
    expect(compiled.errors.map((e) => e.kind)).toEqual(['void_element', 'no_such_sid', 'overlap'])
  })
})

describe('structural ops', () => {
  const T = '<div>\n  <p class="x">Hello <b>big</b> world</p>\n  <img src="a.png">\n</div>'
  it('set_text_node and wrap_text address direct text nodes', () => {
    const p = sidOf(T, 'p')
    const edited = run([{ op: 'set_text_node', sid: p, index: 1, text: ' & universe' }], T).next
    expect(edited).toContain('<b>big</b> &amp; universe</p>')
    const wrapped = run(
      [{ op: 'wrap_text', sid: p, index: 0, start: 0, end: 5, tag: 'em' }],
      T,
    ).next
    expect(wrapped).toContain('<p class="x"><em>Hello</em> <b>big</b> world</p>')
    const linked = run(
      [
        {
          op: 'wrap_text',
          sid: p,
          index: 1,
          start: 1,
          end: 6,
          tag: 'a',
          attrs: { href: 'https://x' },
        },
      ],
      T,
    ).next
    expect(linked).toContain('<b>big</b> <a href="https://x">world</a></p>')
    expect(
      run([{ op: 'wrap_text', sid: p, index: 5, start: 0, end: 1, tag: 'em' }], T).compiled
        .errors[0]?.kind,
    ).toBe('bad_args')
  })
  it('counts an astral numeric character reference as two UTF-16 units', async () => {
    const { decodedToRaw } = await import('../src/renderer/document/ops')
    const decodedA = '\u{1F600}abc'.indexOf('a')
    expect(decodedToRaw('\u{1F600}abc', decodedA)).toBe(2)
    expect(decodedToRaw('&#x1F600;abc', decodedA)).toBe(9)
    expect(decodedToRaw('&#128512;abc', decodedA)).toBe(9)
    expect(decodedToRaw('&#x41;bc', 1)).toBe(6)
    const E = '<p>&#x1F600;abc</p>'
    const p = sidOf(E, 'p')
    expect(
      run([{ op: 'wrap_text', sid: p, index: 0, start: 2, end: 5, tag: 'strong' }], E).next,
    ).toBe('<p>&#x1F600;<strong>abc</strong></p>')
  })
  it('wrap_text maps decoded (DOM) offsets across entities', async () => {
    const { decodedToRaw } = await import('../src/renderer/document/ops')
    const raw = 'A &amp; B&nbsp;C'
    expect(decodedToRaw(raw, 0)).toBe(0)
    expect(decodedToRaw(raw, 2)).toBe(2) // before &amp;
    expect(decodedToRaw(raw, 3)).toBe(7) // after &amp;
    expect(decodedToRaw(raw, 4)).toBe(8) // "B"
    expect(decodedToRaw(raw, 5)).toBe(9) // before &nbsp;
    expect(decodedToRaw(raw, 6)).toBe(15) // "C"
    expect(decodedToRaw(raw, 99)).toBe(raw.length)
    const E = '<p>A &amp; B&nbsp;C</p>'
    const p = sidOf(E, 'p')
    // DOM selection "B" = decoded offsets [4, 5)
    expect(
      run([{ op: 'wrap_text', sid: p, index: 0, start: 4, end: 5, tag: 'strong' }], E).next,
    ).toBe('<p>A &amp; <strong>B</strong>&nbsp;C</p>')
  })
})

describe('preview instrumentation', () => {
  it('adds data-gx-sid to every mapped start tag and appends the inspector before </body>', async () => {
    const { instrumentForPreview } = await import('../src/renderer/preview/instrument')
    const src = '<html><body><p class="a">x</p><img src="i.png"/><br></body></html>'
    const map = buildParseMap(src, 1)
    const out = instrumentForPreview(src, map, 'console.log(1)')
    expect(instrumentForPreview(src, map, "Number('__GX_VERSION__')")).toContain("Number('1')")
    const p = map.elements.find((e) => e.tag === 'p')!
    const img = map.elements.find((e) => e.tag === 'img')!
    const br = map.elements.find((e) => e.tag === 'br')!
    expect(out).toContain(`<p class="a" data-gx-sid="${p.sid}">`)
    expect(out).toContain(`<img src="i.png" data-gx-sid="${img.sid}"/>`)
    expect(out).toContain(`<br data-gx-sid="${br.sid}">`)
    expect(out).toMatch(/<script data-gx-inspector>console\.log\(1\)<\/script><\/body>/)
    // a fragment without </body> still gets the script
    expect(instrumentForPreview('<p>x</p>', buildParseMap('<p>x</p>', 1), 's')).toMatch(
      /<\/p><script data-gx-inspector>s<\/script>$/,
    )
  })

  it('keeps authored data-sid separate from instrumentation', async () => {
    const { instrumentForPreview } = await import('../src/renderer/preview/instrument')
    const src = '<p data-sid="999">x</p>'
    const map = buildParseMap(src, 1)
    const out = instrumentForPreview(src, map, 'inspector')
    const sid = map.elements.find((e) => e.tag === 'p')!.sid
    expect(out).toContain(`data-sid="999" data-gx-sid="${sid}"`)
    expect(out).not.toContain(`data-sid="${sid}"`)
  })

  it('replaces authored instrumentation SIDs in the parsed DOM', async () => {
    const { instrumentForPreview } = await import('../src/renderer/preview/instrument')
    const src =
      '<p data-sid="kept" DATA-GX-SID="999>" data-gx-sid=888>one</p><div title="data-gx-sid=&quot;777&quot;">two</div>'
    const map = buildParseMap(src, 1)
    const out = instrumentForPreview(src, map, 'inspector')
    const dom = new DOMParser().parseFromString(out, 'text/html')
    const p = dom.querySelector('p')!
    const div = dom.querySelector('div')!
    const pSid = map.elements.find((entry) => entry.tag === 'p')!.sid
    const divSid = map.elements.find((entry) => entry.tag === 'div')!.sid
    expect(p.getAttribute('data-sid')).toBe('kept')
    expect(p.getAttribute('data-gx-sid')).toBe(String(pSid))
    expect([...p.attributes].filter((attribute) => attribute.name === 'data-gx-sid')).toHaveLength(
      1,
    )
    expect(div.getAttribute('title')).toBe('data-gx-sid="777"')
    expect(div.getAttribute('data-gx-sid')).toBe(String(divSid))
  })
})
