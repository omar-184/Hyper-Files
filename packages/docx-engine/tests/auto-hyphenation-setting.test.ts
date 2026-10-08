import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { parseDocx, saveDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const BODY = '<w:p><w:r><w:t>body</w:t></w:r></w:p>'
const SETTINGS_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml'
const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

const withSettings = (inner: string) =>
  buildDocx({
    bodyXml: BODY,
    extraParts: [
      {
        path: 'word/settings.xml',
        xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ${W_NS}>${inner}</w:settings>`,
        contentType: SETTINGS_TYPE,
      },
    ],
  })

async function settingsOf(bytes: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(bytes)
  return zip.file('word/settings.xml')!.async('string')
}

describe('Layout > Hyphenation setting', () => {
  it('reads the hyphenation zone and the caps rule', async () => {
    const parsed = await parseDocx(
      await withSettings(
        '<w:defaultTabStop w:val="720"/><w:autoHyphenation/><w:hyphenationZone w:val="425"/><w:doNotHyphenateCaps/>',
      ),
    )
    expect(parsed.autoHyphenation).toBe(true)
    expect(parsed.hyphenationZoneTwips).toBe(425)
    expect(parsed.doNotHyphenateCaps).toBe(true)
  })

  it('writes w:autoHyphenation in schema order and removes it again', async () => {
    const source = await withSettings(
      '<w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/><w:hyphenationZone w:val="357"/>',
    )
    const parsed = await parseDocx(source)
    const on = await settingsOf(
      await saveDocx(parsed, [{ kind: 'original', docxIndex: 0 }], { autoHyphenation: true }),
    )
    expect(on).toContain('<w:defaultTabStop w:val="720"/><w:autoHyphenation/><w:hyphenationZone')
    const reparsed = await parseDocx(
      await saveDocx(parsed, [{ kind: 'original', docxIndex: 0 }], { autoHyphenation: true }),
    )
    expect(reparsed.autoHyphenation).toBe(true)

    const offSource = await withSettings('<w:autoHyphenation w:val="true"/>')
    const off = await settingsOf(
      await saveDocx(await parseDocx(offSource), [{ kind: 'original', docxIndex: 0 }], {
        autoHyphenation: false,
      }),
    )
    expect(off).not.toContain('autoHyphenation')
  })
})
