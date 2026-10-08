import { chmodSync, existsSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import {
  convertLegacyFile,
  defaultSofficeCandidates,
  findSoffice,
  LEGACY_RE,
  LegacyConvertError,
  legacyTargetFor,
  siblingTargetPath,
} from '../src/main/legacy-office/converter'

const fixture = (rel: string) => fileURLToPath(new URL(rel, import.meta.url))
const DOC = fixture('../../../packages/file-parse/tests/fixtures/legacy-sample.doc')
const PPT = fixture('../../../packages/file-parse/tests/fixtures/legacy-sample.ppt')
const XLS = fixture('fixtures/legacy/legacy-sample.xls')

function scratch(): string {
  return mkdtempSync(join(tmpdir(), 'legacy-office-test-'))
}

/** A stand-in soffice: a shell script that runs `body` with $OUT and $IN set. */
function fakeSoffice(body: string): string {
  const dir = scratch()
  const path = join(dir, 'soffice')
  writeFileSync(
    path,
    `#!/bin/sh\nwhile [ $# -gt 0 ]; do case "$1" in --outdir) OUT="$2"; shift;; --convert-to) FILTER="$2"; shift;; *) IN="$1";; esac; shift; done\n${body}\n`,
  )
  chmodSync(path, 0o755)
  return path
}

describe('legacy format table', () => {
  it('maps each legacy extension to the editor format it becomes', () => {
    expect(legacyTargetFor('/a/Report.DOC')).toBe('docx')
    expect(legacyTargetFor('/a/notes.rtf')).toBe('docx')
    expect(legacyTargetFor('/a/b.odt')).toBe('docx')
    expect(legacyTargetFor('/a/Budget.xls')).toBe('xlsx')
    expect(legacyTargetFor('/a/b.ods')).toBe('xlsx')
    expect(legacyTargetFor('/a/Talk.ppt')).toBe('pptx')
    expect(legacyTargetFor('/a/b.odp')).toBe('pptx')
    expect(legacyTargetFor('/a/b.docx')).toBeUndefined()
    expect(LEGACY_RE.test('x.PPT')).toBe(true)
    expect(LEGACY_RE.test('x.pptx')).toBe(false)
  })
})

describe('siblingTargetPath', () => {
  it('names the copy after the original, beside it', () => {
    expect(siblingTargetPath('/docs/Report.doc', () => false)).toBe('/docs/Report.docx')
  })

  it('never reuses a name that already exists', () => {
    const taken = new Set(['/docs/Report.docx', '/docs/Report (converted).docx'])
    expect(siblingTargetPath('/docs/Report.doc', (p) => taken.has(p))).toBe(
      '/docs/Report (converted 2).docx',
    )
  })
})

describe('findSoffice', () => {
  it('prefers the chosen path, then the first existing candidate', () => {
    const real = fakeSoffice('exit 0')
    expect(findSoffice({ chosenPath: real, candidates: [], useRegistry: false })).toBe(real)
    expect(
      findSoffice({
        chosenPath: '/missing/soffice',
        candidates: ['/missing/too', real],
        useRegistry: false,
      }),
    ).toBe(real)
    expect(findSoffice({ candidates: ['/missing'], useRegistry: false })).toBeUndefined()
  })

  it('lists the app add-on folder first, then the standard Windows installs', () => {
    const list = defaultSofficeCandidates(
      { ProgramFiles: 'C:\\Program Files', LOCALAPPDATA: 'C:\\Users\\u\\AppData\\Local' },
      '/addon',
      'win32',
    )
    expect(list[0]).toBe(join('/addon', 'program', 'soffice.exe'))
    expect(list).toContain(join('C:\\Program Files', 'LibreOffice', 'program', 'soffice.exe'))
    expect(list.at(-1)).toBe(
      join('C:\\Users\\u\\AppData\\Local', 'Programs', 'LibreOffice', 'program', 'soffice.exe'),
    )
  })
})

describe.skipIf(process.platform === 'win32')('convertLegacyFile with a stand-in converter', () => {
  const base = () => {
    const dir = scratch()
    return { dir, profileDir: join(dir, 'profile'), tempRoot: join(dir, 'tmp') }
  }

  it('moves the converter output to the destination and cleans up', async () => {
    const { dir, profileDir, tempRoot } = base()
    const soffice = fakeSoffice('printf "converted:%s" "$FILTER" > "$OUT/source.docx"')
    const destination = join(dir, 'Report.docx')
    await convertLegacyFile({ soffice, source: DOC, destination, profileDir, tempRoot })
    expect(readFileSync(destination, 'utf8')).toBe('converted:docx:MS Word 2007 XML')
    expect(existsSync(DOC)).toBe(true)
  })

  it('fails when the converter writes nothing', async () => {
    const { dir, profileDir, tempRoot } = base()
    const soffice = fakeSoffice('exit 0')
    const destination = join(dir, 'Report.docx')
    await expect(
      convertLegacyFile({ soffice, source: DOC, destination, profileDir, tempRoot }),
    ).rejects.toMatchObject({ reason: 'no-output' })
    expect(existsSync(destination)).toBe(false)
  })

  it('fails on a non-zero exit', async () => {
    const { dir, profileDir, tempRoot } = base()
    const soffice = fakeSoffice('echo broken >&2; exit 3')
    await expect(
      convertLegacyFile({
        soffice,
        source: DOC,
        destination: join(dir, 'R.docx'),
        profileDir,
        tempRoot,
      }),
    ).rejects.toBeInstanceOf(LegacyConvertError)
  })

  it('kills a converter that hangs', async () => {
    const { dir, profileDir, tempRoot } = base()
    const soffice = fakeSoffice('sleep 30')
    await expect(
      convertLegacyFile({
        soffice,
        source: DOC,
        destination: join(dir, 'R.docx'),
        profileDir,
        tempRoot,
        timeoutMs: 300,
      }),
    ).rejects.toMatchObject({ reason: 'timeout' })
  })
})

// Real LibreOffice, when this machine has it (developer boxes usually do; CI may not).
const realSoffice = findSoffice({
  candidates: defaultSofficeCandidates(process.env, '/nonexistent'),
  useRegistry: false,
})

describe.skipIf(!realSoffice)('convertLegacyFile with LibreOffice', () => {
  const cases = [
    { source: DOC, out: 'Report.docx', part: 'word/document.xml', text: 'Legacy DOC body text' },
    { source: XLS, out: 'Budget.xlsx', part: 'xl/sharedStrings.xml', text: 'Apples' },
    { source: PPT, out: 'Talk.pptx', part: 'ppt/presentation.xml', text: 'sldIdLst' },
  ]
  for (const c of cases) {
    it(`converts ${c.source.slice(-3)} to ${c.out.slice(-4)}`, { timeout: 120_000 }, async () => {
      const dir = scratch()
      const destination = join(dir, c.out)
      await convertLegacyFile({
        soffice: realSoffice!,
        source: c.source,
        destination,
        profileDir: join(dir, 'profile'),
        tempRoot: join(dir, 'tmp'),
      })
      const zip = await JSZip.loadAsync(readFileSync(destination))
      expect(await zip.file(c.part)?.async('string')).toContain(c.text)
    })
  }
})
