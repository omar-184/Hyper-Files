import { mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync, mkdirSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import {
  dictionaryLanguage,
  installedDictionaryLanguages,
  localDictionaryDownloadUrl,
  seedSpellDictionaries,
} from '../src/spell-dictionaries'

const dirs: string[] = []
const scratch = () => {
  const dir = mkdtempSync(join(tmpdir(), 'spell-dicts-'))
  dirs.push(dir)
  return dir
}
afterEach(() => {
  for (const dir of dirs.splice(0)) rmSync(dir, { recursive: true, force: true })
})

describe('dictionaryLanguage', () => {
  it('reads the language from Chromium dictionary file names', () => {
    expect(dictionaryLanguage('en-US-10-1.bdic')).toBe('en-US')
    expect(dictionaryLanguage('en-GB-oxendict-10-1.bdic')).toBe('en-GB-oxendict')
    expect(dictionaryLanguage('sh-4-0.bdic')).toBe('sh')
    expect(dictionaryLanguage('LICENSE-SCOWL.txt')).toBeNull()
    expect(dictionaryLanguage('en-US.bdic')).toBeNull()
  })
})

describe('seedSpellDictionaries', () => {
  it('copies shipped dictionaries into the target folder, creating it', () => {
    const bundled = scratch()
    writeFileSync(join(bundled, 'en-US-10-1.bdic'), 'us')
    writeFileSync(join(bundled, 'LICENSE-SCOWL.txt'), 'notice')
    const target = join(scratch(), 'Dictionaries')
    seedSpellDictionaries(bundled, target)
    expect(readdirSync(target)).toEqual(['en-US-10-1.bdic'])
    expect(readFileSync(join(target, 'en-US-10-1.bdic'), 'utf8')).toBe('us')
  })

  it('replaces a torn copy, keeps a complete one and leaves other dictionaries alone', () => {
    const bundled = scratch()
    writeFileSync(join(bundled, 'en-US-10-1.bdic'), 'complete')
    writeFileSync(join(bundled, 'en-GB-10-1.bdic'), 'complete')
    const target = scratch()
    writeFileSync(join(target, 'en-US-10-1.bdic'), 'torn')
    writeFileSync(join(target, 'en-GB-10-1.bdic'), 'existing')
    writeFileSync(join(target, 'de-DE-3-0.bdic'), 'downloaded earlier')
    seedSpellDictionaries(bundled, target)
    expect(readFileSync(join(target, 'en-US-10-1.bdic'), 'utf8')).toBe('complete')
    // same size: treated as up to date
    expect(readFileSync(join(target, 'en-GB-10-1.bdic'), 'utf8')).toBe('existing')
    expect(readFileSync(join(target, 'de-DE-3-0.bdic'), 'utf8')).toBe('downloaded earlier')
  })

  it('does nothing when the shipped folder is missing', () => {
    const target = join(scratch(), 'Dictionaries')
    expect(() => seedSpellDictionaries(join(scratch(), 'absent'), target)).not.toThrow()
    expect(installedDictionaryLanguages(target)).toEqual(new Set())
  })
})

describe('installedDictionaryLanguages', () => {
  it('lists the languages with a dictionary on disk', () => {
    const dir = scratch()
    writeFileSync(join(dir, 'en-US-10-1.bdic'), '')
    writeFileSync(join(dir, 'de-DE-3-0.bdic'), '')
    writeFileSync(join(dir, 'notes.txt'), '')
    mkdirSync(join(dir, 'sub'))
    expect(installedDictionaryLanguages(dir)).toEqual(new Set(['en-US', 'de-DE']))
  })
})

describe('localDictionaryDownloadUrl', () => {
  it('is a file URL ending in a slash', () => {
    const url = localDictionaryDownloadUrl(join(tmpdir(), 'dictionaries'))
    expect(url.startsWith('file://')).toBe(true)
    expect(url.endsWith('/dictionaries/')).toBe(true)
  })
})
