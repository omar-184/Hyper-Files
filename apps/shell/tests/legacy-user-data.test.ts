import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'

import { resolveUserDataDir } from '../src/main/legacy-user-data'

let appData: string
const current = () => join(appData, 'Hypercube Office')
const legacy = () => join(appData, 'Hyper-Files')

beforeEach(() => {
  appData = mkdtempSync(join(tmpdir(), 'hypercube-user-data-'))
})

afterEach(() => {
  rmSync(appData, { recursive: true, force: true })
})

describe('resolveUserDataDir', () => {
  it('moves the pre-rename Hyper-Files folder to the new name', () => {
    mkdirSync(legacy())
    writeFileSync(join(legacy(), 'app-settings.json'), '{"language":"ar"}')
    expect(resolveUserDataDir(appData, current())).toBe(current())
    expect(readFileSync(join(current(), 'app-settings.json'), 'utf8')).toBe('{"language":"ar"}')
    expect(existsSync(legacy())).toBe(false)
  })

  it('leaves both folders alone once the new one exists', () => {
    mkdirSync(legacy())
    mkdirSync(current())
    expect(resolveUserDataDir(appData, current())).toBe(current())
    expect(existsSync(legacy())).toBe(true)
  })

  it('uses the new folder on a fresh install', () => {
    expect(resolveUserDataDir(appData, current())).toBe(current())
  })

  it('keeps using the old folder when the move fails', () => {
    mkdirSync(legacy())
    // the target's parent does not exist, so rename throws
    const unreachable = join(appData, 'missing-parent', 'Hypercube Office')
    expect(resolveUserDataDir(appData, unreachable)).toBe(legacy())
  })
})
