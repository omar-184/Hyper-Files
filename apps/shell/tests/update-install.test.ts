import { createHash } from 'node:crypto'
import { existsSync, mkdtempSync, readFileSync, readdirSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'
import {
  createUpdateInstaller,
  parseSha256File,
  pickInstallerAsset,
} from '../src/main/update-install'
import type { UpdateInstallDeps } from '../src/main/update-install'

const BASE = 'https://github.com/omar-184/Hypercube-Office/releases/download/v0.2.0/'
const EXE = 'Hypercube-Office-Setup-0.2.0-x64.exe'
const PAYLOAD = Buffer.from('MZ fake installer bytes '.repeat(50_000))
const SHA = createHash('sha256').update(PAYLOAD).digest('hex')

function release(overrides: { digest?: string; exeUrl?: string; noSum?: boolean } = {}) {
  const assets: Record<string, unknown>[] = [
    {
      name: EXE,
      browser_download_url: overrides.exeUrl ?? BASE + EXE,
      size: PAYLOAD.length,
      digest: overrides.digest ?? `sha256:${SHA}`,
    },
  ]
  if (!overrides.noSum) {
    assets.push({ name: `${EXE}.sha256`, browser_download_url: `${BASE}${EXE}.sha256` })
  }
  return { tag_name: 'v0.2.0', assets }
}

function stream(buf: Buffer, chunk = 64 * 1024): ReadableStream<Uint8Array> {
  let offset = 0
  return new ReadableStream({
    pull(controller) {
      if (offset >= buf.length) return controller.close()
      controller.enqueue(new Uint8Array(buf.subarray(offset, offset + chunk)))
      offset += chunk
    },
  })
}

let dirs: string[] = []
afterEach(() => {
  for (const d of dirs) rmSync(d, { recursive: true, force: true })
  dirs = []
})

function makeDeps(opts: { sumText?: string; payload?: Buffer } = {}) {
  const root = mkdtempSync(join(tmpdir(), 'hf-update-'))
  dirs.push(root)
  const statuses: string[] = []
  const deps: UpdateInstallDeps = {
    supported: true,
    downloadDir: join(root, 'dl'),
    fetch: vi.fn(async (url: string) => {
      if (url.endsWith('.sha256')) {
        return new Response(opts.sumText ?? `${SHA}  ${EXE}\n`)
      }
      const body = opts.payload ?? PAYLOAD
      return new Response(stream(body), {
        headers: { 'content-length': String(body.length) },
      })
    }),
    confirm: vi.fn(async () => true),
    quit: vi.fn(),
    launch: vi.fn(),
    onStatus: (s) => statuses.push(s.state),
  }
  return { deps, statuses }
}

describe('pickInstallerAsset', () => {
  it('finds the installer and checksum of this repository release', () => {
    expect(pickInstallerAsset(release(), '0.2.0')).toEqual({
      name: EXE,
      url: BASE + EXE,
      size: PAYLOAD.length,
      sha256Url: `${BASE}${EXE}.sha256`,
      digest: SHA,
    })
  })

  it('rejects assets hosted elsewhere, missing checksums and other versions', () => {
    expect(
      pickInstallerAsset(release({ exeUrl: `https://evil.example/${EXE}` }), '0.2.0'),
    ).toBeNull()
    expect(pickInstallerAsset(release({ noSum: true }), '0.2.0')).toBeNull()
    expect(pickInstallerAsset(release(), '0.3.0')).toBeNull()
    expect(pickInstallerAsset(null, '0.2.0')).toBeNull()
  })
})

describe('parseSha256File', () => {
  it('reads sha256sum output and checks the file name', () => {
    expect(parseSha256File(`${SHA.toUpperCase()}  ${EXE}\n`, EXE)).toBe(SHA)
    expect(parseSha256File(SHA, EXE)).toBe(SHA)
    expect(parseSha256File(`${SHA}  other.exe`, EXE)).toBeNull()
    expect(parseSha256File('not a hash', EXE)).toBeNull()
  })
})

describe('createUpdateInstaller', () => {
  it('is unsupported outside packaged Windows builds and makes no request', async () => {
    const { deps } = makeDeps()
    const installer = createUpdateInstaller({ ...deps, supported: false })
    expect(installer.status(release(), '0.2.0').state).toBe('unsupported')
    expect((await installer.download(release(), '0.2.0')).state).toBe('unsupported')
    expect(deps.fetch).not.toHaveBeenCalled()
  })

  it('downloads, verifies, and installs only after confirmation and quit', async () => {
    const { deps, statuses } = makeDeps()
    const installer = createUpdateInstaller(deps)
    expect(installer.status(release(), '0.2.0').state).toBe('idle')

    const done = await installer.download(release(), '0.2.0')
    expect(done).toMatchObject({ state: 'ready', version: '0.2.0', received: PAYLOAD.length })
    expect(statuses).toContain('downloading')
    const file = join(deps.downloadDir, EXE)
    expect(readFileSync(file).equals(PAYLOAD)).toBe(true)
    expect(readdirSync(deps.downloadDir)).toEqual([EXE])

    expect(await installer.install()).toBe(true)
    expect(deps.confirm).toHaveBeenCalledWith('0.2.0')
    expect(deps.quit).toHaveBeenCalledTimes(1)
    expect(deps.launch).not.toHaveBeenCalled()
    installer.launchPending()
    expect(deps.launch).toHaveBeenCalledWith(file)
    installer.launchPending()
    expect(deps.launch).toHaveBeenCalledTimes(1)
  })

  it('does nothing when the user cancels the install prompt', async () => {
    const { deps } = makeDeps()
    deps.confirm = vi.fn(async () => false)
    const installer = createUpdateInstaller(deps)
    await installer.download(release(), '0.2.0')
    expect(await installer.install()).toBe(false)
    installer.launchPending()
    expect(deps.quit).not.toHaveBeenCalled()
    expect(deps.launch).not.toHaveBeenCalled()
  })

  it('discards a download whose checksum does not match', async () => {
    const { deps } = makeDeps({ payload: Buffer.from('tampered') })
    const installer = createUpdateInstaller(deps)
    expect((await installer.download(release(), '0.2.0')).state).toBe('failed')
    expect(existsSync(join(deps.downloadDir, EXE))).toBe(false)
    expect(readdirSync(deps.downloadDir)).toEqual([])
    expect(await installer.install()).toBe(false)
  })

  it('fails when the .sha256 file and the GitHub digest disagree', async () => {
    const { deps } = makeDeps()
    const installer = createUpdateInstaller(deps)
    const other = 'f'.repeat(64)
    expect((await installer.download(release({ digest: `sha256:${other}` }), '0.2.0')).state).toBe(
      'failed',
    )
    expect(deps.fetch).toHaveBeenCalledTimes(1)
  })

  it('reports network errors as failed', async () => {
    const { deps } = makeDeps()
    deps.fetch = vi.fn(async () => {
      throw new Error('offline')
    })
    const installer = createUpdateInstaller(deps)
    expect((await installer.download(release(), '0.2.0')).state).toBe('failed')
    expect(installer.status(release(), '0.2.0').state).toBe('failed')
  })

  it('cleans up old downloads at startup', async () => {
    const { deps } = makeDeps()
    const first = createUpdateInstaller(deps)
    await first.download(release(), '0.2.0')
    await createUpdateInstaller(deps).cleanup()
    expect(existsSync(deps.downloadDir)).toBe(false)
  })
})
