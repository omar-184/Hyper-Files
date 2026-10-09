/**
 * In-app update install for packaged Windows builds. Nothing here runs on
 * its own: the user clicks "Download and install" in Settings → About after a
 * check found a newer release, then confirms "Restart and install".
 *
 * Download: the release's NSIS installer and its `.sha256` file, both from
 * this repository's GitHub release over HTTPS. The installer is streamed to
 * a temp folder (never held in memory) and its SHA-256 must match the
 * published `.sha256` file and, when GitHub reports one, the asset digest.
 * Anything else is discarded.
 *
 * Install: the app quits through the normal close flow (unsaved documents
 * prompt first); once it is really quitting, the verified installer is
 * started detached with electron-builder's update flags and reopens the app.
 * The installer is not code-signed, so this relies on HTTPS plus the
 * checksum: it catches a corrupted or tampered download, not a compromised
 * release. Files written by the app carry no Mark-of-the-Web, so SmartScreen
 * does not prompt.
 */

import { createHash } from 'node:crypto'
import { createWriteStream } from 'node:fs'
import { mkdir, rename, rm } from 'node:fs/promises'
import { join } from 'node:path'

import type { UpdateInstallStatus } from '../shared/home-api'
import { REPO } from './update-check'

const DOWNLOAD_PREFIX = `https://github.com/${REPO}/releases/download/`
/** a release installer is ~140 MB; anything far larger is not ours */
const MAX_INSTALLER_BYTES = 1024 * 1024 * 1024
const SAFE_NAME = /^[A-Za-z0-9._-]+$/
const SHA256_HEX = /^[0-9a-f]{64}$/

export interface InstallerAsset {
  name: string
  url: string
  /** published size in bytes, 0 when unknown */
  size: number
  sha256Url: string
  /** GitHub's own asset digest (lowercase hex), when the API reports one */
  digest: string | null
}

interface ReleaseAsset {
  name: string
  browser_download_url: string
  size?: unknown
  digest?: unknown
}

/** the x64 NSIS installer of `version` and its .sha256 sibling, both from this repository's release */
export function pickInstallerAsset(release: unknown, version: string): InstallerAsset | null {
  const r = (release ?? {}) as { tag_name?: unknown; assets?: unknown }
  if (typeof r.tag_name !== 'string' || !Array.isArray(r.assets)) return null
  const prefix = `${DOWNLOAD_PREFIX}${r.tag_name}/`
  const assets = r.assets.filter(
    (a): a is ReleaseAsset =>
      !!a &&
      typeof a.name === 'string' &&
      SAFE_NAME.test(a.name) &&
      typeof a.browser_download_url === 'string' &&
      a.browser_download_url === prefix + a.name,
  )
  const exe = assets.find((a) => a.name.endsWith(`-Setup-${version}-x64.exe`))
  if (!exe) return null
  const sum = assets.find((a) => a.name === `${exe.name}.sha256`)
  if (!sum) return null
  const digest =
    typeof exe.digest === 'string' && /^sha256:[0-9a-f]{64}$/i.test(exe.digest)
      ? exe.digest.slice('sha256:'.length).toLowerCase()
      : null
  return {
    name: exe.name,
    url: exe.browser_download_url,
    size: typeof exe.size === 'number' && exe.size > 0 ? exe.size : 0,
    sha256Url: sum.browser_download_url,
    digest,
  }
}

/** `<hex>  <file name>` (sha256sum format); the name, when present, must be the installer's */
export function parseSha256File(text: string, name: string): string | null {
  const [hash, file] = text.trim().split(/\s+/, 2)
  const hex = (hash ?? '').toLowerCase()
  if (!SHA256_HEX.test(hex)) return null
  if (file !== undefined && file.replace(/^\*/, '') !== name) return null
  return hex
}

export interface UpdateInstallDeps {
  /** packaged Windows build; everything else keeps the release-page link */
  supported: boolean
  /** folder the installer is downloaded into; emptied before each download */
  downloadDir: string
  /** GET a URL (redirects followed); rejects on network errors */
  fetch: (url: string) => Promise<Response>
  /** native "Install version X?" prompt */
  confirm: (version: string) => Promise<boolean>
  /** quit through the normal close flow; unsaved documents may cancel it */
  quit: () => void
  /** start the installer detached; called once the app is really quitting */
  launch: (installerPath: string) => void
  onStatus: (status: UpdateInstallStatus) => void
}

export interface UpdateInstaller {
  /** status for the release a check returned (null: no newer release) */
  status(release: unknown, version: string | null): UpdateInstallStatus
  download(release: unknown, version: string | null): Promise<UpdateInstallStatus>
  install(): Promise<boolean>
  /** will-quit hook: runs the installer when the user asked for it */
  launchPending(): void
  /** drop installers left from earlier downloads (call at startup) */
  cleanup(): Promise<void>
}

export function createUpdateInstaller(deps: UpdateInstallDeps): UpdateInstaller {
  let current: UpdateInstallStatus = { state: 'idle', received: 0, total: 0 }
  let readyPath: string | null = null
  let pendingLaunch: string | null = null
  let inflight: Promise<UpdateInstallStatus> | null = null

  const unsupported = (): UpdateInstallStatus => ({ state: 'unsupported', received: 0, total: 0 })

  function set(next: UpdateInstallStatus): UpdateInstallStatus {
    current = next
    deps.onStatus(next)
    return next
  }

  function status(release: unknown, version: string | null): UpdateInstallStatus {
    if (!deps.supported || !version) return unsupported()
    // a download or a verified installer of this version wins over the release lookup
    if (current.version === version && current.state !== 'idle') return current
    if (!pickInstallerAsset(release, version)) return unsupported()
    return { state: 'idle', version, received: 0, total: 0 }
  }

  async function fetchOk(url: string): Promise<Response> {
    const res = await deps.fetch(url)
    if (!res.ok || !res.body) throw new Error(`HTTP ${res.status}`)
    return res
  }

  async function run(asset: InstallerAsset, version: string): Promise<UpdateInstallStatus> {
    readyPath = null
    const fail = () => set({ state: 'failed', version, received: 0, total: asset.size })
    const target = join(deps.downloadDir, asset.name)
    const partial = `${target}.partial`
    try {
      await rm(deps.downloadDir, { recursive: true, force: true })
      await mkdir(deps.downloadDir, { recursive: true })
      const expected = parseSha256File(await (await fetchOk(asset.sha256Url)).text(), asset.name)
      if (!expected || (asset.digest && asset.digest !== expected)) return fail()

      const res = await fetchOk(asset.url)
      const total = Number(res.headers.get('content-length')) || asset.size
      if (total > MAX_INSTALLER_BYTES) return fail()
      const hash = createHash('sha256')
      const out = createWriteStream(partial)
      let received = 0
      let lastReport = 0
      set({ state: 'downloading', version, received, total })
      try {
        const reader = res.body!.getReader()
        for (;;) {
          const { done, value } = await reader.read()
          if (done) break
          received += value.byteLength
          if (received > MAX_INSTALLER_BYTES) throw new Error('installer too large')
          hash.update(value)
          if (!out.write(value)) await new Promise<void>((r) => out.once('drain', r))
          if (received - lastReport >= 1024 * 1024) {
            lastReport = received
            set({ state: 'downloading', version, received, total })
          }
        }
      } finally {
        await new Promise<void>((resolve, reject) => {
          out.once('error', reject)
          out.end(resolve)
        })
      }
      if (hash.digest('hex') !== expected) {
        await rm(partial, { force: true })
        return fail()
      }
      await rename(partial, target)
      readyPath = target
      return set({ state: 'ready', version, received, total: received })
    } catch {
      await rm(partial, { force: true }).catch(() => {})
      return fail()
    }
  }

  function download(release: unknown, version: string | null): Promise<UpdateInstallStatus> {
    if (inflight) return inflight
    if (!deps.supported || !version) return Promise.resolve(unsupported())
    if (current.state === 'ready' && current.version === version && readyPath) {
      return Promise.resolve(current)
    }
    const asset = pickInstallerAsset(release, version)
    if (!asset) return Promise.resolve(unsupported())
    inflight = run(asset, version).finally(() => {
      inflight = null
    })
    return inflight
  }

  async function install(): Promise<boolean> {
    if (current.state !== 'ready' || !readyPath || !current.version) return false
    if (!(await deps.confirm(current.version))) return false
    // runs on will-quit, so a quit the user cancels (unsaved work) installs nothing now;
    // the update then installs whenever the app next quits
    pendingLaunch = readyPath
    deps.quit()
    return true
  }

  function launchPending(): void {
    if (!pendingLaunch) return
    const path = pendingLaunch
    pendingLaunch = null
    try {
      deps.launch(path)
    } catch {
      // the installer stays on disk; the next "Download and install" starts over
    }
  }

  async function cleanup(): Promise<void> {
    if (!deps.supported || inflight || readyPath) return
    await rm(deps.downloadDir, { recursive: true, force: true }).catch(() => {})
  }

  return { status, download, install, launchPending, cleanup }
}
