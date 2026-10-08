/**
 * Update check. The background check is off until the user turns it on in
 * Settings → About; while off, Hyper-Files makes no network request of its
 * own except when the user clicks "Check now". When on, it asks the GitHub
 * Releases API for the latest published release at most once a day and, if
 * that release is newer, shows one notification per version. Nothing is
 * downloaded here: downloading and installing is a separate, user-started
 * step (update-install.ts). The request carries no identifiers beyond what any HTTPS
 * request to api.github.com does (IP address, a User-Agent with the version).
 */

import type { UpdateCheckResult, UpdateCheckState } from '../shared/home-api'

export const UPDATE_CHECK_KEY = 'updateCheck'
export const UPDATE_LAST_CHECK_KEY = 'updateCheckLastAt'
export const UPDATE_NOTIFIED_KEY = 'updateNotifiedVersion'

export const REPO = 'omar-184/Hyper-Files'
export const LATEST_RELEASE_API = `https://api.github.com/repos/${REPO}/releases/latest`
export const RELEASES_PAGE = `https://github.com/${REPO}/releases`

export const CHECK_INTERVAL_MS = 24 * 60 * 60 * 1000
/** first check waits for the app to settle; later ones re-evaluate on this tick */
const STARTUP_DELAY_MS = 30_000
const TICK_MS = 60 * 60 * 1000

/** '1.2.3' / 'v1.2.3' / '1.2.3-beta.1' → [1, 2, 3] plus the prerelease tail */
function parseVersion(v: string): { nums: number[]; pre: string } | null {
  const m = /^v?(\d+)\.(\d+)\.(\d+)(?:-([0-9A-Za-z.-]+))?(?:\+.*)?$/.exec(v.trim())
  if (!m) return null
  return { nums: [Number(m[1]), Number(m[2]), Number(m[3])], pre: m[4] ?? '' }
}

/** >0 when a is newer than b; null when either is not a version */
export function compareVersions(a: string, b: string): number | null {
  const pa = parseVersion(a)
  const pb = parseVersion(b)
  if (!pa || !pb) return null
  for (let i = 0; i < 3; i++) {
    if (pa.nums[i] !== pb.nums[i]) return pa.nums[i] - pb.nums[i]
  }
  // a release outranks its own prereleases
  if (pa.pre === pb.pre) return 0
  if (!pa.pre) return 1
  if (!pb.pre) return -1
  return pa.pre.localeCompare(pb.pre, undefined, { numeric: true })
}

/** only release pages of this repository may be opened from the result */
export function safeReleaseUrl(url: unknown): string {
  return typeof url === 'string' && url.startsWith(`https://github.com/${REPO}/releases/`)
    ? url
    : RELEASES_PAGE
}

export function interpretLatestRelease(json: unknown, currentVersion: string): UpdateCheckResult {
  const r = (json ?? {}) as { tag_name?: unknown; html_url?: unknown; draft?: unknown }
  const checkedAt = Date.now()
  if (typeof r.tag_name !== 'string' || r.draft === true) {
    return { state: 'failed', checkedAt }
  }
  const cmp = compareVersions(r.tag_name, currentVersion)
  if (cmp === null) return { state: 'failed', checkedAt }
  const version = r.tag_name.replace(/^v/, '')
  if (cmp <= 0) return { state: 'latest', version, checkedAt }
  return { state: 'available', version, url: safeReleaseUrl(r.html_url), checkedAt }
}

export interface UpdateCheckDeps {
  currentVersion: string
  /** GET a URL and parse the JSON body; rejects on network/HTTP errors */
  fetchJson: (url: string) => Promise<{ status: number; body: unknown }>
  readSettings: () => Record<string, unknown>
  writeSettings: (updates: Record<string, unknown>) => void
  /** show a system notification; clicking it opens `url` */
  notify: (version: string, url: string) => void
  now?: () => number
}

export interface UpdateChecker {
  enabled(): boolean
  setEnabled(on: boolean): void
  last(): UpdateCheckResult | null
  /** raw GitHub release payload behind last(); update-install.ts picks the installer from it */
  lastRelease(): unknown
  /** run a check now (Settings "Check now"); never throws */
  checkNow(): Promise<UpdateCheckResult>
  /** begin the background schedule; it does nothing while the setting is off */
  start(): void
  stop(): void
}

export function createUpdateChecker(deps: UpdateCheckDeps): UpdateChecker {
  const now = deps.now ?? Date.now
  let lastResult: UpdateCheckResult | null = null
  let lastReleaseBody: unknown = null
  let inflight: Promise<UpdateCheckResult> | null = null
  let startupTimer: ReturnType<typeof setTimeout> | null = null
  let tickTimer: ReturnType<typeof setInterval> | null = null

  const enabled = () => deps.readSettings()[UPDATE_CHECK_KEY] === true

  async function run(): Promise<UpdateCheckResult> {
    let result: UpdateCheckResult
    lastReleaseBody = null
    try {
      const res = await deps.fetchJson(LATEST_RELEASE_API)
      if (res.status >= 200 && res.status < 300) lastReleaseBody = res.body
      // 404 = the repository has no published release yet
      result =
        res.status === 404
          ? { state: 'latest', checkedAt: now() }
          : res.status >= 200 && res.status < 300
            ? interpretLatestRelease(res.body, deps.currentVersion)
            : { state: 'failed', checkedAt: now() }
    } catch {
      result = { state: 'failed', checkedAt: now() }
    }
    lastResult = result
    if (result.state !== 'failed') {
      try {
        deps.writeSettings({ [UPDATE_LAST_CHECK_KEY]: result.checkedAt })
      } catch {
        // unwritable settings only cost an extra check next launch
      }
    }
    return result
  }

  function checkNow(): Promise<UpdateCheckResult> {
    inflight ??= run().finally(() => {
      inflight = null
    })
    return inflight
  }

  async function backgroundCheck(): Promise<void> {
    if (!enabled()) return
    const lastAt = Number(deps.readSettings()[UPDATE_LAST_CHECK_KEY]) || 0
    if (now() - lastAt < CHECK_INTERVAL_MS) return
    const result = await checkNow()
    if (result.state !== 'available' || !result.version || !result.url) return
    if (deps.readSettings()[UPDATE_NOTIFIED_KEY] === result.version) return
    try {
      deps.writeSettings({ [UPDATE_NOTIFIED_KEY]: result.version })
    } catch {
      // still notify; at worst the same version is announced again next launch
    }
    deps.notify(result.version, result.url)
  }

  function start(): void {
    if (startupTimer || tickTimer) return
    startupTimer = setTimeout(() => {
      startupTimer = null
      void backgroundCheck()
    }, STARTUP_DELAY_MS)
    tickTimer = setInterval(() => void backgroundCheck(), TICK_MS)
    tickTimer.unref?.()
  }

  function stop(): void {
    if (startupTimer) clearTimeout(startupTimer)
    if (tickTimer) clearInterval(tickTimer)
    startupTimer = null
    tickTimer = null
  }

  return {
    enabled,
    setEnabled(on) {
      deps.writeSettings({ [UPDATE_CHECK_KEY]: on })
      if (!on) lastResult = null
    },
    last: () => lastResult,
    lastRelease: () => lastReleaseBody,
    checkNow,
    start,
    stop,
  }
}

export type { UpdateCheckResult, UpdateCheckState }
