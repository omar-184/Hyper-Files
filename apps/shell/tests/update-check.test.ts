import { afterEach, describe, expect, it, vi } from 'vitest'
import {
  CHECK_INTERVAL_MS,
  LATEST_RELEASE_API,
  RELEASES_PAGE,
  UPDATE_CHECK_KEY,
  UPDATE_LAST_CHECK_KEY,
  UPDATE_NOTIFIED_KEY,
  compareVersions,
  createUpdateChecker,
  interpretLatestRelease,
  safeReleaseUrl,
} from '../src/main/update-check'
import type { UpdateCheckDeps } from '../src/main/update-check'

const RELEASE = {
  tag_name: 'v0.2.0',
  html_url: 'https://github.com/omar-184/Hypercube-Office/releases/tag/v0.2.0',
  draft: false,
}

function makeDeps(settings: Record<string, unknown>, overrides: Partial<UpdateCheckDeps> = {}) {
  const store = { ...settings }
  const deps: UpdateCheckDeps = {
    currentVersion: '0.1.0',
    fetchJson: vi.fn(async () => ({ status: 200, body: RELEASE })),
    readSettings: () => store,
    writeSettings: (u) => Object.assign(store, u),
    notify: vi.fn(),
    ...overrides,
  }
  return { deps, store }
}

afterEach(() => {
  vi.useRealTimers()
})

describe('compareVersions', () => {
  it('orders numeric parts and prereleases', () => {
    expect(compareVersions('v0.2.0', '0.1.0')).toBeGreaterThan(0)
    expect(compareVersions('0.1.0', '0.1.0')).toBe(0)
    expect(compareVersions('0.10.0', '0.9.9')).toBeGreaterThan(0)
    expect(compareVersions('1.0.0', '1.0.0-beta.2')).toBeGreaterThan(0)
    expect(compareVersions('1.0.0-beta.10', '1.0.0-beta.2')).toBeGreaterThan(0)
    expect(compareVersions('nightly', '0.1.0')).toBeNull()
  })
})

describe('interpretLatestRelease', () => {
  it('reports a newer release with its page', () => {
    expect(interpretLatestRelease(RELEASE, '0.1.0')).toMatchObject({
      state: 'available',
      version: '0.2.0',
      url: RELEASE.html_url,
    })
  })

  it('reports latest when the release is not newer', () => {
    expect(interpretLatestRelease(RELEASE, '0.2.0').state).toBe('latest')
  })

  it('never hands out a link outside this repository', () => {
    const r = interpretLatestRelease({ ...RELEASE, html_url: 'https://evil.example/x' }, '0.1.0')
    expect(r.url).toBe(RELEASES_PAGE)
    expect(safeReleaseUrl('https://github.com/someone-else/repo/releases/tag/v9')).toBe(
      RELEASES_PAGE,
    )
  })

  it('fails on malformed payloads', () => {
    expect(interpretLatestRelease(null, '0.1.0').state).toBe('failed')
    expect(interpretLatestRelease({ tag_name: 'latest' }, '0.1.0').state).toBe('failed')
  })
})

describe('createUpdateChecker', () => {
  it('is off by default and makes no request in the background while off', async () => {
    vi.useFakeTimers()
    const { deps } = makeDeps({})
    const checker = createUpdateChecker(deps)
    expect(checker.enabled()).toBe(false)
    checker.start()
    await vi.advanceTimersByTimeAsync(3 * CHECK_INTERVAL_MS)
    checker.stop()
    expect(deps.fetchJson).not.toHaveBeenCalled()
  })

  it('checks once a day when on and notifies once per version', async () => {
    vi.useFakeTimers()
    const { deps, store } = makeDeps({ [UPDATE_CHECK_KEY]: true })
    const checker = createUpdateChecker(deps)
    checker.start()
    await vi.advanceTimersByTimeAsync(60_000)
    expect(deps.fetchJson).toHaveBeenCalledTimes(1)
    expect(deps.fetchJson).toHaveBeenCalledWith(LATEST_RELEASE_API)
    expect(deps.notify).toHaveBeenCalledWith('0.2.0', RELEASE.html_url)
    expect(store[UPDATE_NOTIFIED_KEY]).toBe('0.2.0')
    expect(typeof store[UPDATE_LAST_CHECK_KEY]).toBe('number')

    // a few hours later: no new request inside the 24 h window
    await vi.advanceTimersByTimeAsync(5 * 60 * 60 * 1000)
    expect(deps.fetchJson).toHaveBeenCalledTimes(1)

    // next day: checks again, same version is not announced twice
    await vi.advanceTimersByTimeAsync(CHECK_INTERVAL_MS)
    checker.stop()
    expect(deps.fetchJson).toHaveBeenCalledTimes(2)
    expect(deps.notify).toHaveBeenCalledTimes(1)
  })

  it('treats a repository without releases as up to date', async () => {
    const { deps } = makeDeps(
      { [UPDATE_CHECK_KEY]: true },
      { fetchJson: vi.fn(async () => ({ status: 404, body: null })) },
    )
    expect((await createUpdateChecker(deps).checkNow()).state).toBe('latest')
  })

  it('reports network errors as failed without recording a check', async () => {
    const { deps, store } = makeDeps(
      { [UPDATE_CHECK_KEY]: true },
      {
        fetchJson: vi.fn(async () => {
          throw new Error('offline')
        }),
      },
    )
    expect((await createUpdateChecker(deps).checkNow()).state).toBe('failed')
    expect(store[UPDATE_LAST_CHECK_KEY]).toBeUndefined()
  })

  it('persists the toggle', () => {
    const { deps, store } = makeDeps({})
    const checker = createUpdateChecker(deps)
    checker.setEnabled(true)
    expect(store[UPDATE_CHECK_KEY]).toBe(true)
    expect(checker.enabled()).toBe(true)
  })
})
