/**
 * @vitest-environment jsdom
 */
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import type { Root } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { HomeApi } from '../src/shared/home-api'
import { LocaleProvider } from '../src/renderer/src/locale'
import { SettingsModal } from '../src/renderer/src/SettingsModal'

const actEnvironment = globalThis as typeof globalThis & {
  IS_REACT_ACT_ENVIRONMENT?: boolean
}
actEnvironment.IS_REACT_ACT_ENVIRONMENT = true

let host: HTMLDivElement
let root: Root

beforeEach(() => {
  host = document.createElement('div')
  document.body.append(host)
  root = createRoot(host)
})

afterEach(() => {
  act(() => root.unmount())
  host.remove()
})

async function click(button: HTMLButtonElement): Promise<void> {
  await act(async () => {
    button.click()
    await Promise.resolve()
  })
}

async function openSection(section: string, api: Partial<HomeApi>): Promise<void> {
  window.hyperFiles = {
    getTheme: async () => 'system',
    getDefaultSaveDir: async () => '',
    getAnalyticsEnabled: async () => true,
    setAnalyticsEnabled: async () => true,
    getAiPanelPrefs: async () => ({ fontSize: 'default', spellcheck: true }),
    setAiPanelPrefs: async (patch) => ({ fontSize: 'default', spellcheck: true, ...patch }),
    getUpdateChannel: async () => 'stable',
    getAppVersion: async () => '1.0.0',
    githubStars: async () => null,
    ...api,
  } as unknown as HomeApi

  await act(async () => {
    root.render(
      createElement(
        LocaleProvider,
        { initial: 'en' },
        createElement(SettingsModal, {
          status: null,
          loggingOut: false,
          loginWaiting: false,
          loginUrl: null,
          urlCopied: false,
          onOpenLoginUrl: vi.fn(),
          onCopyLoginUrl: vi.fn(),
          onClose: vi.fn(),
          onLogin: vi.fn(),
          onLogout: vi.fn(),
        }),
      ),
    )
    await Promise.resolve()
  })
  const nav = Array.from(host.querySelectorAll<HTMLButtonElement>('.set-nav-item')).find((button) =>
    button.textContent?.includes(section),
  )
  await click(nav!)
}

function row(): HTMLElement | null {
  return (
    Array.from(host.querySelectorAll<HTMLElement>('.set-field')).find((el) =>
      el.textContent?.includes('Check for updates'),
    ) ?? null
  )
}

describe('Settings update-check row', () => {
  it('is off by default but still checks on demand', async () => {
    const check = vi.fn(async () => ({ state: 'latest' as const, checkedAt: 1 }))
    await openSection('About', {
      getUpdateCheck: async () => ({ enabled: false, last: null }),
      checkForUpdates: check,
    })
    const field = row()!
    expect(field.textContent).toContain('Off by default.')
    expect(field.querySelector('[role="switch"]')!.getAttribute('aria-checked')).toBe('false')
    expect(check).not.toHaveBeenCalled()
    const checkNow = field.querySelector<HTMLButtonElement>('.set-btn')!
    expect(checkNow.textContent).toBe('Check now')
    await click(checkNow)
    await act(async () => {
      await Promise.resolve()
    })
    expect(check).toHaveBeenCalledTimes(1)
    expect(row()!.textContent).toContain('You have the latest version.')
  })

  it('turns on and offers the download page when the app cannot install itself', async () => {
    const setUpdateCheck = vi.fn(async () => {})
    const openUpdatePage = vi.fn(async () => {})
    await openSection('About', {
      getUpdateCheck: async () => ({ enabled: false, last: null }),
      setUpdateCheck,
      checkForUpdates: async () => ({ state: 'available', version: '0.2.0', checkedAt: 1 }),
      getUpdateInstall: async () => ({ state: 'unsupported', received: 0, total: 0 }),
      openUpdatePage,
    })
    await click(row()!.querySelector<HTMLButtonElement>('[role="switch"]')!)
    expect(setUpdateCheck).toHaveBeenCalledWith(true)
    await click(row()!.querySelector<HTMLButtonElement>('.set-btn')!)
    await act(async () => {
      await Promise.resolve()
    })
    expect(row()!.textContent).toContain('Version 0.2.0 is available.')
    const download = row()!.querySelector<HTMLButtonElement>('.set-btn')!
    expect(download.textContent).toBe('Download')
    await click(download)
    expect(openUpdatePage).toHaveBeenCalledTimes(1)
  })

  it('downloads and installs in the app', async () => {
    const downloadUpdate = vi.fn(async () => ({
      state: 'ready' as const,
      version: '0.2.0',
      received: 10,
      total: 10,
    }))
    const installUpdate = vi.fn(async () => true)
    await openSection('About', {
      getUpdateCheck: async () => ({
        enabled: true,
        last: { state: 'available', version: '0.2.0', checkedAt: 1 },
      }),
      getUpdateInstall: async () => ({ state: 'idle', version: '0.2.0', received: 0, total: 0 }),
      downloadUpdate,
      installUpdate,
    })
    await act(async () => {
      await Promise.resolve()
    })
    const start = row()!.querySelector<HTMLButtonElement>('.set-btn')!
    expect(start.textContent).toBe('Download and install')
    await click(start)
    await act(async () => {
      await Promise.resolve()
    })
    expect(downloadUpdate).toHaveBeenCalledTimes(1)
    expect(row()!.textContent).toContain('Version 0.2.0 is downloaded and verified.')
    const restart = row()!.querySelector<HTMLButtonElement>('.set-btn')!
    expect(restart.textContent).toBe('Restart and install')
    await click(restart)
    expect(installUpdate).toHaveBeenCalledTimes(1)
  })
})
