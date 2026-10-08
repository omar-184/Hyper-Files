import type { DesktopApi } from '../shared/desktop-api'

declare global {
  interface Window {
    readonly desktopApi: DesktopApi
  }
}

export {}
