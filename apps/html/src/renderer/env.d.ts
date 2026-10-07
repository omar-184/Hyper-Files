/// <reference types="vite/client" />

import type { HtmlApi } from '../shared/ipc'

declare global {
  interface Window {
    htmlApi: HtmlApi
  }
}

export {}
