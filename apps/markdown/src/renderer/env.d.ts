/// <reference types="vite/client" />

import type { MarkdownApi } from '../shared/ipc'

declare global {
  interface Window {
    markdownApi: MarkdownApi
  }
}

export {}
