/// <reference types="vite/client" />
import type { SlidesApi } from '../shared/ipc'

declare global {
  interface Window {
    slidesApi: SlidesApi
  }
}

export {}
