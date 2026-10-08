import type { Params } from '@genoffice/i18n'
import { translateTools, type ToolStringKey } from '../i18n/strings-tools'
import { useI18n } from '../locale'

export type ToolsT = (key: ToolStringKey, params?: Params) => string

/** The tools dictionary in the Home language. */
export function useToolsI18n(): { t: ToolsT; dateLocale: string } {
  const { lang, dateLocale } = useI18n()
  return { t: (key, params) => translateTools(lang, key, params), dateLocale }
}

export function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(0)} KB`
  if (bytes < 1024 * 1024 * 1024) return `${(bytes / 1024 / 1024).toFixed(1)} MB`
  return `${(bytes / 1024 / 1024 / 1024).toFixed(2)} GB`
}

export function baseName(path: string): string {
  return path.replace(/^.*[\\/]/, '')
}
