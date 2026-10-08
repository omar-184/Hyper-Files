import { createI18n, defineStrings } from '@genoffice/i18n'
import { zh } from './tools/zh'
import { en } from './tools/en'
import { ja } from './tools/ja'
import { ko } from './tools/ko'
import { fr } from './tools/fr'
import { de } from './tools/de'
import { es } from './tools/es'
import { th } from './tools/th'
import { id } from './tools/id'
import { ru } from './tools/ru'
import { ar } from './tools/ar'
import { pt } from './tools/pt'
import { it } from './tools/it'
import { pl } from './tools/pl'
import { cs } from './tools/cs'
import { nl } from './tools/nl'
import { ms } from './tools/ms'
import { he } from './tools/he'
import { hi } from './tools/hi'
import { zhTW } from './tools/zh-TW'
import { vi } from './tools/vi'

/** User-visible strings for the PDF tools area on Home */
export const toolStrings = defineStrings({
  zh,
  en,
  ja,
  ko,
  fr,
  de,
  es,
  th,
  id,
  ru,
  ar,
  pt,
  it,
  pl,
  cs,
  nl,
  ms,
  he,
  hi,
  'zh-TW': zhTW,
  vi,
})

export type ToolStringKey = keyof typeof toolStrings.zh

export const translateTools = createI18n(toolStrings)
