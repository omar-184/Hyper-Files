import { createI18n, defineStrings } from '@genoffice/i18n'

const en = {
  secTemplates: 'Templates',
  tplLetter: 'Letter',
  tplLetterSub: 'Formal letter',
  tplResume: 'Resume',
  tplResumeSub: 'One page',
  tplBudget: 'Monthly budget',
  tplBudgetSub: 'Planned vs. actual',
  tplInvoice: 'Invoice',
  tplInvoiceSub: 'Items and tax',
  tplPresentation: 'Presentation',
  tplPresentationSub: 'Five slides',
}

/** Home ▸ Templates; English until each locale is translated (English UI first). */
export const templateStrings = defineStrings({
  zh: en,
  en,
  ja: en,
  ko: en,
  fr: en,
  de: en,
  es: en,
  th: en,
  id: en,
  ru: en,
  ar: en,
  pt: en,
  it: en,
  pl: en,
  cs: en,
  nl: en,
  ms: en,
  he: en,
  hi: en,
  'zh-TW': en,
  vi: en,
})

export type TemplateStringKey = keyof typeof en

export const translateTemplates = createI18n(templateStrings)
