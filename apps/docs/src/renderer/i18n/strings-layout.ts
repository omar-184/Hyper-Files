import { defineStrings } from '@genoffice/i18n'

const en = {
  ribbonLineNumbers: 'Line Numbers',
  ribbonLineNumbersTip: 'Number the lines in the margin beside the text',
  ribbonLineNumbersNone: 'None',
  ribbonLineNumbersContinuous: 'Continuous',
  ribbonLineNumbersEachPage: 'Restart Each Page',
  ribbonLineNumbersEachSection: 'Restart Each Section',
  ribbonLineNumbersSuppress: 'Suppress for Current Paragraph',
  ribbonLineNumbersOptions: 'Line Numbering Options…',
  ribbonLineNumbersStartAt: 'Start at',
  ribbonLineNumbersFromText: 'From text',
  ribbonLineNumbersCountBy: 'Count by',
  ribbonLineNumbersNumbering: 'Numbering',
}

/**
 * New Layout tab controls fall back to English until each locale has reviewed
 * terminology (same approach as strings-table.ts).
 */
export const layoutStrings = defineStrings({
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
