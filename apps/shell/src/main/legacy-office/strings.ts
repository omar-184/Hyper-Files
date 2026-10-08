import { createI18n, LANGS, type Lang } from '@genoffice/i18n'

/** Legacy Office add-on strings. English everywhere until translations land. */
const en = {
  kindWord: 'Word 97-2003 document',
  kindExcel: 'Excel 97-2003 workbook',
  kindPowerPoint: 'PowerPoint 97-2003 presentation',
  kindRtf: 'Rich Text document',
  kindOdt: 'OpenDocument text',
  kindOds: 'OpenDocument spreadsheet',
  kindOdp: 'OpenDocument presentation',
  menuAddon: 'Old Office Formats Add-on…',
  missingTitle: 'Old Office formats add-on',
  missingMessage: '“{name}” is a {kind}. Opening it needs the free Old Office Formats add-on.',
  missingDetail:
    'The add-on is LibreOffice, used only to convert the file to {target}. It is not part of the Hypercube Office install, so the base app stays small. Install LibreOffice (from libreoffice.org or any offline copy), then open the file again. If it is installed somewhere unusual, choose “Locate LibreOffice…”.',
  btnLocate: 'Locate LibreOffice…',
  btnGetAddon: 'Get LibreOffice',
  btnCancel: 'Cancel',
  btnClose: 'Close',
  btnForget: 'Forget this location',
  locateTitle: 'Locate LibreOffice (soffice)',
  locateFilter: 'LibreOffice program',
  locateInvalid:
    'That file is not the LibreOffice program. Choose soffice.exe in LibreOffice’s “program” folder.',
  convertTitle: 'Convert to {target}',
  convertMessage: 'Convert “{name}” to {target}?',
  convertDetail:
    'Hypercube Office edits this {kind} as a {target} file. The converted copy is saved as “{target_name}” next to the original. The original file is not changed.',
  btnConvert: 'Convert and open',
  btnConvertElsewhere: 'Save copy elsewhere…',
  dontAskAgain: 'Always convert without asking',
  saveTitle: 'Save the converted copy',
  failedMessage: 'Could not convert “{name}”',
  failedTimeout:
    'LibreOffice did not finish in time. The file may be very large, damaged, or password protected.',
  failedGeneric: 'LibreOffice could not convert the file. It may be damaged or password protected.',
  statusFound: 'The Old Office Formats add-on is ready.',
  statusFoundDetail:
    'Using LibreOffice at:\n{path}\n\nHypercube Office opens .doc, .xls, .ppt, .rtf, .odt, .ods and .odp files by converting them to .docx, .xlsx or .pptx. LibreOffice runs only while a file is converting.',
  statusMissing: 'The Old Office Formats add-on is not installed.',
  statusMissingDetail:
    'Install LibreOffice to open .doc, .xls (with formatting), .ppt, .rtf, .odt, .ods and .odp files. Without it, .xls files still open with values and formulas only.',
}

export type LegacyStringKey = keyof typeof en

const dicts = Object.fromEntries(LANGS.map((lang) => [lang, en])) as Record<Lang, typeof en>

const translate = createI18n(dicts)

export function legacyT(
  lang: Lang,
  key: LegacyStringKey,
  params?: Record<string, string | number>,
): string {
  return translate(lang, key, params)
}
