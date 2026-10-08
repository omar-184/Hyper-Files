/// Shell side of the Old Office Formats add-on: finds LibreOffice, asks before
/// writing a converted copy, runs the conversion and hands the result to the
/// normal open route. See converter.ts for why LibreOffice and how it runs.
import { constants, existsSync } from 'node:fs'
import { access } from 'node:fs/promises'
import { basename, dirname, extname, join } from 'node:path'
import { app, BrowserWindow, dialog, shell, type MenuItemConstructorOptions } from 'electron'
import type { Lang } from '@genoffice/i18n'
import { readAppSettings, writeAppSettings } from '../app-settings'
import {
  convertLegacyFile,
  defaultSofficeCandidates,
  findSoffice,
  LegacyConvertError,
  legacyTargetFor,
  siblingTargetPath,
  type LegacyTarget,
} from './converter'
import { legacyT, type LegacyStringKey } from './strings'

export { LEGACY_EXTENSIONS, LEGACY_RE } from './converter'

const GET_LIBREOFFICE_URL = 'https://www.libreoffice.org/download/download-libreoffice/'
/** app-settings.json keys */
const SETTING_PATH = 'legacyOfficeSofficePath'
const SETTING_NO_PROMPT = 'legacyOfficeConvertWithoutAsking'

export interface LegacyOfficeDeps {
  settingsPath: () => string
  lang: () => Lang
  /** the window dialogs attach to; undefined = app-modal */
  parentWindow: () => BrowserWindow | null
  /** opens a converted .docx/.xlsx/.pptx through the shell router */
  openConverted: (path: string) => void
}

let deps: LegacyOfficeDeps | null = null

export function configureLegacyOffice(next: LegacyOfficeDeps): void {
  deps = next
}

function requireDeps(): LegacyOfficeDeps {
  if (!deps) throw new Error('legacy office add-on used before configureLegacyOffice')
  return deps
}

function t(key: LegacyStringKey, params?: Record<string, string | number>): string {
  return legacyT(requireDeps().lang(), key, params)
}

/** app-owned folder a portable LibreOffice can be unpacked into */
function addonDir(): string {
  return join(app.getPath('userData'), 'addons', 'libreoffice')
}

function chosenPath(): string | undefined {
  const value = readAppSettings(requireDeps().settingsPath())[SETTING_PATH]
  return typeof value === 'string' && value ? value : undefined
}

let cachedSoffice: { path: string | undefined } | null = null

/** The LibreOffice program, or undefined when the add-on is not installed. */
export function locateSoffice(refresh = false): string | undefined {
  if (process.env.HYPERFILES_SOFFICE !== undefined) {
    return process.env.HYPERFILES_SOFFICE || undefined
  }
  // a hit is cached for the session; a miss is re-checked on every open so a
  // LibreOffice installed while the app runs is picked up without a restart
  if (!refresh && cachedSoffice?.path && existsSync(cachedSoffice.path)) return cachedSoffice.path
  const path = findSoffice({
    chosenPath: chosenPath(),
    candidates: defaultSofficeCandidates(process.env, addonDir()),
  })
  cachedSoffice = { path }
  return path
}

const KIND_KEYS: Record<string, LegacyStringKey> = {
  doc: 'kindWord',
  xls: 'kindExcel',
  ppt: 'kindPowerPoint',
  rtf: 'kindRtf',
  odt: 'kindOdt',
  ods: 'kindOds',
  odp: 'kindOdp',
}

function kindOf(path: string): string {
  const key = KIND_KEYS[extname(path).slice(1).toLowerCase()]
  return key ? t(key) : extname(path)
}

async function showBox(
  options: Electron.MessageBoxOptions,
): Promise<Electron.MessageBoxReturnValue> {
  const parent = requireDeps().parentWindow()
  return parent ? dialog.showMessageBox(parent, options) : dialog.showMessageBox(options)
}

/**
 * Route hook for the shell: true when this path is the add-on's to open.
 * `.xls` is the add-on's only when LibreOffice is present; otherwise the
 * sheets app keeps its built-in (values and formulas only) import.
 */
export function shouldOpenAsLegacy(path: string): boolean {
  const target = legacyTargetFor(path)
  if (!target) return false
  if (extname(path).toLowerCase() === '.xls') return locateSoffice() !== undefined
  return true
}

/** Convert and open; every outcome (missing add-on, cancel, failure) is handled here. */
export async function openLegacyDocument(source: string): Promise<void> {
  const target = legacyTargetFor(source)
  if (!target) return
  let soffice = locateSoffice()
  if (!soffice) {
    soffice = await promptForAddon(source, target)
    if (!soffice) return
  }
  const destination = await chooseDestination(source, target)
  if (!destination) return
  const parent = requireDeps().parentWindow()
  // the taskbar shows activity: a cold LibreOffice start takes a few seconds
  parent?.setProgressBar(2, { mode: 'indeterminate' })
  try {
    await convertLegacyFile({
      soffice,
      source,
      destination,
      profileDir: join(app.getPath('userData'), 'addons', 'libreoffice-profile'),
      tempRoot: join(app.getPath('temp'), 'hyper-files-legacy'),
    })
  } catch (error) {
    console.warn('[legacy-office] conversion failed:', error)
    const timedOut = error instanceof LegacyConvertError && error.reason === 'timeout'
    await showBox({
      type: 'error',
      message: t('failedMessage', { name: basename(source) }),
      detail: t(timedOut ? 'failedTimeout' : 'failedGeneric'),
      buttons: [t('btnClose')],
    })
    return
  } finally {
    if (parent && !parent.isDestroyed()) parent.setProgressBar(-1)
  }
  requireDeps().openConverted(destination)
}

async function promptForAddon(source: string, target: LegacyTarget): Promise<string | undefined> {
  const { response } = await showBox({
    type: 'info',
    title: t('missingTitle'),
    message: t('missingMessage', { name: basename(source), kind: kindOf(source) }),
    detail: t('missingDetail', { target: `.${target}` }),
    buttons: [t('btnLocate'), t('btnGetAddon'), t('btnCancel')],
    defaultId: 0,
    cancelId: 2,
    noLink: true,
  })
  if (response === 0) return pickSoffice()
  if (response === 1) void shell.openExternal(GET_LIBREOFFICE_URL)
  return undefined
}

/** File picker for soffice; the choice is remembered in app-settings.json. */
async function pickSoffice(): Promise<string | undefined> {
  const parent = requireDeps().parentWindow()
  const options: Electron.OpenDialogOptions = {
    title: t('locateTitle'),
    properties: ['openFile'],
    filters: process.platform === 'win32' ? [{ name: t('locateFilter'), extensions: ['exe'] }] : [],
  }
  const result = parent
    ? await dialog.showOpenDialog(parent, options)
    : await dialog.showOpenDialog(options)
  const picked = result.filePaths[0]
  if (result.canceled || !picked) return undefined
  if (!/^soffice(\.exe|\.com)?$/i.test(basename(picked))) {
    await showBox({ type: 'warning', message: t('locateInvalid'), buttons: [t('btnClose')] })
    return undefined
  }
  writeAppSettings(requireDeps().settingsPath(), { [SETTING_PATH]: picked })
  cachedSoffice = { path: picked }
  return picked
}

/** Beside the original by default (asked once, unless "don't ask" is on). */
async function chooseDestination(
  source: string,
  target: LegacyTarget,
): Promise<string | undefined> {
  const sibling = siblingTargetPath(source)
  const settings = readAppSettings(requireDeps().settingsPath())
  if (settings[SETTING_NO_PROMPT] !== true) {
    const { response, checkboxChecked } = await showBox({
      type: 'question',
      title: t('convertTitle', { target: `.${target}` }),
      message: t('convertMessage', { name: basename(source), target: `.${target}` }),
      detail: t('convertDetail', {
        kind: kindOf(source),
        target: `.${target}`,
        target_name: basename(sibling),
      }),
      buttons: [t('btnConvert'), t('btnConvertElsewhere'), t('btnCancel')],
      defaultId: 0,
      cancelId: 2,
      noLink: true,
      checkboxLabel: t('dontAskAgain'),
    })
    if (response === 2) return undefined
    if (checkboxChecked) {
      try {
        writeAppSettings(requireDeps().settingsPath(), { [SETTING_NO_PROMPT]: true })
      } catch (error) {
        console.warn('[legacy-office] could not remember the choice:', error)
      }
    }
    if (response === 1) return askSavePath(sibling, target)
  }
  return (await folderWritable(dirname(source))) ? sibling : askSavePath(sibling, target)
}

async function folderWritable(dir: string): Promise<boolean> {
  try {
    await access(dir, constants.W_OK)
    return true
  } catch {
    return false
  }
}

async function askSavePath(suggested: string, target: LegacyTarget): Promise<string | undefined> {
  const parent = requireDeps().parentWindow()
  const defaultPath = (await folderWritable(dirname(suggested)))
    ? suggested
    : join(app.getPath('documents'), basename(suggested))
  const options: Electron.SaveDialogOptions = {
    title: t('saveTitle'),
    defaultPath,
    filters: [{ name: target.toUpperCase(), extensions: [target] }],
  }
  const result = parent
    ? await dialog.showSaveDialog(parent, options)
    : await dialog.showSaveDialog(options)
  if (result.canceled || !result.filePath) return undefined
  return extname(result.filePath).toLowerCase() === `.${target}`
    ? result.filePath
    : `${result.filePath}.${target}`
}

/** Help menu entry: shows whether the add-on is found and lets the user fix it. */
export function legacyOfficeMenuItem(): MenuItemConstructorOptions {
  return { label: t('menuAddon'), click: () => void showAddonStatus() }
}

async function showAddonStatus(): Promise<void> {
  const found = locateSoffice(true)
  const chosen = chosenPath()
  if (found) {
    const buttons = [t('btnClose'), t('btnLocate')]
    if (chosen) buttons.push(t('btnForget'))
    const { response } = await showBox({
      type: 'info',
      title: t('missingTitle'),
      message: t('statusFound'),
      detail: t('statusFoundDetail', { path: found }),
      buttons,
      defaultId: 0,
      cancelId: 0,
      noLink: true,
    })
    if (response === 1) await pickSoffice()
    if (response === 2) {
      writeAppSettings(requireDeps().settingsPath(), { [SETTING_PATH]: '' })
      cachedSoffice = null
    }
    return
  }
  const { response } = await showBox({
    type: 'info',
    title: t('missingTitle'),
    message: t('statusMissing'),
    detail: t('statusMissingDetail'),
    buttons: [t('btnLocate'), t('btnGetAddon'), t('btnClose')],
    defaultId: 0,
    cancelId: 2,
    noLink: true,
  })
  if (response === 0) await pickSoffice()
  if (response === 1) void shell.openExternal(GET_LIBREOFFICE_URL)
}
