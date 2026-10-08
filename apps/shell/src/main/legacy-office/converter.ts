/// Legacy Office add-on: converts Word/Excel/PowerPoint 97-2003 files (and
/// the OpenDocument/RTF siblings LibreOffice reads) to their OOXML twins so
/// the regular editors can open them.
///
/// The converter is LibreOffice (MPL-2.0), run as a separate headless
/// process. It is never bundled with Hyper-Files: the add-on is "a
/// LibreOffice install (or portable copy) the app can find". Nothing here
/// touches the network, and the process only lives for one conversion, so a
/// 4 GB machine pays for it only while a file is being converted.
///
/// This module has no Electron imports so it can be unit-tested directly.
import { spawn, spawnSync } from 'node:child_process'
import { existsSync } from 'node:fs'
import { copyFile, mkdir, mkdtemp, readdir, rename, rm, stat } from 'node:fs/promises'
import { basename, dirname, extname, join } from 'node:path'
import { pathToFileURL } from 'node:url'

export type LegacyTarget = 'docx' | 'xlsx' | 'pptx'

/** Every extension the add-on converts, and the editor format it becomes. */
export const LEGACY_FORMATS: Readonly<Record<string, LegacyTarget>> = {
  doc: 'docx',
  rtf: 'docx',
  odt: 'docx',
  xls: 'xlsx',
  ods: 'xlsx',
  ppt: 'pptx',
  odp: 'pptx',
}

export const LEGACY_EXTENSIONS: readonly string[] = Object.keys(LEGACY_FORMATS)

/** Matches any path the add-on converts. */
export const LEGACY_RE = new RegExp(`\\.(${LEGACY_EXTENSIONS.join('|')})$`, 'i')

/** LibreOffice export filter per target; naming it pins the exact OOXML flavor. */
const EXPORT_FILTERS: Readonly<Record<LegacyTarget, string>> = {
  docx: 'docx:MS Word 2007 XML',
  xlsx: 'xlsx:Calc MS Excel 2007 XML',
  pptx: 'pptx:Impress MS PowerPoint 2007 XML',
}

export function legacyTargetFor(path: string): LegacyTarget | undefined {
  return LEGACY_FORMATS[extname(path).slice(1).toLowerCase()]
}

const SOFFICE_EXE = process.platform === 'win32' ? 'soffice.exe' : 'soffice'

/**
 * Where a LibreOffice install usually lives. `addonDir` is the app-owned
 * folder a portable LibreOffice can be unpacked into; it is checked first so
 * a copy made for Hyper-Files wins over a system install.
 */
export function defaultSofficeCandidates(
  env: NodeJS.ProcessEnv,
  addonDir: string,
  platform: NodeJS.Platform = process.platform,
): string[] {
  const exe = platform === 'win32' ? 'soffice.exe' : 'soffice'
  const candidates = [
    join(addonDir, 'program', exe),
    join(addonDir, 'LibreOffice', 'program', exe),
    join(addonDir, 'App', 'libreoffice', 'program', exe),
  ]
  if (platform === 'win32') {
    for (const root of [env.ProgramFiles, env['ProgramFiles(x86)'], env.ProgramW6432]) {
      if (root) candidates.push(join(root, 'LibreOffice', 'program', exe))
    }
    if (env.LOCALAPPDATA) {
      candidates.push(join(env.LOCALAPPDATA, 'Programs', 'LibreOffice', 'program', exe))
    }
  } else if (platform === 'darwin') {
    candidates.push('/Applications/LibreOffice.app/Contents/MacOS/soffice')
  } else {
    candidates.push(
      '/usr/bin/soffice',
      '/usr/local/bin/soffice',
      '/usr/lib/libreoffice/program/soffice',
      '/opt/libreoffice/program/soffice',
    )
  }
  return candidates
}

/** The install path LibreOffice's Windows installer records in the registry. */
function registrySoffice(): string | undefined {
  if (process.platform !== 'win32') return undefined
  for (const key of [
    'HKLM\\SOFTWARE\\LibreOffice\\UNO\\InstallPath',
    'HKCU\\SOFTWARE\\LibreOffice\\UNO\\InstallPath',
  ]) {
    const result = spawnSync('reg', ['query', key, '/ve'], {
      encoding: 'utf8',
      windowsHide: true,
      timeout: 5000,
    })
    const match = /REG_SZ\s+(.+)$/m.exec(result.stdout ?? '')
    if (match?.[1]) {
      const exe = join(match[1].trim(), SOFFICE_EXE)
      if (existsSync(exe)) return exe
    }
  }
  return undefined
}

/**
 * The LibreOffice executable to use: the path the user picked, then the
 * known locations, then (Windows) the registry. undefined = add-on missing.
 */
export function findSoffice(options: {
  chosenPath?: string | undefined
  candidates: readonly string[]
  useRegistry?: boolean
}): string | undefined {
  if (options.chosenPath && existsSync(options.chosenPath)) return options.chosenPath
  for (const candidate of options.candidates) if (existsSync(candidate)) return candidate
  return options.useRegistry === false ? undefined : registrySoffice()
}

export class LegacyConvertError extends Error {
  constructor(
    message: string,
    readonly reason: 'timeout' | 'failed' | 'no-output',
    readonly detail = '',
  ) {
    super(message)
  }
}

export interface ConvertOptions {
  soffice: string
  source: string
  /** where the converted file must end up (its extension is the target format) */
  destination: string
  /** a private LibreOffice profile, so a LibreOffice window the user has open
   *  never captures the job and the job never touches their settings */
  profileDir: string
  /** scratch root; a per-job directory is created and removed under it */
  tempRoot: string
  timeoutMs?: number
}

const DEFAULT_TIMEOUT_MS = 120_000
const MAX_STDERR = 4096

/** One conversion at a time: each soffice run is a few hundred MB on its own. */
let queue: Promise<unknown> = Promise.resolve()

export function convertLegacyFile(options: ConvertOptions): Promise<string> {
  const job = queue.then(() => runConversion(options))
  queue = job.catch(() => undefined)
  return job
}

async function runConversion(options: ConvertOptions): Promise<string> {
  const target = extname(options.destination).slice(1).toLowerCase() as LegacyTarget
  const filter = EXPORT_FILTERS[target]
  if (!filter) throw new LegacyConvertError(`unsupported target .${target}`, 'failed')
  await mkdir(options.tempRoot, { recursive: true })
  const workDir = await mkdtemp(join(options.tempRoot, 'job-'))
  try {
    // soffice names its output after the input's stem; a fixed ASCII copy of
    // the input keeps odd characters in the user's file name out of argv and
    // out of the output-name guess
    const inputExt = extname(options.source).toLowerCase()
    const input = join(workDir, `source${inputExt}`)
    await copyFile(options.source, input)
    const outDir = join(workDir, 'out')
    await mkdir(outDir)
    await runSoffice(options, [
      `-env:UserInstallation=${pathToFileURL(options.profileDir).href}`,
      '--headless',
      '--invisible',
      '--norestore',
      '--nolockcheck',
      '--nodefault',
      '--nologo',
      '--convert-to',
      filter,
      '--outdir',
      outDir,
      input,
    ])
    const produced = (await readdir(outDir)).find((name) =>
      name.toLowerCase().endsWith(`.${target}`),
    )
    if (!produced) throw new LegacyConvertError('LibreOffice produced no file', 'no-output')
    const output = join(outDir, produced)
    if ((await stat(output)).size === 0) {
      throw new LegacyConvertError('LibreOffice produced an empty file', 'no-output')
    }
    await moveInto(output, options.destination)
    return options.destination
  } finally {
    await rm(workDir, { recursive: true, force: true }).catch(() => undefined)
  }
}

function runSoffice(options: ConvertOptions, args: string[]): Promise<void> {
  return new Promise((resolve, reject) => {
    const child = spawn(options.soffice, args, {
      windowsHide: true,
      stdio: ['ignore', 'ignore', 'pipe'],
      cwd: dirname(options.soffice),
    })
    let stderr = ''
    child.stderr.on('data', (chunk: Buffer) => {
      if (stderr.length < MAX_STDERR) stderr += chunk.toString()
    })
    const timer = setTimeout(() => {
      child.kill()
      reject(new LegacyConvertError('LibreOffice took too long', 'timeout', stderr))
    }, options.timeoutMs ?? DEFAULT_TIMEOUT_MS)
    child.on('error', (error) => {
      clearTimeout(timer)
      reject(new LegacyConvertError(error.message, 'failed', stderr))
    })
    child.on('exit', (code) => {
      clearTimeout(timer)
      if (code === 0) resolve()
      else reject(new LegacyConvertError(`LibreOffice exited with code ${code}`, 'failed', stderr))
    })
  })
}

/** rename when on one volume; copy + delete when the temp dir is elsewhere */
async function moveInto(from: string, to: string): Promise<void> {
  try {
    await rename(from, to)
  } catch {
    await copyFile(from, to)
  }
}

/**
 * The converted file's default home: beside the original, named after it.
 * An existing file is never overwritten; "Report.docx" taken means
 * "Report (converted).docx", then "Report (converted 2).docx", and so on.
 */
export function siblingTargetPath(
  source: string,
  exists: (path: string) => boolean = existsSync,
): string {
  const target = legacyTargetFor(source)
  if (!target) throw new Error(`not a legacy file: ${source}`)
  const dir = dirname(source)
  const stem = basename(source, extname(source))
  const first = join(dir, `${stem}.${target}`)
  if (!exists(first)) return first
  for (let i = 1; ; i++) {
    const candidate = join(dir, `${stem} (converted${i === 1 ? '' : ` ${i}`}).${target}`)
    if (!exists(candidate)) return candidate
  }
}
