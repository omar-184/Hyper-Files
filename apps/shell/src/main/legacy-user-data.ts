/**
 * The app was called Hyper-Files up to 0.1.x. Electron derives the userData
 * folder from productName, so after the rename an upgraded install would start
 * with empty settings, recents and starred files. On first launch the old
 * folder is moved to the new name; if that fails (e.g. a file is locked) the
 * app keeps using the old folder rather than starting blank.
 */
import { existsSync, renameSync } from 'node:fs'
import { join } from 'node:path'

export const LEGACY_PRODUCT_NAME = 'Hyper-Files'

/** the userData folder to use: `current`, after moving the legacy folder into place when needed */
export function resolveUserDataDir(appData: string, current: string): string {
  const legacy = join(appData, LEGACY_PRODUCT_NAME)
  if (existsSync(current) || !existsSync(legacy)) return current
  try {
    renameSync(legacy, current)
    return current
  } catch {
    return legacy
  }
}
