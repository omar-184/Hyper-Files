import { rm } from 'node:fs/promises'
import { basename, dirname, join, resolve } from 'node:path'

const UUID = '[0-9a-f]{8}-[0-9a-f]{4}-[1-8][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}'
const SESSION_SNAPSHOT_NAME = new RegExp(`^${UUID}\\.xlsx$`, 'i')
const IMPORT_DIRECTORY_NAME = new RegExp(`^${UUID}$`, 'i')

export interface SessionTemporaryResources {
  readonly snapshotPath: string
  readonly importTempDir?: string | undefined
}

interface CleanupSessionResourcesOptions extends SessionTemporaryResources {
  readonly tempRoot: string
  readonly closeSidecar: () => Promise<unknown>
}

function isDirectOwnedChild(path: string, parent: string, namePattern: RegExp): boolean {
  const resolvedPath = resolve(path)
  return dirname(resolvedPath) === resolve(parent) && namePattern.test(basename(resolvedPath))
}

/**
 * Close the sidecar first, remove its snapshot second, then remove a converted
 * CSV/XLS directory. Every filesystem deletion is constrained to an app-owned
 * temp root and filename shape.
 */
export async function cleanupSessionResources(
  options: CleanupSessionResourcesOptions,
): Promise<void> {
  try {
    await options.closeSidecar()
  } catch {
    // Continue with owned temp cleanup even when the sidecar is already gone.
  }
  const snapshotRoot = join(options.tempRoot, 'genoffice-sheets-sessions')
  if (isDirectOwnedChild(options.snapshotPath, snapshotRoot, SESSION_SNAPSHOT_NAME)) {
    await rm(options.snapshotPath, { force: true }).catch(() => undefined)
  }
  if (options.importTempDir !== undefined) {
    await cleanupImportTempDirectory(options.tempRoot, options.importTempDir)
  }
}

export async function cleanupImportTempDirectory(
  tempRoot: string,
  importTempDir: string,
): Promise<void> {
  const importRoot = join(tempRoot, 'genoffice-imports')
  if (!isDirectOwnedChild(importTempDir, importRoot, IMPORT_DIRECTORY_NAME)) return
  await rm(importTempDir, { recursive: true, force: true }).catch(() => undefined)
}
