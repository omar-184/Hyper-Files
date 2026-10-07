export interface OpenFileResult {
  path: string
  name: string
  /** one-shot URL serving the docx bytes (fetch it exactly once) */
  dataUrl: string
  /** sha256 of the original file; original archived under this hash */
  hash: string
  /** the on-disk file is password protected (opened via decrypt; saves re-encrypt) */
  encrypted?: boolean
  /** content came from a newer crash-recovery copy and still needs an explicit save */
  recovered?: boolean
}

/** Password-protected (ECMA-376 encrypted) docx: the renderer prompts for the
 *  open password and retries via openDocxDecrypt. */
export interface OpenFileNeedsPassword {
  needsPassword: true
  path: string
  name: string
}

export type OpenDocxResult = OpenFileResult | OpenFileNeedsPassword | null

/** result of an openDocxDecrypt attempt; wrong-password keeps the prompt open */
export type DecryptOpenResult =
  | { ok: true; result: OpenFileResult }
  | { ok: false; reason: 'wrong-password' | 'unsupported' | 'error'; error?: string }

export interface PickImageResult {
  /** raw image bytes, base64 encoded */
  base64: string
  mime: 'image/png' | 'image/jpeg' | 'image/gif'
  name: string
}

import type { HeadlessExportTarget } from '@genoffice/electron-utils/headless-export'
import type { FaceVerticalMetrics } from '@genoffice/font-metrics'

export type { FaceVerticalMetrics }

/** an open docs tab, for View → Switch Tab */
export interface DocsTabInfo {
  id: string
  title: string
  focused: boolean
}

/** commands dispatched from the native application menu to the renderer */
export type MenuCommand =
  | 'new'
  | 'open'
  | 'open-path'
  | 'save'
  | 'save-as'
  | 'undo'
  | 'redo'
  | 'zoom-in'
  | 'zoom-out'
  | 'zoom-100'
  | 'zoom-set'
  | 'zoom-page-width'
  | 'zoom-whole-page'
  | 'toggle-dark'
  | 'insert-table'
  | 'insert-image'
  | 'insert-page-break'
  | 'insert-link'
  | 'insert-equation'
  | 'insert-comment'
  | 'font-dialog'
  | 'paragraph-dialog'
  | 'bold'
  | 'italic'
  | 'underline'
  | 'align-left'
  | 'align-center'
  | 'align-right'
  | 'align-justify'
  | 'page-setup'
  | 'find'
  | 'replace'
  | 'goto'
  | 'print'
  | 'export-pdf'
  | 'export-html'
  | 'export-images'
  | 'word-count'
  | 'autocorrect-options'
  | 'preferences'
  | 'table-insert-cells'
  | 'table-insert-rows-above'
  | 'table-insert-rows-below'
  | 'table-insert-cols-left'
  | 'table-insert-cols-right'
  | 'table-delete-table'
  | 'table-delete-columns'
  | 'table-delete-rows'
  | 'table-delete-cells'
  | 'table-select-table'
  | 'table-select-column'
  | 'table-select-row'
  | 'table-select-cell'
  | 'table-merge-cells'
  | 'table-split-cells'
  | 'table-split-table'
  | 'table-autofit-contents'
  | 'table-autofit-window'
  | 'table-autofit-fixed'
  | 'table-distribute-rows'
  | 'table-distribute-columns'
  | 'table-repeat-header'
  | 'table-gridlines'
  | 'table-properties'
  | 'shortcuts'

export type UiTheme = 'light' | 'dark' | 'system'

/**
 * Document page theme preference (#1811): what the canvas/paper does relative
 * to the UI theme. 'follow' keeps the previous single-theme behavior.
 */
export type DocTheme = 'follow' | 'light' | 'dark'

/** shell-wide AutoSave default; updatedAt is 0 until the user has ever set it */
export interface AutoSaveDefault {
  on: boolean
  updatedAt: number
}

export type ZoteroCommand =
  'addEditCitation' | 'addEditBibliography' | 'refresh' | 'setDocPrefs' | 'removeCodes'

export type ZoteroCommandErrorCode =
  'connection-refused' | 'unsupported-command' | 'operation-failed'

export interface ZoteroCommandResult {
  ok: boolean
  errorCode?: ZoteroCommandErrorCode
  error?: string
}

export interface ZoteroRendererRequest {
  requestId: string
  command: string
  args: unknown[]
}

export interface ZoteroRendererResponse {
  requestId: string
  ok: boolean
  result?: unknown
  error?: string
}

/** Chromium's misspelling data for a claimed body right-click (`seq` = the claim it answers) */
export interface ContextMenuRequest {
  seq: number
  misspelledWord: string
  suggestions: string[]
}

export interface SpellLanguages {
  active: string[]
  /** empty on macOS: the OS checker picks the language itself */
  available: string[]
}

export interface DesktopApi {
  /** current UI language (persisted by the shell in app-settings.json) */
  getLanguage(): Promise<'zh' | 'en' | 'ja' | 'ko' | 'fr' | 'de' | 'es' | 'th' | 'id' | 'ru' | 'ar'>
  /** OS regional-settings locale (BCP 47); Word derives the new-document paper size from it */
  getSystemLocale(): Promise<string>
  /** language switched from the shell home page */
  onLanguageChanged(
    handler: (
      lang: 'zh' | 'en' | 'ja' | 'ko' | 'fr' | 'de' | 'es' | 'th' | 'id' | 'ru' | 'ar',
    ) => void,
  ): () => void
  /** current UI theme preference (persisted by the shell in app-settings.json) */
  getTheme(): Promise<UiTheme>
  /** theme switched from the shell home page */
  onThemeChanged(handler: (theme: UiTheme) => void): () => void
  /** current document page theme preference (#1811, persisted by the shell in app-settings.json) */
  getDocumentTheme(): Promise<DocTheme>
  /** document page theme switched from the shell home page */
  onDocumentThemeChanged(handler: (theme: DocTheme) => void): () => void
  /** shell-wide AutoSave default (see useAutoSavePref) */
  getAutoSaveDefault(): Promise<AutoSaveDefault>
  onAutoSaveDefaultChanged(handler: (value: AutoSaveDefault) => void): () => void
  /** press on the shell chrome (tab strip is a sibling WebContentsView whose
   *  clicks produce no DOM event here) — dismiss open popovers */
  onChromePressed(handler: () => void): () => void
  /** invoke Zotero's word-processor integration and service its document callbacks */
  zoteroCommand(command: ZoteroCommand): Promise<ZoteroCommandResult>
  onZoteroRequest(handler: (request: ZoteroRendererRequest) => void): () => void
  respondToZotero(response: ZoteroRendererResponse): void
  openDocx(): Promise<OpenDocxResult>
  openDocxPath(path: string): Promise<OpenDocxResult>
  confirmDocumentReplace(): Promise<boolean>
  /** decrypt-and-open a password-protected docx (path from a needsPassword result) */
  openDocxDecrypt(path: string, password: string): Promise<DecryptOpenResult>
  /** w:altChunk HTML rendered through html2docx in a hidden window; null when conversion fails */
  convertAltChunkHtml(html: string): Promise<Uint8Array | null>
  /** Review > Protect: set (or clear with null) the desired next-save password;
   *  filePath null = document not saved yet, applied on its first successful save */
  setDocPassword(filePath: string | null, password: string | null): Promise<{ ok: boolean }>
  /** snapshot the current intent sequence before replacement cleanup is queued */
  docPasswordIntentRevision(): Promise<number>
  /** discard prior-document intents through a captured revision */
  discardDocPasswordIntents(throughRevision: number): Promise<{ ok: boolean }>
  /** mark the renderer ready and consume a file passed by Finder/Explorer at launch */
  consumePendingOpenDocx(): Promise<OpenDocxResult>
  /** returns true when this tab was created via "New Document" and should start blank */
  consumeNewBlankDoc(): Promise<boolean>
  /** Headless export mode: the path and format this hidden renderer must export, null in normal use */
  consumeHeadlessExport(): Promise<HeadlessExportTarget | null>
  /** Headless export mode: report the export outcome so the main process can quit */
  headlessExportDone(result: { ok: boolean; error?: string }): void
  /** receive documents opened from Finder/Explorer while the app is running */
  onOpenDocx(handler: (result: Exclude<OpenDocxResult, null>) => void): () => void
  /** File was renamed externally (renamed in the shell Home list) — pushes old and new paths; renderer syncs its save path and title bar */
  onRenamedDocx(handler: (paths: { oldPath: string; newPath: string }) => void): () => void
  /** auto=true marks an autosave: an externally modified file then fails with
   *  reason 'external-modified' instead of prompting (manual saves get an
   *  Overwrite/Cancel dialog in the main process) */
  saveDocx(
    path: string,
    data: ArrayBuffer,
    auto?: boolean,
  ): Promise<{
    ok: boolean
    error?: string
    reason?: 'external-modified'
    /** a newer password choice arrived after this save's snapshot */
    passwordIntentPending?: boolean
    /** one-shot URL of the saved document in full when an encrypted save absorbed
     *  lazily served pictures: the renderer reparses from it and leaves lazy mode */
    dataUrl?: string
  }>
  /** crash-recovery copy of a dirty document, stored under userData */
  writeRecoveryCopy(path: string, data: ArrayBuffer): Promise<{ ok: boolean }>
  /** tab closed but webContents kept alive (shell freeze workaround) — stop background timers */
  onTeardown(handler: () => void): () => void
  /** one trusted space keystroke into this webContents — the only thing that
   *  makes Blink respell existing text after the spellcheck attribute turns
   *  back on (r168); the caller pauses the PM DOM observer and removes the
   *  space again by script */
  respellKick(): Promise<void>
  /** append one line to userData/spell-diag.log (size-capped) — field
   *  spellcheck failures are intermittent and platform-bound, so the
   *  toggle/kick lifecycle keeps a trace support can ask users for */
  spellDiag(line: string): void
  /** opt this renderer into claiming right-clicks: claimed clicks get no native menu */
  armContextMenu(): void
  /** synchronous, from the DOM contextmenu handler: the React menu answers this
   *  right-click, so Blink's request for it must not pop the native menu */
  claimContextMenu(seq: number): void
  /** Chromium's misspelling data for a claimed click */
  onContextMenuRequest(handler: (request: ContextMenuRequest) => void): () => void
  spellAddWord(word: string): Promise<boolean>
  /** Word's Ignore All: skipped while this document is open, forgotten when it closes */
  spellIgnoreWord(word: string): Promise<boolean>
  /** Blink-side replacement of the misspelled word under the last right-click */
  spellReplace(word: string): Promise<void>
  spellLanguages(): Promise<SpellLanguages>
  spellSetLanguages(langs: string[]): Promise<SpellLanguages>
  /** sourcePath: the document's current path — Save As uses its desired next-save
   *  password and commits that state to the chosen path only after success */
  saveDocxAs(
    defaultName: string,
    data: ArrayBuffer,
    sourcePath?: string | null,
  ): Promise<{
    ok: boolean
    path?: string
    error?: string
    passwordIntentPending?: boolean
    dataUrl?: string
  }>
  /** first save of a new document: silently writes into the default folder, no dialog */
  saveDocxNew(
    defaultName: string,
    data: ArrayBuffer,
  ): Promise<{
    ok: boolean
    path?: string
    error?: string
    passwordIntentPending?: boolean
    dataUrl?: string
  }>
  getRecentFiles(): Promise<string[]>
  pickImage(): Promise<PickImageResult | null>
  /** vertical metrics of an installed family (exact name match), null when missing */
  fontMetrics(family: string): Promise<FaceVerticalMetrics | null>
  /** system print dialog for the current window; ok=false without error = canceled.
   *  scale: print scale inverting the preview's print zoom (print-zoom.ts) */
  print(scale?: number): Promise<{ ok: boolean; error?: string }>
  /** render the document to PDF and ask where to save; size in twips.
   *  outPath is only honored when a previous export dialog chose that exact path */
  exportPdf(
    defaultName: string,
    pageWidthTwips: number,
    pageHeightTwips: number,
    outPath?: string,
    scale?: number,
  ): Promise<{ ok: boolean; path?: string; error?: string }>
  exportHtml(
    defaultName: string,
    html: string,
    outPath?: string,
  ): Promise<{ ok: boolean; path?: string; error?: string }>
  /** Mixed paper-size export: produce a set of PDF bytes (base64) at given sizes per the current print layout */
  printPdfBuffer(
    pageWidthTwips: number,
    pageHeightTwips: number,
    scale?: number,
  ): Promise<{ ok: boolean; base64?: string; error?: string }>
  /** Merge grouped PDF fragments in order and write to disk (missing outPath opens
   *  the save dialog; a given outPath must come from a previous export dialog) */
  saveMergedPdf(
    defaultName: string,
    base64Parts: string[],
    outPath?: string,
  ): Promise<{ ok: boolean; path?: string; error?: string }>
  /** Export as images: the folder picker plus a pre-authorized temp PDF path the
   *  regular PDF export writes to silently (no reveal, no open) */
  pickExportImagesTarget(): Promise<{ dir: string; pdfPath: string } | null>
  /** Read back and delete the temp PDF written for an image export */
  takeExportPdf(pdfPath: string): Promise<{ ok: boolean; base64?: string; error?: string }>
  /** Write one page PNG into the folder chosen by pickExportImagesTarget */
  writeExportImage(
    dir: string,
    fileName: string,
    pngBase64: string,
  ): Promise<{ ok: boolean; path?: string; error?: string }>
  /** Save a picture the renderer displays (data URL) through a Save dialog */
  saveImageAs(src: string): Promise<{ ok: boolean; path?: string; error?: string }>
  /** Native context menu "View Image" on a chrome surface */
  onViewImage(handler: (src: string) => void): () => void
  /** copy an embedded picture to the OS clipboard as a real bitmap + <img>
   *  html (r136: copying an image exported only the protected placeholder) */
  copyImageToClipboard(dataUrl: string, metaJson?: string): Promise<boolean>
  /** download a pasted web image (bitmap-less HTML paste) as base64; null on failure */
  fetchImage(url: string): Promise<{ base64: string; mime: string } | null>
  /** absolute path of a File dropped onto the window (Electron webUtils) */
  getPathForFile(file: File): string
  /** View → New Tab: open another docs tab, optionally loading the same document */
  openNewTab(openPath?: string | null): Promise<void>
  /** all open docs tabs, for View → Switch Tab */
  listDocsTabs(): Promise<DocsTabInfo[]>
  focusDocsTab(id: string): Promise<void>
  /** subscribe to native menu commands; returns unsubscribe */
  onMenuCommand(handler: (command: MenuCommand, payload?: string) => void): () => void
  /** Close guard: main process queries pre-close state (dirty flag + autosave switch; if autosave is on, save silently without a dialog) */
  onCloseCheck(handler: () => void): () => void
  reportCloseCheck(state: { dirty: boolean; autoSave: boolean; filePath?: string | null }): void
  /** Close guard chose "Save": main process asks the renderer to run the full save flow */
  onCloseSaveRequest(handler: () => void): () => void
  reportCloseSaveResult(ok: boolean): void
  /** keep the native View menu's checkbox items in sync with renderer state */
  reportViewMenuState(state: { darkCanvas: boolean }): void
}

/** mirrors VIEW_IMAGE_CHANNEL in @genoffice/electron-utils (kept literal so the preload stays free of main-only deps) */
export const VIEW_IMAGE_CHANNEL = 'genoffice:view-image'
