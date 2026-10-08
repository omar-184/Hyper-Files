/// Hard cap on journaled cell edits (and recalculated formula cached
/// values) carried inline in one save request. Bulk actions — pasting,
/// sorting, or moving a large range — journal every affected cell, so this
/// must accommodate whole-sheet-scale edit sets. Shared by the zod save
/// schema and the preload's hand-rolled validator (which stays zod-free).
export const MAX_SAVE_EDITS = 1_000_000

/// Above this, the renderer stages edits through the chunked transfer
/// instead of the inline request: live object graphs cross the context
/// bridge property by property (~38s for a million edits measured on an M-
/// series laptop), while the transfer's JSON-string chunks cross as flat
/// strings in milliseconds.
export const SAVE_EDITS_INLINE_MAX = 10_000

/// Chunked save-edit transfer: the renderer uploads the edits as
/// JSON-serialized slices of SAVE_EDITS_CHUNK_MAX edits (a flat string per
/// IPC message instead of a structured-clone spike), then sends the save
/// request with an editsTransferId instead of an inline edits array. The
/// absolute ceiling is MAX_SAVE_EDITS_TOTAL — roughly "one bulk action over
/// an entire very large sheet"; past that the bottleneck is main-process
/// memory, not transport.
export const SAVE_EDITS_CHUNK_MAX = 100_000
export const MAX_SAVE_EDITS_TOTAL = 10_000_000
/// Byte-scale sanity cap on one serialized chunk (typical 100k-edit chunks
/// serialize to ~10–20MB; styles and rich runs can grow them further).
export const SAVE_EDITS_CHUNK_JSON_MAX = 256_000_000

/// Upper bound on one CSV export's serialized text (shared by the zod schema
/// and the preload validator, like MAX_SAVE_EDITS above).
export const MAX_CSV_EXPORT_CHARS = 64_000_000

/// One PDF-export header/footer template (shared by the zod schema and the
/// preload validator). `&G` pictures ride along as base64 data URLs — the
/// sidecar skips pictures over 2 MiB, so three slots fit under this cap.
export const MAX_PDF_TEMPLATE_CHARS = 12_000_000

/// VML shape id of a header/footer picture slot: left/center/right ×
/// header/footer, with an EVEN or FIRST suffix for the page variants.
export const HEADER_FOOTER_PICTURE_POSITION = /^[LCR][HF](EVEN|FIRST)?$/

export const IPC_CHANNELS = {
  selectWorkbook: 'workbook:select',
  /** Multi-file picker + sidecar sessions for merging into the current workbook */
  selectWorkbooksForMerge: 'workbook:select-for-merge',
  readWorkbookRange: 'workbook:read-range',
  readWorkbookFormulas: 'workbook:read-formulas',
  /// The sidecar process died; every session id the renderer holds is gone.
  /// A positive crash signal, so recovery never has to infer a crash from a
  /// session guard's error text (the Save swap and closeWorkbook reject ids
  /// with that same text on purpose).
  sidecarCrashed: 'workbook:sidecar-crashed',
  /// Re-open an already-known path through the normal open path (the same
  /// pipeline selectWorkbook runs), for recovering from a sidecar crash.
  reopenWorkbook: 'workbook:reopen',
  recalcWorkbook: 'workbook:recalc',
  readWorkbookMedia: 'workbook:read-media',
  readPivotDefinition: 'workbook:read-pivot-definition',
  readLocalImage: 'shell:read-local-image',
  closeWorkbook: 'workbook:close',
  saveWorkbook: 'workbook:save',
  /** Chunked upload of a large save's cell edits, consumed by the next save */
  saveEditsBegin: 'workbook:save-edits-begin',
  saveEditsChunk: 'workbook:save-edits-chunk',
  saveEditsAbort: 'workbook:save-edits-abort',
  /** Crash-recovery copy of a dirty workbook, written under userData */
  writeWorkbookRecovery: 'workbook:write-recovery',
  /** Main found a newer recovery copy while opening; renderer shows the styled prompt */
  recoveryPrompt: 'workbook:recovery-prompt',
  recoveryPromptReply: 'workbook:recovery-prompt-reply',
  workbookRenamed: 'workbook:renamed',
  pendingEditsChanged: 'workbook:pending-edits',
  closeSaveRequest: 'workbook:close-save-request',
  closeSaveResult: 'workbook:close-save-result',
  exportPdf: 'workbook:export-pdf',
  printWorkbook: 'workbook:print',
  exportCsv: 'workbook:export-csv',
  csvSaveConfirm: 'workbook:csv-save-confirm',
  openExternal: 'shell:open-external',
  menuAction: 'menu:action',
  captureScreenSources: 'sheets:capture-screen-sources',
  captureScreenSource: 'sheets:capture-screen-source',
} as const
