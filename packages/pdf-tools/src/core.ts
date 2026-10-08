import type * as MupdfModule from 'mupdf'

/** The MuPDF.js module. The host loads it (worker thread, test runner) and passes it in. */
export type Mupdf = typeof MupdfModule
export type PDFDocument = MupdfModule.PDFDocument
export type PDFPage = MupdfModule.PDFPage
export type PDFObject = MupdfModule.PDFObject

/** A file handed to a tool. */
export interface ToolInput {
  /** file name with extension, used to name outputs */
  name: string
  bytes: Uint8Array
  /** open password for an encrypted PDF */
  password?: string
}

/** A file a tool produced; the host picks the folder and resolves name clashes. */
export interface ToolOutput {
  name: string
  bytes: Uint8Array
}

export type PdfToolErrorCode =
  'password-required' | 'wrong-password' | 'not-pdf' | 'damaged' | 'bad-input' | 'unsupported'

export class PdfToolError extends Error {
  constructor(
    readonly code: PdfToolErrorCode,
    message: string,
    /** the input file the error is about, when there is one */
    readonly file?: string,
  ) {
    super(message)
    this.name = 'PdfToolError'
  }
}

/** Save options for ordinary outputs: drop unused objects, compress streams. */
export const SAVE_DEFAULT = 'garbage=compact,compress,encrypt=none'

const PDF_MAGIC = [0x25, 0x50, 0x44, 0x46] // %PDF

/** True when `%PDF` appears in the first KiB (some files carry a junk prefix). */
export function looksLikePdf(bytes: Uint8Array): boolean {
  const end = Math.min(bytes.length - PDF_MAGIC.length, 1024)
  for (let i = 0; i <= end; i++) {
    if (PDF_MAGIC.every((b, k) => bytes[i + k] === b)) return true
  }
  return false
}

/**
 * Open a PDF, authenticating with its password when it has one. MuPDF repairs
 * broken cross-reference tables while opening, so a damaged file that still
 * has readable objects comes back usable.
 */
export function openPdf(m: Mupdf, input: ToolInput): PDFDocument {
  if (!looksLikePdf(input.bytes)) {
    throw new PdfToolError('not-pdf', `${input.name} is not a PDF file`, input.name)
  }
  let doc: PDFDocument
  try {
    doc = new m.PDFDocument(input.bytes)
  } catch (err) {
    throw new PdfToolError(
      'damaged',
      `${input.name} is damaged and could not be read (${(err as Error).message})`,
      input.name,
    )
  }
  if (doc.needsPassword()) {
    if (!input.password) {
      doc.destroy()
      throw new PdfToolError('password-required', `${input.name} needs a password`, input.name)
    }
    if (doc.authenticatePassword(input.password) === 0) {
      doc.destroy()
      throw new PdfToolError('wrong-password', `wrong password for ${input.name}`, input.name)
    }
  }
  if (doc.countPages() < 1) {
    doc.destroy()
    throw new PdfToolError('damaged', `${input.name} has no readable pages`, input.name)
  }
  return doc
}

/**
 * Resolve an indirect reference. MuPDF.js returns a document-less Null for a
 * missing key, whose resolve() throws, so every lookup goes through here.
 */
export function deref(obj: PDFObject): PDFObject {
  return obj.isNull() ? obj : obj.resolve()
}

export function savePdf(doc: PDFDocument, options = SAVE_DEFAULT): Uint8Array {
  const buf = doc.saveToBuffer(options)
  // copy out of the wasm heap: the Buffer is freed with the document
  const bytes = buf.asUint8Array().slice()
  buf.destroy()
  return bytes
}

/** "Report.final.pdf" -> "Report.final" */
export function stem(name: string): string {
  const base = name.replace(/^.*[\\/]/, '')
  const dot = base.lastIndexOf('.')
  return dot > 0 ? base.slice(0, dot) : base
}

/** Run `fn` with an open document and always free it. */
export function withPdf<T>(m: Mupdf, input: ToolInput, fn: (doc: PDFDocument) => T): T {
  const doc = openPdf(m, input)
  try {
    return fn(doc)
  } finally {
    doc.destroy()
  }
}
