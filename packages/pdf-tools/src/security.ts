import {
  type Mupdf,
  PdfToolError,
  type ToolInput,
  type ToolOutput,
  savePdf,
  stem,
  withPdf,
} from './core'

/** What a reader may do without the owner password. */
export interface Permissions {
  print: boolean
  copy: boolean
  edit: boolean
  annotate: boolean
  fillForms: boolean
  assemble: boolean
}

export const ALL_PERMISSIONS: Permissions = {
  print: true,
  copy: true,
  edit: true,
  annotate: true,
  fillForms: true,
  assemble: true,
}

/**
 * The /P value (PDF 32000-1 table 22). Bits 7-8 and 13-32 must be 1; bits
 * 1-2 must be 0. Accessibility extraction (bit 10) is always allowed, and
 * high-quality print (bit 12) follows print.
 */
export function permissionBits(p: Permissions): number {
  let bits = 0xfffff0c0 | (1 << 9)
  if (p.print) bits |= (1 << 2) | (1 << 11)
  if (p.edit) bits |= 1 << 3
  if (p.copy) bits |= 1 << 4
  if (p.annotate) bits |= 1 << 5
  if (p.fillForms) bits |= 1 << 8
  if (p.assemble) bits |= 1 << 10
  // as a signed 32-bit integer, which is how /P is written
  return bits | 0
}

/** Option values cannot carry a comma: the save-options string splits on it. */
function checkPassword(pw: string, label: string) {
  if (pw.includes(',')) {
    throw new PdfToolError('bad-input', `the ${label} password cannot contain a comma`)
  }
}

export interface ProtectOptions {
  /** needed to open the file; empty means anyone can open it */
  userPassword: string
  /** needed to change permissions; defaults to the user password */
  ownerPassword?: string
  permissions?: Permissions
}

/** Encrypt with AES-256. */
export function protectPdf(m: Mupdf, input: ToolInput, opts: ProtectOptions): ToolOutput {
  const user = opts.userPassword
  const owner = opts.ownerPassword || user
  if (!user && !owner) throw new PdfToolError('bad-input', 'enter a password')
  checkPassword(user, 'open')
  checkPassword(owner, 'permissions')
  const perms = permissionBits(opts.permissions ?? ALL_PERMISSIONS)
  return withPdf(m, input, (doc) => ({
    name: `${stem(input.name)}-protected.pdf`,
    bytes: savePdf(
      doc,
      `garbage=compact,compress,encrypt=aes-256,user-password=${user},owner-password=${owner},permissions=${perms}`,
    ),
  }))
}

/** Remove encryption; the input's password must already be set on `input`. */
export function unlockPdf(m: Mupdf, input: ToolInput): ToolOutput {
  return withPdf(m, input, (doc) => ({
    name: `${stem(input.name)}-unlocked.pdf`,
    bytes: savePdf(doc),
  }))
}
