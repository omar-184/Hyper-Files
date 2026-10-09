import { containsAscii } from '../shared/ascii-search'

/**
 * A certificate signature covers the exact bytes of the file, and every write path here
 * rewrites the whole file, so writing into a signed PDF silently invalidates its
 * signature. This module decides when to ask first.
 */

/**
 * Every signature value dictionary carries /ByteRange. Signers write it uncompressed,
 * because they patch its offsets in place after laying out the file, so a native byte
 * search finds it without parsing.
 */
export function hasSignatureMarker(bytes: Uint8Array): boolean {
  return containsAscii(bytes, '/ByteRange')
}

/** What wants to write into the open file */
export type SignedWriteKind = 'save' | 'autosave' | 'pageOp'

/** The user's answer to the warning */
export type SignedWriteChoice = 'copy' | 'anyway' | 'cancel'

/** proceed: write as asked; copy: write nothing, offer Save As; cancel: write nothing */
export type SignedWriteDecision = 'proceed' | 'copy' | 'cancel'

export async function decideSignedWrite(args: {
  /** the file, as last read by this view, carries a signature */
  signed: boolean
  /** the user already chose to overwrite this signed file in this view */
  acknowledged: boolean
  kind: SignedWriteKind
  ask: (kind: Exclude<SignedWriteKind, 'autosave'>) => Promise<SignedWriteChoice>
}): Promise<SignedWriteDecision> {
  const { signed, acknowledged, kind, ask } = args
  if (!signed || acknowledged) return 'proceed'
  // Autosave never asks (it would interrupt every 30 seconds) and never writes: the
  // signed file keeps its pending edits until the user saves explicitly
  if (kind === 'autosave') return 'cancel'
  const choice = await ask(kind)
  return choice === 'anyway' ? 'proceed' : choice
}
