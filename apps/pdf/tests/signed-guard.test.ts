import { describe, expect, it, vi } from 'vitest'
import { PDFDocument, PDFName, PDFNumber } from 'pdf-lib'
import {
  decideSignedWrite,
  hasSignatureMarker,
  type SignedWriteChoice,
} from '../src/main/signed-guard'

const enc = (s: string) => new TextEncoder().encode(s)

describe('hasSignatureMarker', () => {
  it('finds a signature value dictionary', () => {
    const sig = '<< /Type /Sig /Filter /Adobe.PPKLite /ByteRange [0 840 960 240] /Contents <00> >>'
    expect(hasSignatureMarker(enc(`%PDF-1.7\n5 0 obj\n${sig}\nendobj\n`))).toBe(true)
  })

  it('ignores ordinary files', () => {
    expect(hasSignatureMarker(enc('%PDF-1.7\n1 0 obj << /Type /Catalog >> endobj'))).toBe(false)
    expect(hasSignatureMarker(new Uint8Array())).toBe(false)
  })

  it('finds the marker in a file written by pdf-lib', async () => {
    const doc = await PDFDocument.create()
    doc.addPage()
    expect(hasSignatureMarker(await doc.save())).toBe(false)
    const range = doc.context.obj([0, 0, 0, 0].map((n) => PDFNumber.of(n)))
    const sig = doc.context.obj({ Type: 'Sig', ByteRange: range })
    doc.catalog.set(PDFName.of('TestSig'), doc.context.register(sig))
    // Signers never put the signature dictionary in an object stream (its offsets are
    // patched in place), so write it the same way here
    expect(hasSignatureMarker(await doc.save({ useObjectStreams: false }))).toBe(true)
  })
})

describe('decideSignedWrite', () => {
  const asker = (choice: SignedWriteChoice) => vi.fn(async () => choice)

  it('writes unsigned files without asking', async () => {
    const ask = asker('cancel')
    for (const kind of ['save', 'autosave', 'pageOp'] as const) {
      expect(await decideSignedWrite({ signed: false, acknowledged: false, kind, ask })).toBe(
        'proceed',
      )
    }
    expect(ask).not.toHaveBeenCalled()
  })

  it('does not ask again once the user chose to write anyway', async () => {
    const ask = asker('cancel')
    for (const kind of ['save', 'autosave', 'pageOp'] as const) {
      expect(await decideSignedWrite({ signed: true, acknowledged: true, kind, ask })).toBe(
        'proceed',
      )
    }
    expect(ask).not.toHaveBeenCalled()
  })

  it('never lets autosave write into a signed file, and never interrupts for it', async () => {
    const ask = asker('anyway')
    expect(
      await decideSignedWrite({ signed: true, acknowledged: false, kind: 'autosave', ask }),
    ).toBe('cancel')
    expect(ask).not.toHaveBeenCalled()
  })

  it('asks before a save or page change and follows the answer', async () => {
    const cases: [SignedWriteChoice, string][] = [
      ['anyway', 'proceed'],
      ['copy', 'copy'],
      ['cancel', 'cancel'],
    ]
    for (const kind of ['save', 'pageOp'] as const) {
      for (const [choice, decision] of cases) {
        const ask = asker(choice)
        expect(await decideSignedWrite({ signed: true, acknowledged: false, kind, ask })).toBe(
          decision,
        )
        expect(ask).toHaveBeenCalledWith(kind)
      }
    }
  })
})
