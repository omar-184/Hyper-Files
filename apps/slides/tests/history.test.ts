import { describe, expect, it, vi } from 'vitest'
import {
  addPicture,
  createBlankPptx,
  openPptx,
  replacePictureBytes,
  type PictureElement,
} from '@genoffice/pptx-engine'
import type { Session } from '../src/main/session-state'
import {
  beginHistoryBatch,
  endHistoryBatch,
  pushHistory,
  restoreSnapshot,
  settleStaleHistoryBatch,
  takeSnapshot,
} from '../src/main/session-state'

vi.mock('electron', () => ({
  BrowserWindow: { getFocusedWindow: () => null },
}))
vi.mock('../src/main/fonts', () => ({
  createSystemFontMetrics: () => ({}),
}))

function sessionWith(value: string): Session {
  return {
    path: '',
    fitWidthPx: 1280,
    undoStack: [],
    redoStack: [],
    opened: {
      deck: {
        slides: [{ value }],
        size: { cx: 1, cy: 1 },
      },
      archive: { entries: new Map([['deck', new Uint8Array([value.length])]]) },
    },
  } as unknown as Session
}

function valueOf(session: Session): string {
  return (session.opened.deck.slides[0] as unknown as { value: string }).value
}

function setValue(session: Session, value: string): void {
  ;(session.opened.deck.slides[0] as unknown as { value: string }).value = value
}

describe('Slides main-process history batching', () => {
  it('collapses several edits into one pre-run snapshot', () => {
    const session = sessionWith('before')
    beginHistoryBatch(session)
    pushHistory(session)
    setValue(session, 'first')
    pushHistory(session)
    setValue(session, 'second')
    endHistoryBatch(session)

    expect(session.undoStack).toHaveLength(1)
    restoreSnapshot(session, session.undoStack[0]!)
    expect(valueOf(session)).toBe('before')
  })

  it('supports nested batching inside an outer batch', () => {
    const session = sessionWith('before')
    beginHistoryBatch(session)
    beginHistoryBatch(session)
    pushHistory(session)
    setValue(session, 'after')
    endHistoryBatch(session)
    expect(session.historyBatch?.depth).toBe(1)
    endHistoryBatch(session)
    expect(session.undoStack).toHaveLength(1)
  })

  it('does not create a history step when every edit is rolled back', () => {
    const session = sessionWith('before')
    beginHistoryBatch(session)
    pushHistory(session)
    session.undoStack.pop()
    endHistoryBatch(session)
    expect(session.undoStack).toHaveLength(0)
  })

  it('restores the deck size on undo', () => {
    const session = sessionWith('before')
    pushHistory(session)
    session.opened.deck.size = { cx: 2, cy: 3 }
    setValue(session, 'after')

    restoreSnapshot(session, session.undoStack.pop()!)
    expect(session.opened.deck.size).toEqual({ cx: 1, cy: 1 })
    expect(valueOf(session)).toBe('before')
  })

  it('returns the pre-run snapshot from the outermost batch end with edits', () => {
    const session = sessionWith('before')
    beginHistoryBatch(session)
    beginHistoryBatch(session)
    pushHistory(session)
    setValue(session, 'after')
    expect(endHistoryBatch(session)).toBeNull()
    const before = endHistoryBatch(session)
    expect(before).not.toBeNull()
    expect((before!.slides[0] as unknown as { value: string }).value).toBe('before')

    const emptyRun = sessionWith('untouched')
    beginHistoryBatch(emptyRun)
    expect(endHistoryBatch(emptyRun)).toBeNull()
  })

  it('undo restores the archive-only dirty flag with the deck', () => {
    const session = sessionWith('before')
    pushHistory(session)
    setValue(session, 'notes edited')
    session.metaDirty = true

    session.redoStack.push(takeSnapshot(session))
    restoreSnapshot(session, session.undoStack.pop()!)
    expect(session.metaDirty).toBe(false)

    restoreSnapshot(session, session.redoStack.pop()!)
    expect(session.metaDirty).toBe(true)
  })

  it('undo → edit → redo replays the state that was undone, not a mutated copy', () => {
    const session = sessionWith('before')
    pushHistory(session)
    setValue(session, 'edited')

    // ⌘Z
    session.redoStack.push(takeSnapshot(session))
    restoreSnapshot(session, session.undoStack.pop()!)
    expect(valueOf(session)).toBe('before')

    // typing after the undo must not rewrite the redo snapshot in place
    setValue(session, 'typed after undo')
    restoreSnapshot(session, session.redoStack.pop()!)
    expect(valueOf(session)).toBe('edited')
  })

  it('a batch left open by a crashed tool path is collapsed so undo still works', () => {
    const session = sessionWith('before')
    // run begins a batch, a tool nests another, then the tool path dies without ending either
    beginHistoryBatch(session)
    pushHistory(session)
    setValue(session, 'edited')
    beginHistoryBatch(session)
    expect(session.historyBatch).toBeDefined()

    settleStaleHistoryBatch(session)
    expect(session.historyBatch).toBeUndefined()
    expect(session.undoStack.length).toBe(1)

    restoreSnapshot(session, session.undoStack.pop()!)
    expect(valueOf(session)).toBe('before')
  })

  it('undo and redo restore resource bytes and relationships pruned by picture replacement', async () => {
    const png = new Uint8Array(
      Buffer.from(
        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M/wHwAF/gL+XwV2AAAAAElFTkSuQmCC',
        'base64',
      ),
    )
    const gif = new Uint8Array(
      Buffer.from('R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7', 'base64'),
    )
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const picture = addPicture(opened, slide, {
      bytes: png,
      ext: 'png',
      offset: { x: 0, y: 0, cx: 914400, cy: 914400 },
    })!
    const session: Session = {
      path: '',
      opened,
      fitWidthPx: 1280,
      undoStack: [],
      redoStack: [],
    }
    const oldMedia = picture.mediaRef
    const slash = slide.path.lastIndexOf('/')
    const relsPath = `${slide.path.slice(0, slash)}/_rels/${slide.path.slice(slash + 1)}.rels`
    const oldRels = opened.archive.readText(relsPath)!

    // The IPC mutation path snapshots first, then replacement prunes the old
    // relationship/media as soon as the element points at the new bytes.
    pushHistory(session)
    expect(replacePictureBytes(opened, slide, picture.id, gif, 'gif')).toBe(true)
    const newMedia = picture.mediaRef
    const newRels = opened.archive.readText(relsPath)!
    expect(opened.archive.entries.has(oldMedia)).toBe(false)
    expect(opened.archive.readBytes(newMedia)).toEqual(gif)

    // Production undo stores the cleaned state for redo, then restores the
    // pre-mutation snapshot.
    session.redoStack.push(takeSnapshot(session))
    restoreSnapshot(session, session.undoStack.pop()!)
    expect(session.opened.archive.readBytes(oldMedia)).toEqual(png)
    expect(session.opened.archive.entries.has(newMedia)).toBe(false)
    expect(session.opened.archive.readText(relsPath)).toBe(oldRels)
    expect(
      (
        session.opened.deck.slides[0]!.elements.find(
          (element) => element.type === 'picture',
        ) as PictureElement
      ).mediaRef,
    ).toBe(oldMedia)

    // Production redo snapshots the restored state and reapplies the cleaned
    // archive/model snapshot.
    session.undoStack.push(takeSnapshot(session))
    restoreSnapshot(session, session.redoStack.pop()!)
    expect(session.opened.archive.entries.has(oldMedia)).toBe(false)
    expect(session.opened.archive.readBytes(newMedia)).toEqual(gif)
    expect(session.opened.archive.readText(relsPath)).toBe(newRels)
    expect(
      (
        session.opened.deck.slides[0]!.elements.find(
          (element) => element.type === 'picture',
        ) as PictureElement
      ).mediaRef,
    ).toBe(newMedia)
  })
})
