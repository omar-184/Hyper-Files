import { type Mupdf, type ToolInput, type ToolOutput, savePdf, stem, withPdf } from './core'

/**
 * Burn annotations and form fields into the page content so they print the
 * same everywhere and can no longer be edited.
 */
export function flattenPdf(m: Mupdf, input: ToolInput): ToolOutput {
  return withPdf(m, input, (doc) => {
    doc.bake(true, true)
    return { name: `${stem(input.name)}-flattened.pdf`, bytes: savePdf(doc) }
  })
}

/**
 * Rewrite a damaged PDF. MuPDF rebuilds a broken cross-reference table while
 * opening; saving with garbage collection writes a clean file from what it found.
 */
export function repairPdf(m: Mupdf, input: ToolInput): ToolOutput & { repaired: boolean } {
  return withPdf(m, input, (doc) => ({
    name: `${stem(input.name)}-repaired.pdf`,
    bytes: savePdf(doc, 'garbage=compact,compress,clean,encrypt=none'),
    repaired: doc.wasRepaired(),
  }))
}

export interface Metadata {
  title?: string
  author?: string
  subject?: string
  keywords?: string
}

const META_KEYS = {
  title: 'info:Title',
  author: 'info:Author',
  subject: 'info:Subject',
  keywords: 'info:Keywords',
} as const

export function readMetadata(m: Mupdf, input: ToolInput): Metadata {
  return withPdf(m, input, (doc) => {
    const out: Metadata = {}
    for (const [k, key] of Object.entries(META_KEYS) as [keyof Metadata, string][]) {
      const v = doc.getMetaData(key)
      if (v) out[k] = v
    }
    return out
  })
}

/** Set the document properties; a field left undefined keeps its value, '' clears it. */
export function writeMetadata(m: Mupdf, input: ToolInput, meta: Metadata): ToolOutput {
  return withPdf(m, input, (doc) => {
    for (const [k, key] of Object.entries(META_KEYS) as [keyof Metadata, string][]) {
      const v = meta[k]
      if (v !== undefined) doc.setMetaData(key, v)
    }
    return { name: `${stem(input.name)}-properties.pdf`, bytes: savePdf(doc) }
  })
}
