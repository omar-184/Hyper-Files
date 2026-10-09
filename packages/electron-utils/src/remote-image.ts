/// Downloader for web pictures the user brings in (pasted web content, a picture
/// address, a web image in an HTML document). Browser-like headers and a couple
/// of retries turn most transient CDN refusals into successful inserts.

import { fetchWithSsrfGuard, type FetchWithSsrfGuardOptions } from './safe-remote-url'

const RETRY_DELAYS_MS: readonly number[] = [500, 1500]

/** Cap for a single downloaded picture; a pasted or typed address can point anywhere. */
export const MAX_REMOTE_IMAGE_BYTES = 50 * 1024 * 1024

export class ResponseTooLargeError extends Error {
  constructor(public readonly maxBytes: number) {
    super(`response larger than ${Math.round(maxBytes / (1024 * 1024))} MB`)
  }
}

function declaredLength(resp: Response): number | undefined {
  const raw = resp.headers.get('content-length')
  if (raw === null) return undefined
  const n = Number(raw)
  return Number.isFinite(n) && n >= 0 ? n : undefined
}

/**
 * Body bytes of `resp`, capped at `maxBytes`. A Content-Length above the cap
 * is refused before reading; otherwise (chunked transfers carry none, and a
 * header can lie) the stream is counted as it arrives and cancelled the moment
 * the cap is passed, so an oversized body never sits in memory whole.
 */
export async function readBodyCapped(resp: Response, maxBytes: number): Promise<Uint8Array> {
  const declared = declaredLength(resp)
  if (declared !== undefined && declared > maxBytes) {
    await resp.body?.cancel().catch(() => {})
    throw new ResponseTooLargeError(maxBytes)
  }
  if (!resp.body) {
    const bytes = new Uint8Array(await resp.arrayBuffer())
    if (bytes.byteLength > maxBytes) throw new ResponseTooLargeError(maxBytes)
    return bytes
  }
  const reader = resp.body.getReader()
  const chunks: Uint8Array[] = []
  let total = 0
  for (;;) {
    const { done, value } = await reader.read()
    if (done) break
    total += value.byteLength
    if (total > maxBytes) {
      await reader.cancel().catch(() => {})
      throw new ResponseTooLargeError(maxBytes)
    }
    chunks.push(value)
  }
  if (chunks.length === 1) return chunks[0]!
  const out = new Uint8Array(total)
  let offset = 0
  for (const chunk of chunks) {
    out.set(chunk, offset)
    offset += chunk.byteLength
  }
  return out
}

export function remoteImageHeaders(): Record<string, string> {
  const headers: Record<string, string> = {
    'User-Agent': 'Mozilla/5.0',
    // Only advertise formats the insert pipelines can label correctly: callers
    // map non-png/gif responses to JPEG, so preferring avif/webp would invite
    // content-negotiating CDNs to send bytes that end up mislabeled.
    Accept: 'image/png,image/jpeg,image/gif,image/*;q=0.8,*/*;q=0.5',
  }
  return headers
}

/**
 * fetchWithSsrfGuard specialized for image downloads: browser-like headers
 * and retries on transient failures
 * (network errors, 403/408/429, 5xx). An SSRF-blocked URL still returns null
 * immediately — that outcome never changes on retry.
 */
export async function fetchRemoteImage(
  rawUrl: string,
  options: Pick<FetchWithSsrfGuardOptions, 'fetchImpl'> & {
    retryDelaysMs?: readonly number[]
  } = {},
): Promise<Response | null> {
  const { retryDelaysMs = RETRY_DELAYS_MS, ...guardOptions } = options
  // only network images: local file URLs are never read through this path
  if (rawUrl.startsWith('file:')) return null
  const headers = remoteImageHeaders()
  for (let attempt = 0; ; attempt++) {
    let resp: Response | null = null
    let threw = false
    try {
      resp = await fetchWithSsrfGuard(rawUrl, { ...guardOptions, headers })
    } catch {
      threw = true
    }
    if (resp?.ok) return resp
    if (resp === null && !threw) return null // blocked by the SSRF guard: permanent
    const transient =
      threw ||
      (resp !== null &&
        (resp.status === 403 || resp.status === 408 || resp.status === 429 || resp.status >= 500))
    const delay = retryDelaysMs[attempt]
    if (!transient || delay === undefined) return resp
    await new Promise((r) => setTimeout(r, delay))
  }
}
