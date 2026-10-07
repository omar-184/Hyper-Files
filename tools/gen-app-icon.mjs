/**
 * Generates the Hyper-Files app icon set from apps/shell/brand/app-icon.svg:
 * build/icon.png, build/icon-mac.png, build/icons/<n>x<n>.png, build/icon.ico
 * and build/icon.icns for the shell, plus the docs app's standalone copies.
 *
 * The SVG is rasterized with Playwright's Chromium (transparent background).
 * ICO and ICNS are written as PNG-entry containers by hand, so the script runs
 * on any host.
 *
 * Regenerate after changing app-icon.svg:
 *   node tools/gen-app-icon.mjs
 */

import { mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { chromium } from 'playwright'

const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const svgPath = join(root, 'apps/shell/brand/app-icon.svg')
const OUT_DIRS = [join(root, 'apps/shell/build'), join(root, 'apps/docs/build')]

const SET_SIZES = [16, 32, 48, 64, 128, 256, 512, 1024]
const WIN_SIZES = [16, 24, 32, 48, 64, 128, 256]
// ICNS PNG entry types by pixel size
const ICNS_TYPES = [
  ['icp4', 16],
  ['icp5', 32],
  ['icp6', 64],
  ['ic07', 128],
  ['ic08', 256],
  ['ic09', 512],
  ['ic10', 1024],
]

/** ICO container with PNG-compressed entries (supported since Vista). */
function buildIco(entries) {
  const header = Buffer.alloc(6)
  header.writeUInt16LE(0, 0)
  header.writeUInt16LE(1, 2) // type: icon
  header.writeUInt16LE(entries.length, 4)
  const dir = Buffer.alloc(16 * entries.length)
  let offset = header.length + dir.length
  entries.forEach(({ size, png }, i) => {
    const o = i * 16
    dir.writeUInt8(size >= 256 ? 0 : size, o) // 0 means 256
    dir.writeUInt8(size >= 256 ? 0 : size, o + 1)
    dir.writeUInt8(0, o + 2) // palette
    dir.writeUInt8(0, o + 3) // reserved
    dir.writeUInt16LE(1, o + 4) // planes
    dir.writeUInt16LE(32, o + 6) // bpp
    dir.writeUInt32LE(png.length, o + 8)
    dir.writeUInt32LE(offset, o + 12)
    offset += png.length
  })
  return Buffer.concat([header, dir, ...entries.map((e) => e.png)])
}

/** ICNS container with PNG entries (macOS 10.7+). */
function buildIcns(entries) {
  const chunks = entries.map(({ type, png }) => {
    const head = Buffer.alloc(8)
    head.write(type, 0, 'ascii')
    head.writeUInt32BE(png.length + 8, 4)
    return Buffer.concat([head, png])
  })
  const head = Buffer.alloc(8)
  head.write('icns', 0, 'ascii')
  head.writeUInt32BE(8 + chunks.reduce((n, c) => n + c.length, 0), 4)
  return Buffer.concat([head, ...chunks])
}

async function renderPng(page, svgDataUrl, size) {
  await page.setViewportSize({ width: size, height: size })
  await page.setContent(
    `<body style="margin:0"><img src="${svgDataUrl}" style="display:block;width:${size}px;height:${size}px"></body>`,
  )
  return page.screenshot({ omitBackground: true })
}

const svg = readFileSync(svgPath)
const dataUrl = `data:image/svg+xml;base64,${svg.toString('base64')}`
// CHROMIUM_PATH points at a system Chromium when Playwright's own build is absent
const browser = await chromium.launch({
  headless: true,
  executablePath: process.env.CHROMIUM_PATH || undefined,
})
const page = await browser.newPage({ deviceScaleFactor: 1 })

try {
  const pngs = new Map()
  for (const size of new Set([...SET_SIZES, ...WIN_SIZES])) {
    pngs.set(size, await renderPng(page, dataUrl, size))
  }
  const ico = buildIco(WIN_SIZES.map((size) => ({ size, png: pngs.get(size) })))
  const icns = buildIcns(ICNS_TYPES.map(([type, size]) => ({ type, png: pngs.get(size) })))

  for (const outDir of OUT_DIRS) {
    writeFileSync(join(outDir, 'icon.png'), pngs.get(1024))
    writeFileSync(join(outDir, 'icon-mac.png'), pngs.get(1024))
    writeFileSync(join(outDir, 'icon.ico'), ico)
    writeFileSync(join(outDir, 'icon.icns'), icns)
  }
  const setDir = join(OUT_DIRS[0], 'icons')
  mkdirSync(setDir, { recursive: true })
  for (const size of SET_SIZES) {
    writeFileSync(join(setDir, `${size}x${size}.png`), pngs.get(size))
  }
  console.log('generated app icons')
} finally {
  await browser.close()
}
