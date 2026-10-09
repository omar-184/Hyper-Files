import type { ToolId, ToolOptions } from '../../../shared/pdf-tools-api'
import type { ToolStringKey } from '../i18n/strings-tools'

export type ToolGroup = 'organize' | 'optimize' | 'convert' | 'security' | 'edit'

export interface ToolDef {
  id: ToolId
  group: ToolGroup
  name: ToolStringKey
  desc: ToolStringKey
  /** what the file list takes */
  input: 'pdf' | 'image'
  /** combining tools need at least this many files */
  minFiles: number
}

export const GROUPS: { id: ToolGroup; label: ToolStringKey }[] = [
  { id: 'organize', label: 'groupOrganize' },
  { id: 'optimize', label: 'groupOptimize' },
  { id: 'convert', label: 'groupConvert' },
  { id: 'security', label: 'groupSecurity' },
  { id: 'edit', label: 'groupEdit' },
]

export const TOOLS: ToolDef[] = [
  {
    id: 'merge',
    group: 'organize',
    name: 'toolMerge',
    desc: 'toolMergeDesc',
    input: 'pdf',
    minFiles: 2,
  },
  {
    id: 'split',
    group: 'organize',
    name: 'toolSplit',
    desc: 'toolSplitDesc',
    input: 'pdf',
    minFiles: 1,
  },
  {
    id: 'extract',
    group: 'organize',
    name: 'toolExtract',
    desc: 'toolExtractDesc',
    input: 'pdf',
    minFiles: 1,
  },
  {
    id: 'delete',
    group: 'organize',
    name: 'toolDelete',
    desc: 'toolDeleteDesc',
    input: 'pdf',
    minFiles: 1,
  },
  {
    id: 'rotate',
    group: 'organize',
    name: 'toolRotate',
    desc: 'toolRotateDesc',
    input: 'pdf',
    minFiles: 1,
  },
  {
    id: 'reorder',
    group: 'organize',
    name: 'toolReorder',
    desc: 'toolReorderDesc',
    input: 'pdf',
    minFiles: 1,
  },
  {
    id: 'compress',
    group: 'optimize',
    name: 'toolCompress',
    desc: 'toolCompressDesc',
    input: 'pdf',
    minFiles: 1,
  },
  {
    id: 'repair',
    group: 'optimize',
    name: 'toolRepair',
    desc: 'toolRepairDesc',
    input: 'pdf',
    minFiles: 1,
  },
  {
    id: 'flatten',
    group: 'optimize',
    name: 'toolFlatten',
    desc: 'toolFlattenDesc',
    input: 'pdf',
    minFiles: 1,
  },
  {
    id: 'ocr',
    group: 'convert',
    name: 'toolOcr',
    desc: 'toolOcrDesc',
    input: 'pdf',
    minFiles: 1,
  },
  {
    id: 'images-to-pdf',
    group: 'convert',
    name: 'toolImagesToPdf',
    desc: 'toolImagesToPdfDesc',
    input: 'image',
    minFiles: 1,
  },
  {
    id: 'pdf-to-images',
    group: 'convert',
    name: 'toolPdfToImages',
    desc: 'toolPdfToImagesDesc',
    input: 'pdf',
    minFiles: 1,
  },
  {
    id: 'extract-images',
    group: 'convert',
    name: 'toolExtractImages',
    desc: 'toolExtractImagesDesc',
    input: 'pdf',
    minFiles: 1,
  },
  {
    id: 'protect',
    group: 'security',
    name: 'toolProtect',
    desc: 'toolProtectDesc',
    input: 'pdf',
    minFiles: 1,
  },
  {
    id: 'unlock',
    group: 'security',
    name: 'toolUnlock',
    desc: 'toolUnlockDesc',
    input: 'pdf',
    minFiles: 1,
  },
  {
    id: 'watermark',
    group: 'edit',
    name: 'toolWatermark',
    desc: 'toolWatermarkDesc',
    input: 'pdf',
    minFiles: 1,
  },
  {
    id: 'page-numbers',
    group: 'edit',
    name: 'toolPageNumbers',
    desc: 'toolPageNumbersDesc',
    input: 'pdf',
    minFiles: 1,
  },
  {
    id: 'properties',
    group: 'edit',
    name: 'toolProperties',
    desc: 'toolPropertiesDesc',
    input: 'pdf',
    minFiles: 1,
  },
]

export function toolDef(id: ToolId): ToolDef {
  return TOOLS.find((t) => t.id === id)!
}

/** Fresh options each time a tool opens. */
export function defaultOptions(): { [T in ToolId]: ToolOptions[T] } {
  return {
    merge: {},
    split: { kind: 'every', size: 1 },
    extract: { pages: '' },
    delete: { pages: '' },
    rotate: { angle: 90, pages: '' },
    reorder: { order: '' },
    compress: { level: 'balanced' },
    'images-to-pdf': { pageSize: 'a4', orientation: 'auto', margin: 18 },
    'pdf-to-images': { format: 'png', dpi: 150, quality: 85, pages: '' },
    'extract-images': {},
    protect: {
      userPassword: '',
      ownerPassword: '',
      allowPrint: true,
      allowCopy: true,
      allowEdit: true,
    },
    unlock: {},
    watermark: {
      text: '',
      fontSize: 72,
      color: '#c42b1c',
      opacity: 0.25,
      rotation: 'diagonal',
      position: 'center',
      pages: '',
    },
    'page-numbers': {
      format: '{n}',
      position: 'bottom-center',
      fontSize: 11,
      margin: 28,
      firstNumber: 1,
      pages: '',
    },
    flatten: {},
    repair: {},
    properties: {},
    ocr: { pages: '' },
  }
}
