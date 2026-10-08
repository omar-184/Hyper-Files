import { HOME_TEMPLATE_IDS, type HomeTemplateId } from '../shared/home-api'

/** Home "New from template": which editor opens each starter file, and its default name */
export const HOME_TEMPLATES: Record<
  HomeTemplateId,
  { readonly kind: 'doc' | 'sheet' | 'slide'; readonly fileName: string }
> = {
  letter: { kind: 'doc', fileName: 'Letter.docx' },
  resume: { kind: 'doc', fileName: 'Resume.docx' },
  budget: { kind: 'sheet', fileName: 'Monthly Budget.xlsx' },
  invoice: { kind: 'sheet', fileName: 'Invoice.xlsx' },
  presentation: { kind: 'slide', fileName: 'Presentation.pptx' },
}

export function isHomeTemplateId(value: unknown): value is HomeTemplateId {
  return typeof value === 'string' && (HOME_TEMPLATE_IDS as readonly string[]).includes(value)
}

/** The starter file's bytes; the encoded data loads only when a template is picked */
export async function homeTemplateBytes(id: HomeTemplateId): Promise<Buffer> {
  const { HOME_TEMPLATE_DATA } = await import('./home-templates-data')
  return Buffer.from(HOME_TEMPLATE_DATA[id], 'base64')
}
