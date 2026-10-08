import { PDFHexString, PDFName, PDFNumber, PDFRadioGroup, degrees, rgb } from 'pdf-lib'
import type { PDFDocument, PDFForm, PDFPage, PDFTextField } from 'pdf-lib'
import type { DrawingInput } from '../shared/ipc'

export type FieldDrawing = Extract<DrawingInput, { kind: 'field' }>

type FieldAppearanceOptions = NonNullable<Parameters<PDFTextField['addToPage']>[1]>

const BORDER_WIDTH = 1
/** Neutral widget border, part of the saved document (not themed) */
const BORDER_RGB: [number, number, number] = [0.55, 0.55, 0.6]

/** Field flag bit 2 (Required), PDF 32000 table 221 */
const FF_REQUIRED = 1 << 1

/**
 * pdf-lib's addToPage takes an origin + size that it turns (by the widget
 * rotation and border width) into the widget /Rect. Invert that so the saved
 * /Rect is exactly the rect the user drew in page space.
 */
export function widgetPlacement(
  rect: readonly [number, number, number, number],
  rotation: number,
  borderWidth = BORDER_WIDTH,
): { x: number; y: number; width: number; height: number } {
  const [x1, y1, x2, y2] = rect
  const w = x2 - x1
  const h = y2 - y1
  const b = borderWidth / 2
  const r = (((Math.round(rotation / 90) * 90) % 360) + 360) % 360
  if (r === 90) return { x: x1 + w - b, y: y1 + b, width: h - borderWidth, height: w - borderWidth }
  if (r === 180)
    return { x: x1 + w - b, y: y1 + h - b, width: w - borderWidth, height: h - borderWidth }
  if (r === 270)
    return { x: x1 + b, y: y1 + h - b, width: h - borderWidth, height: w - borderWidth }
  return { x: x1 + b, y: y1 + b, width: w - borderWidth, height: h - borderWidth }
}

/** Dots make pdf-lib build a field hierarchy; keep each new field a plain top-level name */
export function cleanFieldName(name: string): string {
  return name.trim().replace(/\./g, '_')
}

function freeName(form: PDFForm, name: string): string {
  if (!form.getFieldMaybe(name)) return name
  for (let n = 2; ; n++) {
    const next = `${name}_${n}`
    if (!form.getFieldMaybe(next)) return next
  }
}

/**
 * Create the AcroForm fields placed with the form designer. Widgets turn with
 * the page (/MK /R) so their contents read upright on rotated pages. A name
 * already taken in the file gets a numeric suffix, except radio buttons, which
 * join the existing radio group of that name.
 */
export function addFormFields(pdfDoc: PDFDocument, pages: PDFPage[], fields: FieldDrawing[]): void {
  if (fields.length === 0) return
  const form = pdfDoc.getForm()
  // Names created in this request: pending widgets of one name are one field
  const created = new Map<string, string>()
  for (const f of fields) {
    const page = pages[f.pageIndex]
    if (!page) continue
    const rotation = page.getRotation().angle
    const options: FieldAppearanceOptions = {
      ...widgetPlacement(f.rect, rotation),
      rotate: degrees((((Math.round(rotation / 90) * 90) % 360) + 360) % 360),
      borderWidth: BORDER_WIDTH,
      borderColor: rgb(...BORDER_RGB),
      backgroundColor: undefined,
      textColor: rgb(0, 0, 0),
    }
    const base = cleanFieldName(f.name) || f.fieldType
    const key = `${f.fieldType}:${base}`
    if (f.fieldType === 'radio') {
      let name = created.get(key)
      if (!name) {
        const existing = form.getFieldMaybe(base)
        name = existing instanceof PDFRadioGroup ? base : freeName(form, base)
        created.set(key, name)
      }
      const group =
        (form.getFieldMaybe(name) as PDFRadioGroup | undefined) ?? form.createRadioGroup(name)
      const taken = new Set(group.getOptions())
      const wanted = f.exportValue?.trim() || 'Choice'
      let value = wanted
      for (let n = 2; taken.has(value); n++) value = `${wanted} ${n}`
      group.addOptionToPage(value, page, options)
      if (f.required) group.enableRequired()
      continue
    }
    const name = freeName(form, base)
    if (f.fieldType === 'text') {
      const field = form.createTextField(name)
      if (f.multiline) field.enableMultiline()
      field.addToPage(page, options)
      if (f.required) field.enableRequired()
    } else if (f.fieldType === 'checkbox') {
      const field = form.createCheckBox(name)
      field.addToPage(page, options)
      if (f.required) field.enableRequired()
    } else if (f.fieldType === 'dropdown') {
      const field = form.createDropdown(name)
      const list = (f.options ?? []).map((o) => o.trim()).filter((o) => o.length > 0)
      if (list.length > 0) field.setOptions([...new Set(list)])
      field.addToPage(page, options)
      if (f.required) field.enableRequired()
    } else {
      addSignatureField(pdfDoc, form, page, name, f, rotation)
    }
  }
}

/** pdf-lib cannot create /Sig fields: write the merged field/widget dictionary directly */
function addSignatureField(
  pdfDoc: PDFDocument,
  form: PDFForm,
  page: PDFPage,
  name: string,
  f: FieldDrawing,
  rotation: number,
): void {
  const [x1, y1, x2, y2] = f.rect
  const widget = pdfDoc.context.obj({
    Type: 'Annot',
    Subtype: 'Widget',
    FT: 'Sig',
    Rect: [x1, y1, x2, y2],
    F: 4,
    P: page.ref,
    BS: { W: BORDER_WIDTH },
    MK: { BC: BORDER_RGB, R: (((Math.round(rotation / 90) * 90) % 360) + 360) % 360 },
  })
  widget.set(PDFName.of('T'), PDFHexString.fromText(name))
  if (f.required) widget.set(PDFName.of('Ff'), PDFNumber.of(FF_REQUIRED))
  const ref = pdfDoc.context.register(widget)
  form.acroForm.addField(ref)
  page.node.addAnnot(ref)
}
