import { describe, expect, it } from 'vitest'
import {
  PDFCheckBox,
  PDFDict,
  PDFDocument,
  PDFDropdown,
  PDFName,
  PDFNumber,
  PDFRadioGroup,
  PDFTextField,
  degrees,
} from 'pdf-lib'
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs'
import { applySaveRequest } from '../src/main/save-pdf'
import { cleanFieldName, widgetPlacement } from '../src/main/form-fields'
import type { DrawingInput, NewFieldType, SavePdfRequest } from '../src/shared/ipc'

const request = (drawings: DrawingInput[]): SavePdfRequest => ({
  path: '/tmp/test.pdf',
  markups: [],
  drawings,
  formValues: [],
  stamps: [],
})

const field = (
  fieldType: NewFieldType,
  name: string,
  over: Partial<Extract<DrawingInput, { kind: 'field' }>> = {},
): DrawingInput => ({
  kind: 'field',
  pageIndex: 0,
  rect: [100, 600, 260, 624],
  fieldType,
  name,
  ...over,
})

async function save(drawings: DrawingInput[], prepare?: (doc: PDFDocument) => void) {
  const src = await PDFDocument.create()
  src.addPage([612, 792])
  src.addPage([612, 792]).setRotation(degrees(90))
  prepare?.(src)
  const { bytes } = await applySaveRequest(await src.save(), request(drawings))
  return bytes
}

describe('widgetPlacement', () => {
  it('maps back to the drawn rect through pdf-lib for every rotation', async () => {
    const rect: [number, number, number, number] = [100, 200, 260, 230]
    for (const rotation of [0, 90, 180, 270]) {
      const doc = await PDFDocument.create()
      const page = doc.addPage([612, 792])
      const tf = doc.getForm().createTextField(`f${rotation}`)
      tf.addToPage(page, {
        ...widgetPlacement(rect, rotation),
        rotate: degrees(rotation),
        borderWidth: 1,
      })
      const r = tf.acroField.getWidgets()[0]!.getRectangle()
      expect(
        [r.x, r.y, r.x + r.width, r.y + r.height].map((v) => Math.round(v * 100) / 100),
      ).toEqual(rect)
    }
  })

  it('keeps names flat', () => {
    expect(cleanFieldName('  Address.Line 1 ')).toBe('Address_Line 1')
  })
})

describe('form designer save', () => {
  it('creates text, checkbox, dropdown, radio and signature fields', async () => {
    const bytes = await save([
      field('text', 'Full name', { required: true }),
      field('text', 'Notes', { rect: [100, 400, 400, 480], multiline: true }),
      field('checkbox', 'Agree', { rect: [100, 560, 114, 574] }),
      field('dropdown', 'Country', { options: ['Egypt', 'Jordan', 'Egypt', ' '] }),
      field('radio', 'Gender', { rect: [100, 520, 112, 532], exportValue: 'Male' }),
      field('radio', 'Gender', { rect: [140, 520, 152, 532], exportValue: 'Female' }),
      field('signature', 'Signature', { rect: [300, 100, 500, 150] }),
    ])
    const form = (await PDFDocument.load(bytes)).getForm()
    const name = form.getField('Full name')
    expect(name).toBeInstanceOf(PDFTextField)
    expect(name.isRequired()).toBe(true)
    expect((form.getField('Notes') as PDFTextField).isMultiline()).toBe(true)
    expect(form.getField('Agree')).toBeInstanceOf(PDFCheckBox)
    expect((form.getField('Country') as PDFDropdown).getOptions()).toEqual(['Egypt', 'Jordan'])
    expect((form.getField('Gender') as PDFRadioGroup).getOptions()).toEqual(['Male', 'Female'])
    const sig = form.getField('Signature')
    expect(sig.acroField.dict.get(PDFName.of('FT'))).toBe(PDFName.of('Sig'))

    // pdf.js (the viewer) sees the same fields and their kinds
    const doc = await getDocument({ data: bytes }).promise
    const annots = await (await doc.getPage(1)).getAnnotations()
    const byName = new Map(annots.map((a: { fieldName: string }) => [a.fieldName, a]))
    expect(byName.get('Full name')?.fieldType).toBe('Tx')
    expect(byName.get('Agree')?.checkBox).toBe(true)
    expect(byName.get('Country')?.combo).toBe(true)
    expect(byName.get('Signature')?.fieldType).toBe('Sig')
    expect(annots.filter((a: { fieldName: string }) => a.fieldName === 'Gender')).toHaveLength(2)
  })

  it('suffixes a name already in the file but joins an existing radio group', async () => {
    const bytes = await save(
      [
        field('text', 'Name'),
        field('radio', 'Choice', { exportValue: 'B', rect: [100, 500, 112, 512] }),
      ],
      (doc) => {
        const page = doc.getPage(0)
        doc.getForm().createTextField('Name').addToPage(page, { x: 10, y: 10 })
        const g = doc.getForm().createRadioGroup('Choice')
        g.addOptionToPage('A', page, { x: 10, y: 100, width: 12, height: 12 })
      },
    )
    const form = (await PDFDocument.load(bytes)).getForm()
    expect(form.getFieldMaybe('Name_2')).toBeInstanceOf(PDFTextField)
    expect((form.getField('Choice') as PDFRadioGroup).getOptions()).toEqual(['A', 'B'])
  })

  it('turns widgets with a rotated page and keeps the drawn rect', async () => {
    const bytes = await save([
      field('text', 'Rotated', { pageIndex: 1, rect: [300, 200, 330, 400] }),
    ])
    const tf = (await PDFDocument.load(bytes)).getForm().getTextField('Rotated')
    const widget = tf.acroField.getWidgets()[0]!
    const r = widget.getRectangle()
    expect([r.x, r.y, r.x + r.width, r.y + r.height]).toEqual([300, 200, 330, 400])
    const mk = widget.dict.lookup(PDFName.of('MK'), PDFDict)
    expect(mk.lookup(PDFName.of('R'), PDFNumber).asNumber()).toBe(90)
  })
})
