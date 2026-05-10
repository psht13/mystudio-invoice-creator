import { Workbook } from 'exceljs'
import type { GenerateWorkbookFromTemplateParams, WorkbookWriter } from '../../application/ports/workbook-writer'
import { DATA_END_ROW, DATA_START_ROW, TOTAL_FORMULA } from '../../domain/invoice/constants'
import { sanitizeHours } from '../../domain/invoice/hours'

type ExcelJSLoadBuffer = Parameters<Workbook['xlsx']['load']>[0]

function toArrayBuffer(value: ArrayBuffer | Uint8Array): ArrayBuffer {
  if (value instanceof ArrayBuffer) {
    return value
  }

  const copy = new Uint8Array(value.byteLength)
  copy.set(value)
  return copy.buffer
}

export class ExcelJSWorkbookWriter implements WorkbookWriter {
  async generateWorkbookFromTemplate(params: GenerateWorkbookFromTemplateParams): Promise<ArrayBuffer> {
    const workbook = new Workbook()
    await workbook.xlsx.load(params.template as ExcelJSLoadBuffer)

    const worksheet = workbook.getWorksheet('Invoice') ?? workbook.worksheets[0]
    if (!worksheet) {
      throw new Error('Template worksheet is missing')
    }

    workbook.calcProperties.fullCalcOnLoad = true

    worksheet.getCell('D5').value = params.periodLabel
    worksheet.getCell('D8').value = params.personName

    for (let row = DATA_START_ROW; row <= DATA_END_ROW; row += 1) {
      worksheet.getCell(`B${row}`).value = null
      worksheet.getCell(`E${row}`).value = null
    }

    params.entries.forEach((entry, index) => {
      const row = DATA_START_ROW + index
      worksheet.getCell(`B${row}`).value = entry.label

      if (entry.hours !== null) {
        worksheet.getCell(`E${row}`).value = entry.hours
      }
    })

    worksheet.getCell('E29').value = {
      formula: TOTAL_FORMULA,
      result: sanitizeHours(params.totalHours),
    }

    const written = await workbook.xlsx.writeBuffer()
    return toArrayBuffer(written)
  }
}
