import type { SheetEntry } from '../../domain/invoice/types'

export type GenerateWorkbookFromTemplateParams = {
  template: ArrayBuffer
  personName: string
  periodLabel: string
  entries: SheetEntry[]
  totalHours: number
}

export interface WorkbookWriter {
  generateWorkbookFromTemplate(params: GenerateWorkbookFromTemplateParams): Promise<ArrayBuffer>
}
