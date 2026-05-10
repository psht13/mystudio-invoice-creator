import type { GenerateWorkbookFromTemplateParams, WorkbookWriter } from '../../application/ports/workbook-writer'

export class LazyExcelJSWorkbookWriter implements WorkbookWriter {
  #writerPromise: Promise<WorkbookWriter> | null = null

  async #getWriter(): Promise<WorkbookWriter> {
    this.#writerPromise ??= import('./exceljs-workbook-writer').then(
      ({ ExcelJSWorkbookWriter }) => new ExcelJSWorkbookWriter(),
    )

    return this.#writerPromise
  }

  async generateWorkbookFromTemplate(params: GenerateWorkbookFromTemplateParams): Promise<ArrayBuffer> {
    const writer = await this.#getWriter()
    return writer.generateWorkbookFromTemplate(params)
  }
}
