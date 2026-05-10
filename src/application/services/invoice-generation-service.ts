import {
  buildDayRecords,
  buildInvoiceFileName,
  buildSheetEntriesWithTotal,
  formatDate,
  getPeriod,
  listPeriodDays,
  type DayOverride,
  type RunDay,
  type SheetEntry,
} from '../../domain/invoice'
import type { FileDownloader } from '../ports/file-downloader'
import type { TemplateLoader } from '../ports/template-loader'
import type { WorkbookWriter } from '../ports/workbook-writer'
import { validateInvoicePersonName, validateInvoiceSheetRowCapacity } from './invoice-generation-validation'

export type GenerateInvoiceRequest = {
  personName: string
  monthValue: string
  runDay: RunDay
  overrides: Record<string, DayOverride>
}

export type InvoiceGenerationNotice = {
  type: 'success' | 'error'
  text: string
}

export type GenerateInvoiceSuccess = {
  ok: true
  notice: InvoiceGenerationNotice
  fileName: string
  templatePath: string
  periodLabel: string
  entries: SheetEntry[]
  totalHours: number
}

export type GenerateInvoiceError = {
  ok: false
  notice: InvoiceGenerationNotice
}

export type GenerateInvoiceResult = GenerateInvoiceSuccess | GenerateInvoiceError

export type InvoiceGenerationServicePorts = {
  templateLoader: TemplateLoader
  workbookWriter: WorkbookWriter
  fileDownloader: FileDownloader
}

const TEMPLATE_PATH_BY_RUN_DAY: Record<RunDay, string> = {
  '1': '/templates/template-1st.xlsx',
  '16': '/templates/template-16th.xlsx',
}

function toErrorNotice(error: unknown): InvoiceGenerationNotice {
  return {
    type: 'error',
    text: error instanceof Error ? error.message : 'Workbook generation failed.',
  }
}

export class InvoiceGenerationService {
  readonly #templateLoader: TemplateLoader
  readonly #workbookWriter: WorkbookWriter
  readonly #fileDownloader: FileDownloader

  constructor(ports: InvoiceGenerationServicePorts) {
    this.#templateLoader = ports.templateLoader
    this.#workbookWriter = ports.workbookWriter
    this.#fileDownloader = ports.fileDownloader
  }

  async generateInvoice(request: GenerateInvoiceRequest): Promise<GenerateInvoiceResult> {
    const personNameValidation = validateInvoicePersonName(request.personName)

    if (!personNameValidation.ok) {
      return {
        ok: false,
        notice: personNameValidation.notice,
      }
    }

    const period = getPeriod(request.monthValue, request.runDay)
    const periodDays = listPeriodDays(period.startDate, period.endDate)
    const days = buildDayRecords(periodDays, request.overrides)
    const { entries, totalHours } = buildSheetEntriesWithTotal(days)
    const capacityValidation = validateInvoiceSheetRowCapacity(entries)

    if (!capacityValidation.ok) {
      return {
        ok: false,
        notice: capacityValidation.notice,
      }
    }

    const templatePath = TEMPLATE_PATH_BY_RUN_DAY[request.runDay]
    const periodLabel = `${formatDate(period.startDate)} - ${formatDate(period.endDate)}`
    const fileName = buildInvoiceFileName(request.personName, period.invoiceDate)

    try {
      const template = await this.#templateLoader.loadTemplate(templatePath)
      const workbookBuffer = await this.#workbookWriter.generateWorkbookFromTemplate({
        template,
        personName: personNameValidation.trimmedPersonName,
        periodLabel,
        entries,
        totalHours,
      })

      await this.#fileDownloader.downloadFile({
        fileName,
        buffer: workbookBuffer,
      })

      return {
        ok: true,
        notice: {
          type: 'success',
          text: `Downloaded ${fileName}`,
        },
        fileName,
        templatePath,
        periodLabel,
        entries,
        totalHours,
      }
    } catch (error) {
      return {
        ok: false,
        notice: toErrorNotice(error),
      }
    }
  }
}
