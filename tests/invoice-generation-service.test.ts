import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { DownloadFileParams, FileDownloader } from '../src/application/ports/file-downloader'
import type { TemplateLoader } from '../src/application/ports/template-loader'
import type { GenerateWorkbookFromTemplateParams, WorkbookWriter } from '../src/application/ports/workbook-writer'
import {
  validateInvoicePersonName,
  validateInvoiceSheetRowCapacity,
} from '../src/application/services/invoice-generation-validation'
import { DATA_ROW_CAPACITY } from '../src/domain/invoice/constants'
import type { DayOverride, RunDay } from '../src/domain/invoice'

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

class FakeTemplateLoader implements TemplateLoader {
  readonly paths: string[] = []
  result = new ArrayBuffer(4)
  error: unknown

  async loadTemplate(templatePath: string): Promise<ArrayBuffer> {
    this.paths.push(templatePath)

    if (this.error) {
      throw this.error
    }

    return this.result
  }
}

class FakeWorkbookWriter implements WorkbookWriter {
  readonly requests: GenerateWorkbookFromTemplateParams[] = []
  result = new ArrayBuffer(8)
  error: unknown

  async generateWorkbookFromTemplate(params: GenerateWorkbookFromTemplateParams): Promise<ArrayBuffer> {
    this.requests.push(params)

    if (this.error) {
      throw this.error
    }

    return this.result
  }
}

class FakeFileDownloader implements FileDownloader {
  readonly downloads: DownloadFileParams[] = []
  error: unknown

  async downloadFile(params: DownloadFileParams): Promise<void> {
    this.downloads.push(params)

    if (this.error) {
      throw this.error
    }
  }
}

type FakePorts = {
  templateLoader: FakeTemplateLoader
  workbookWriter: FakeWorkbookWriter
  fileDownloader: FakeFileDownloader
}

function createFakePorts(): FakePorts {
  return {
    templateLoader: new FakeTemplateLoader(),
    workbookWriter: new FakeWorkbookWriter(),
    fileDownloader: new FakeFileDownloader(),
  }
}

async function createService(ports = createFakePorts()) {
  const { InvoiceGenerationService } = await import('../src/application/services/invoice-generation-service')

  return {
    service: new InvoiceGenerationService(ports),
    ports,
  }
}

function request(overrides: Record<string, DayOverride> = {}, runDay: RunDay = '16') {
  return {
    personName: ' Pavlo Yurchenko ',
    monthValue: '2026-02',
    runDay,
    overrides,
  }
}

afterEach(() => {
  vi.useRealTimers()
  vi.doUnmock('../src/domain/invoice')
  vi.resetModules()
})

describe('InvoiceGeneration validation helpers', () => {
  it('returns the existing empty-name notice text and trimmed valid names', () => {
    expect(validateInvoicePersonName('   ')).toEqual({
      ok: false,
      notice: {
        type: 'error',
        text: 'Name is required before generating the invoice.',
      },
    })
    expect(validateInvoicePersonName(' Pavlo Yurchenko ')).toEqual({
      ok: true,
      trimmedPersonName: 'Pavlo Yurchenko',
    })
  })

  it('returns the existing capacity notice text for rows beyond the template range', () => {
    const overflowingEntries = Array.from({ length: DATA_ROW_CAPACITY + 1 }, (_, index) => ({
      label: `row ${index + 1}`,
      hours: 8,
    }))

    expect(validateInvoiceSheetRowCapacity([])).toEqual({
      ok: true,
      rowCount: 0,
      capacity: DATA_ROW_CAPACITY,
    })
    expect(validateInvoiceSheetRowCapacity(overflowingEntries)).toEqual({
      ok: false,
      notice: {
        type: 'error',
        text: `Template limit reached: ${DATA_ROW_CAPACITY + 1} rows needed, ${DATA_ROW_CAPACITY} rows available.`,
      },
      rowCount: DATA_ROW_CAPACITY + 1,
      capacity: DATA_ROW_CAPACITY,
    })
  })
})

describe('InvoiceGenerationService', () => {
  it('keeps browser-specific behavior out of the service module', async () => {
    const source = await readFile(
      path.resolve(__dirname, '../src/application/services/invoice-generation-service.ts'),
      'utf8',
    )

    expect(source).not.toMatch(/(?:from|import)\s+['"]react['"]/)
    expect(source).not.toMatch(/\b(window|document|fetch|Blob|File|URL|navigator|localStorage|sessionStorage)\b/)
  })

  it('coordinates the success path and returns the existing success notice', async () => {
    const { service, ports } = await createService()

    const result = await service.generateInvoice(
      request({
        '2026-02-03': { status: 'WORK', hours: 5 },
        '2026-02-10': { status: 'OOO', hours: 8 },
      }),
    )

    expect(result).toMatchObject({
      ok: true,
      notice: {
        type: 'success',
        text: 'Downloaded Pavlo Yurchenko Time Tracking - 2_16_2026.xlsx',
      },
      fileName: 'Pavlo Yurchenko Time Tracking - 2_16_2026.xlsx',
      templatePath: '/templates/template-16th.xlsx',
      periodLabel: '2/1/2026 - 2/15/2026',
      totalHours: 69,
    })
    expect(ports.templateLoader.paths).toEqual(['/templates/template-16th.xlsx'])

    expect(ports.workbookWriter.requests).toHaveLength(1)
    const workbookRequest = ports.workbookWriter.requests[0]
    expect(workbookRequest.template).toBe(ports.templateLoader.result)
    expect(workbookRequest.personName).toBe('Pavlo Yurchenko')
    expect(workbookRequest.periodLabel).toBe('2/1/2026 - 2/15/2026')
    expect(workbookRequest.totalHours).toBe(69)
    expect(workbookRequest.entries).toEqual([
      { label: '2/1/2026 (Weekend)', hours: null },
      { label: '2/2/2026', hours: 8 },
      { label: '2/3/2026', hours: 5 },
      { label: '2/4/2026', hours: 8 },
      { label: '2/5/2026', hours: 8 },
      { label: '2/6/2026', hours: 8 },
      { label: '2/7/2026 - 2/8/2026 (Weekend)', hours: null },
      { label: '2/9/2026', hours: 8 },
      { label: '2/10/2026 (OOO)', hours: null },
      { label: '2/11/2026', hours: 8 },
      { label: '2/12/2026', hours: 8 },
      { label: '2/13/2026', hours: 8 },
      { label: '2/14/2026 - 2/15/2026 (Weekend)', hours: null },
    ])

    expect(ports.fileDownloader.downloads).toHaveLength(1)
    expect(ports.fileDownloader.downloads[0].fileName).toBe('Pavlo Yurchenko Time Tracking - 2_16_2026.xlsx')
    expect(ports.fileDownloader.downloads[0].buffer).toBe(ports.workbookWriter.result)
  })

  it('returns the existing empty-name error and does not call ports', async () => {
    const { service, ports } = await createService()

    const result = await service.generateInvoice({
      ...request(),
      personName: '   ',
    })

    expect(result).toEqual({
      ok: false,
      notice: {
        type: 'error',
        text: 'Name is required before generating the invoice.',
      },
    })
    expect(ports.templateLoader.paths).toEqual([])
    expect(ports.workbookWriter.requests).toEqual([])
    expect(ports.fileDownloader.downloads).toEqual([])
  })

  it('returns the existing capacity error before loading the template', async () => {
    vi.doMock('../src/domain/invoice', async (importOriginal) => {
      const actual = await importOriginal<typeof import('../src/domain/invoice')>()

      return {
        ...actual,
        buildSheetEntriesWithTotal: vi.fn(() => ({
          entries: Array.from({ length: actual.DATA_ROW_CAPACITY + 1 }, (_, index) => ({
            label: `row ${index + 1}`,
            hours: 8,
          })),
          totalHours: 0,
        })),
      }
    })
    const { DATA_ROW_CAPACITY } = await import('../src/domain/invoice')
    const { service, ports } = await createService()

    const result = await service.generateInvoice(request())

    expect(result).toEqual({
      ok: false,
      notice: {
        type: 'error',
        text: `Template limit reached: ${DATA_ROW_CAPACITY + 1} rows needed, ${DATA_ROW_CAPACITY} rows available.`,
      },
    })
    expect(ports.templateLoader.paths).toEqual([])
    expect(ports.workbookWriter.requests).toEqual([])
    expect(ports.fileDownloader.downloads).toEqual([])
  })

  it('preserves invalid month fallback behavior in service results', async () => {
    vi.useFakeTimers()
    vi.setSystemTime(new Date(2026, 4, 10, 12))
    const { service, ports } = await createService()

    const result = await service.generateInvoice({
      ...request(),
      monthValue: 'not-a-month',
    })

    expect(result).toMatchObject({
      ok: true,
      notice: {
        type: 'success',
        text: 'Downloaded Pavlo Yurchenko Time Tracking - 5_16_2026.xlsx',
      },
      fileName: 'Pavlo Yurchenko Time Tracking - 5_16_2026.xlsx',
      templatePath: '/templates/template-16th.xlsx',
      periodLabel: '5/1/2026 - 5/15/2026',
    })
    expect(ports.templateLoader.paths).toEqual(['/templates/template-16th.xlsx'])
  })

  it('preserves sanitized hour values in service entries and totals', async () => {
    const { service, ports } = await createService()

    const result = await service.generateInvoice(
      request({
        '2026-02-03': { status: 'WORK', hours: 7.456 },
        '2026-02-04': { status: 'WORK', hours: -1 },
        '2026-02-05': { status: 'WORK', hours: Number.NaN },
        '2026-02-06': { status: 'WORK', hours: Number.POSITIVE_INFINITY },
      }),
    )

    expect(result).toMatchObject({
      ok: true,
      totalHours: 55.46,
    })
    expect(ports.workbookWriter.requests[0].entries).toEqual(
      expect.arrayContaining([
        { label: '2/3/2026', hours: 7.46 },
        { label: '2/4/2026', hours: 0 },
        { label: '2/5/2026', hours: 0 },
        { label: '2/6/2026', hours: 0 },
      ]),
    )
    expect(ports.workbookWriter.requests[0].totalHours).toBe(55.46)
  })

  it('returns a template load failure without asking later ports to run', async () => {
    const ports = createFakePorts()
    ports.templateLoader.error = new Error('Failed to load template file: /templates/template-16th.xlsx')
    const { service } = await createService(ports)

    const result = await service.generateInvoice(request())

    expect(result).toEqual({
      ok: false,
      notice: {
        type: 'error',
        text: 'Failed to load template file: /templates/template-16th.xlsx',
      },
    })
    expect(ports.templateLoader.paths).toEqual(['/templates/template-16th.xlsx'])
    expect(ports.workbookWriter.requests).toEqual([])
    expect(ports.fileDownloader.downloads).toEqual([])
  })

  it('returns a workbook writer failure without downloading', async () => {
    const ports = createFakePorts()
    ports.workbookWriter.error = new Error('Workbook writer failed.')
    const { service } = await createService(ports)

    const result = await service.generateInvoice(request())

    expect(result).toEqual({
      ok: false,
      notice: {
        type: 'error',
        text: 'Workbook writer failed.',
      },
    })
    expect(ports.templateLoader.paths).toEqual(['/templates/template-16th.xlsx'])
    expect(ports.workbookWriter.requests).toHaveLength(1)
    expect(ports.fileDownloader.downloads).toEqual([])
  })

  it('returns a downloader failure after the workbook is written', async () => {
    const ports = createFakePorts()
    ports.fileDownloader.error = new Error('Downloader failed.')
    const { service } = await createService(ports)

    const result = await service.generateInvoice(request())

    expect(result).toEqual({
      ok: false,
      notice: {
        type: 'error',
        text: 'Downloader failed.',
      },
    })
    expect(ports.templateLoader.paths).toEqual(['/templates/template-16th.xlsx'])
    expect(ports.workbookWriter.requests).toHaveLength(1)
    expect(ports.fileDownloader.downloads).toHaveLength(1)
  })

  it.each([
    ['1', '/templates/template-1st.xlsx', 'Pavlo Yurchenko Time Tracking - 2_1_2026.xlsx'],
    ['16', '/templates/template-16th.xlsx', 'Pavlo Yurchenko Time Tracking - 2_16_2026.xlsx'],
  ] as const)('selects the %s run-day template path', async (runDay, templatePath, fileName) => {
    const { service, ports } = await createService()

    const result = await service.generateInvoice(request({}, runDay))

    expect(result).toMatchObject({
      ok: true,
      fileName,
      templatePath,
      notice: {
        type: 'success',
        text: `Downloaded ${fileName}`,
      },
    })
    expect(ports.templateLoader.paths).toEqual([templatePath])
  })
})
