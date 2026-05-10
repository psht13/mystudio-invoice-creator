import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { Workbook, type Cell, type Worksheet } from 'exceljs'
import JSZip from 'jszip'
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { WorkbookWriter } from '../src/application/ports/workbook-writer'
import { DATA_END_ROW, DATA_ROW_CAPACITY, DATA_START_ROW } from '../src/domain/invoice/constants'
import { buildDayRecords } from '../src/domain/invoice/day-records'
import { formatDate, formatFileDate } from '../src/domain/invoice/dates'
import { buildInvoiceFileName, sanitizeFileName } from '../src/domain/invoice/filenames'
import { calculateTotalWorkHours, formatHours, sanitizeHours } from '../src/domain/invoice/hours'
import { getDefaultInvoiceSelection, getPeriod, listPeriodDays, parseMonthValue } from '../src/domain/invoice/periods'
import { buildSheetEntries, buildSheetEntriesWithTotal } from '../src/domain/invoice/sheet-entries'
import type { DayOverride, DayRecord } from '../src/domain/invoice/types'
import {
  isRunDayValue,
  parseRunDayValue,
  validateHoursSanitization,
  validateMonthValueParsing,
  validatePersonNamePresence,
  validateSheetRowCapacity,
} from '../src/domain/invoice/validation'
import { ExcelJSWorkbookWriter } from '../src/infrastructure/excel/exceljs-workbook-writer'
import { LazyExcelJSWorkbookWriter } from '../src/infrastructure/excel/lazy-exceljs-workbook-writer'

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

const DOMAIN_MODULE_FILES = [
  'constants.ts',
  'types.ts',
  'dates.ts',
  'periods.ts',
  'day-records.ts',
  'sheet-entries.ts',
  'hours.ts',
  'filenames.ts',
  'validation.ts',
  'index.ts',
]

function toArrayBuffer(buffer: Buffer): ArrayBuffer {
  const copy = new Uint8Array(buffer.byteLength)
  copy.set(buffer)
  return copy.buffer
}

async function writeWorkbook(workbook: Workbook): Promise<ArrayBuffer> {
  const buffer = await workbook.xlsx.writeBuffer()
  return toArrayBuffer(buffer)
}

async function loadTemplate(name: string): Promise<ArrayBuffer> {
  const templatePath = path.resolve(__dirname, '../public/templates', name)
  const buffer = await readFile(templatePath)
  return toArrayBuffer(buffer)
}

async function loadWorkbook(buffer: ArrayBuffer): Promise<Workbook> {
  const workbook = new Workbook()
  await workbook.xlsx.load(buffer)
  return workbook
}

async function expectFullCalcOnLoad(buffer: ArrayBuffer): Promise<void> {
  const zip = await JSZip.loadAsync(buffer)
  const workbookXmlFile = zip.file('xl/workbook.xml')

  if (!workbookXmlFile) {
    throw new Error('Workbook XML is missing')
  }

  await expect(workbookXmlFile.async('string')).resolves.toMatch(/<calcPr\b[^>]*\bfullCalcOnLoad="1"/)
}

function utcDate(year: number, monthIndex: number, day: number): Date {
  return new Date(Date.UTC(year, monthIndex, day))
}

function isoDates(dates: Date[]): string[] {
  return dates.map((date) => date.toISOString().slice(0, 10))
}

function dayRecord(iso: string, status: DayRecord['status'], hours: number): DayRecord {
  return {
    iso,
    date: new Date(`${iso}T00:00:00.000Z`),
    status,
    hours,
  }
}

function expectInvoiceWorksheet(workbook: Workbook): Worksheet {
  const worksheet = workbook.getWorksheet('Invoice')
  expect(worksheet).toBeDefined()
  return worksheet as Worksheet
}

function normalizeCellValue(value: unknown): unknown {
  if (value instanceof Date) {
    return value.toISOString()
  }

  return value
}

function rowsBtoE(worksheet: Worksheet): unknown[][] {
  const rows: unknown[][] = []

  for (let row = DATA_START_ROW; row <= DATA_END_ROW; row += 1) {
    rows.push(['B', 'C', 'D', 'E'].map((column) => normalizeCellValue(worksheet.getCell(`${column}${row}`).value)))
  }

  return rows
}

function styleFingerprint(cell: Cell) {
  return {
    alignment: cell.alignment,
    border: cell.border,
    fill: cell.fill,
    font: cell.font,
    numFmt: cell.numFmt,
  }
}

function expectStylesPreserved(templateWorksheet: Worksheet, generatedWorksheet: Worksheet, addresses: string[]) {
  for (const address of addresses) {
    expect(styleFingerprint(generatedWorksheet.getCell(address))).toEqual(
      styleFingerprint(templateWorksheet.getCell(address)),
    )
  }
}

function expectTotalFormula(worksheet: Worksheet, expectedResult: number) {
  const value = worksheet.getCell('E29').value as { formula?: string; result?: number }

  expect(value.formula).toBe('SUM(E15:E28)')
  expect(value.result).toBe(expectedResult)
}

async function workbookSemanticFingerprint(buffer: ArrayBuffer) {
  const workbook = await loadWorkbook(buffer)
  const worksheet = expectInvoiceWorksheet(workbook)

  return {
    sheetNames: workbook.worksheets.map((sheet) => sheet.name),
    periodLabel: normalizeCellValue(worksheet.getCell('D5').value),
    personName: normalizeCellValue(worksheet.getCell('D8').value),
    rows: rowsBtoE(worksheet),
    total: normalizeCellValue(worksheet.getCell('E29').value),
  }
}

afterEach(() => {
  vi.useRealTimers()
})

describe('invoice domain module boundary', () => {
  it('does not import UI, workbook, build-tool, or browser dependencies', async () => {
    for (const fileName of DOMAIN_MODULE_FILES) {
      const source = await readFile(path.resolve(__dirname, '../src/domain/invoice', fileName), 'utf8')

      expect(source, fileName).not.toMatch(/(?:from|import)\s+['"](react|exceljs|vite)['"]/)
      expect(source, fileName).not.toMatch(
        /\b(window|document|fetch|Blob|File|URL|navigator|localStorage|sessionStorage|HTMLElement|HTMLAnchorElement)\b/,
      )
    }
  })
})

describe('invoice formatting and filename helpers', () => {
  it('formats date and hour labels exactly as the invoice UI expects', () => {
    const leapDay = utcDate(2024, 1, 29)

    expect(formatDate(leapDay)).toBe('2/29/2024')
    expect(formatFileDate(leapDay)).toBe('2_29_2024')
    expect(formatHours(8)).toBe('8')
    expect(formatHours(8.5)).toBe('8.50')
    expect(formatHours(8.456)).toBe('8.46')
  })

  it('sanitizes hours by clamping invalid values and rounding to cents', () => {
    expect(sanitizeHours(-1)).toBe(0)
    expect(sanitizeHours(Number.NaN)).toBe(0)
    expect(sanitizeHours(Number.POSITIVE_INFINITY)).toBe(0)
    expect(sanitizeHours(Number.NEGATIVE_INFINITY)).toBe(0)
    expect(sanitizeHours(7.454)).toBe(7.45)
    expect(sanitizeHours(7.456)).toBe(7.46)
  })

  it('sanitizes invoice file names without changing the generated naming pattern', () => {
    expect(sanitizeFileName('')).toBe('Employee')
    expect(sanitizeFileName('   ')).toBe('Employee')
    expect(sanitizeFileName(' <Pa:v/lo\\Yu|r?chen*ko> ')).toBe('PavloYurchenko')

    expect(buildInvoiceFileName(' <Pa:v/lo\\Yu|r?chen*ko> ', utcDate(2026, 0, 1))).toBe(
      'PavloYurchenko Time Tracking - 1_1_2026.xlsx',
    )
    expect(buildInvoiceFileName('***', utcDate(2026, 0, 1))).toBe('Employee Time Tracking - 1_1_2026.xlsx')
  })
})

describe('invoice validation helpers', () => {
  it('validates person name presence while exposing the same trimmed value used for generation', () => {
    expect(validatePersonNamePresence('')).toEqual({
      ok: false,
      reason: 'missing-person-name',
    })
    expect(validatePersonNamePresence('   ')).toEqual({
      ok: false,
      reason: 'missing-person-name',
    })
    expect(validatePersonNamePresence(' Pavlo Yurchenko ')).toEqual({
      ok: true,
      trimmedName: 'Pavlo Yurchenko',
    })
  })

  it('validates run-day values and preserves the existing non-16 period fallback', () => {
    expect(isRunDayValue('1')).toBe(true)
    expect(isRunDayValue('16')).toBe(true)
    expect(isRunDayValue('2')).toBe(false)
    expect(isRunDayValue(16)).toBe(false)

    expect(parseRunDayValue('1')).toBe('1')
    expect(parseRunDayValue('16')).toBe('16')
    expect(parseRunDayValue('2')).toBe('1')
    expect(getPeriod('2026-03', '2' as never)).toMatchObject({
      invoiceDate: utcDate(2026, 2, 1),
      startDate: utcDate(2026, 1, 16),
      endDate: utcDate(2026, 1, 28),
    })
  })

  it('parses month values explicitly and keeps the current-month fallback for invalid values', () => {
    const fallbackDate = new Date(2026, 4, 10, 12)

    expect(validateMonthValueParsing('2026-02', fallbackDate)).toEqual({
      ok: true,
      year: 2026,
      monthIndex: 1,
    })
    expect(validateMonthValueParsing('not-a-month', fallbackDate)).toEqual({
      ok: false,
      reason: 'invalid-month-value',
      fallback: {
        year: 2026,
        monthIndex: 4,
      },
    })
    expect(validateMonthValueParsing('2026-00', fallbackDate)).toEqual({
      ok: false,
      reason: 'invalid-month-value',
      fallback: {
        year: 2026,
        monthIndex: 4,
      },
    })
    expect(validateMonthValueParsing('2026-13', fallbackDate)).toEqual({
      ok: false,
      reason: 'invalid-month-value',
      fallback: {
        year: 2026,
        monthIndex: 4,
      },
    })
  })

  it('reports hour sanitization decisions without changing sanitized values', () => {
    expect(validateHoursSanitization(8)).toEqual({ value: 8, sanitized: false })
    expect(validateHoursSanitization(7.454)).toEqual({ value: 7.45, sanitized: true })
    expect(validateHoursSanitization(7.456)).toEqual({ value: 7.46, sanitized: true })
    expect(validateHoursSanitization(-1)).toEqual({ value: 0, sanitized: true })
    expect(validateHoursSanitization(Number.NaN)).toEqual({ value: 0, sanitized: true })
    expect(validateHoursSanitization(Number.POSITIVE_INFINITY)).toEqual({ value: 0, sanitized: true })
  })

  it('validates sheet row capacity at the existing template row boundary', () => {
    expect(validateSheetRowCapacity(DATA_ROW_CAPACITY)).toEqual({
      ok: true,
      rowCount: DATA_ROW_CAPACITY,
      capacity: DATA_ROW_CAPACITY,
    })
    expect(validateSheetRowCapacity(DATA_ROW_CAPACITY + 1)).toEqual({
      ok: false,
      reason: 'sheet-row-capacity-exceeded',
      rowCount: DATA_ROW_CAPACITY + 1,
      capacity: DATA_ROW_CAPACITY,
    })
  })
})

describe('invoice defaults and period parsing', () => {
  it('defaults to current month 16th before day 16', () => {
    const selection = getDefaultInvoiceSelection(new Date('2026-02-15T12:00:00Z'))

    expect(selection).toEqual({
      monthValue: '2026-02',
      runDay: '16',
    })
  })

  it('defaults to next month 1st on/after day 16', () => {
    const selection = getDefaultInvoiceSelection(new Date('2026-02-18T12:00:00Z'))

    expect(selection).toEqual({
      monthValue: '2026-03',
      runDay: '1',
    })
  })

  it('defaults January day 1 runs to the 16th invoice for the current month', () => {
    const selection = getDefaultInvoiceSelection(new Date(2026, 0, 1, 12))

    expect(selection).toEqual({
      monthValue: '2026-01',
      runDay: '16',
    })
  })

  it('crosses from December to January for next-month 1st runs', () => {
    const selection = getDefaultInvoiceSelection(new Date(2026, 11, 16, 12))

    expect(selection).toEqual({
      monthValue: '2027-01',
      runDay: '1',
    })
  })

  it('parses valid month values and falls back to the current month for invalid input', () => {
    vi.useFakeTimers()
    vi.setSystemTime(new Date(2026, 4, 10, 12))

    expect(parseMonthValue('2026-02')).toEqual({ year: 2026, monthIndex: 1 })
    expect(parseMonthValue('not-a-month')).toEqual({ year: 2026, monthIndex: 4 })
    expect(parseMonthValue('2026-00')).toEqual({ year: 2026, monthIndex: 4 })
    expect(parseMonthValue('2026-13')).toEqual({ year: 2026, monthIndex: 4 })
  })

  it('builds the correct half-month period ranges', () => {
    const firstRunPeriod = getPeriod('2026-03', '1')
    expect(firstRunPeriod).toMatchObject({
      invoiceDate: new Date(Date.UTC(2026, 2, 1)),
      startDate: new Date(Date.UTC(2026, 1, 16)),
      endDate: new Date(Date.UTC(2026, 1, 28)),
    })

    const sixteenthRunPeriod = getPeriod('2026-02', '16')
    expect(sixteenthRunPeriod).toMatchObject({
      invoiceDate: new Date(Date.UTC(2026, 1, 16)),
      startDate: new Date(Date.UTC(2026, 1, 1)),
      endDate: new Date(Date.UTC(2026, 1, 15)),
    })
  })

  it('handles leap-year February and non-leap February for 1st runs', () => {
    const leapPeriod = getPeriod('2024-03', '1')
    const nonLeapPeriod = getPeriod('2026-03', '1')

    expect(leapPeriod).toMatchObject({
      invoiceDate: utcDate(2024, 2, 1),
      startDate: utcDate(2024, 1, 16),
      endDate: utcDate(2024, 1, 29),
    })
    expect(isoDates(listPeriodDays(leapPeriod.startDate, leapPeriod.endDate))).toHaveLength(14)

    expect(nonLeapPeriod).toMatchObject({
      invoiceDate: utcDate(2026, 2, 1),
      startDate: utcDate(2026, 1, 16),
      endDate: utcDate(2026, 1, 28),
    })
    expect(isoDates(listPeriodDays(nonLeapPeriod.startDate, nonLeapPeriod.endDate))).toHaveLength(13)
  })

  it('builds January 1st runs from the previous December and lists days inclusively', () => {
    const period = getPeriod('2026-01', '1')

    expect(period).toMatchObject({
      invoiceDate: utcDate(2026, 0, 1),
      startDate: utcDate(2025, 11, 16),
      endDate: utcDate(2025, 11, 31),
    })
    expect(isoDates(listPeriodDays(period.startDate, period.endDate))).toEqual([
      '2025-12-16',
      '2025-12-17',
      '2025-12-18',
      '2025-12-19',
      '2025-12-20',
      '2025-12-21',
      '2025-12-22',
      '2025-12-23',
      '2025-12-24',
      '2025-12-25',
      '2025-12-26',
      '2025-12-27',
      '2025-12-28',
      '2025-12-29',
      '2025-12-30',
      '2025-12-31',
    ])
  })
})

describe('sheet entry generation and totals', () => {
  it('builds day records with weekend status taking precedence over overrides', () => {
    const days = buildDayRecords(listPeriodDays(utcDate(2026, 1, 14), utcDate(2026, 1, 17)), {
      '2026-02-14': { status: 'WORK', hours: 12 },
      '2026-02-15': { status: 'HOLIDAY', hours: 8 },
      '2026-02-16': { status: 'WORK', hours: -3 },
      '2026-02-17': { status: 'WORK', hours: Number.POSITIVE_INFINITY },
    })

    expect(days).toEqual([
      dayRecord('2026-02-14', 'WEEKEND', 0),
      dayRecord('2026-02-15', 'WEEKEND', 0),
      dayRecord('2026-02-16', 'WORK', -3),
      dayRecord('2026-02-17', 'WORK', Number.POSITIVE_INFINITY),
    ])
    expect(calculateTotalWorkHours(days)).toBe(0)
  })

  it('calculates totals from sanitized work hours only', () => {
    const days: DayRecord[] = [
      dayRecord('2026-02-02', 'WORK', 1.235),
      dayRecord('2026-02-03', 'WORK', Number.NaN),
      dayRecord('2026-02-04', 'WORK', Number.POSITIVE_INFINITY),
      dayRecord('2026-02-05', 'WORK', -2),
      dayRecord('2026-02-06', 'HOLIDAY', 99),
    ]

    expect(calculateTotalWorkHours(days)).toBe(1.24)
  })

  it('sanitizes work entry hours before rendering sheet rows', () => {
    const entries = buildSheetEntries([
      dayRecord('2026-02-02', 'WORK', 7.456),
      dayRecord('2026-02-03', 'WORK', -1),
      dayRecord('2026-02-04', 'WORK', Number.NaN),
    ])

    expect(entries).toEqual([
      { label: '2/2/2026', hours: 7.46 },
      { label: '2/3/2026', hours: 0 },
      { label: '2/4/2026', hours: 0 },
    ])
  })

  it('builds entries and totals together without changing separate calculation results', () => {
    const days: DayRecord[] = [
      dayRecord('2026-02-02', 'WORK', 7.456),
      dayRecord('2026-02-03', 'WORK', -1),
      dayRecord('2026-02-04', 'HOLIDAY', 8),
      dayRecord('2026-02-05', 'HOLIDAY', 8),
      dayRecord('2026-02-06', 'OOO', 8),
    ]

    const summary = buildSheetEntriesWithTotal(days)

    expect(summary.entries).toEqual(buildSheetEntries(days))
    expect(summary.totalHours).toBe(calculateTotalWorkHours(days))
  })

  it('groups non-work ranges and computes correct total work hours', () => {
    const period = getPeriod('2026-03', '1')
    const periodDays = listPeriodDays(period.startDate, period.endDate)

    const overrides: Record<string, DayOverride> = {
      '2026-02-16': { status: 'WORK', hours: 6 },
      '2026-02-23': { status: 'OOO', hours: 8 },
      '2026-02-24': { status: 'VACATION', hours: 8 },
      '2026-02-25': { status: 'VACATION', hours: 8 },
      '2026-02-26': { status: 'HOLIDAY', hours: 8 },
    }

    const days = buildDayRecords(periodDays, overrides)
    const totalHours = calculateTotalWorkHours(days)
    const entries = buildSheetEntries(days)

    expect(totalHours).toBe(46)
    expect(entries).toEqual([
      { label: '2/16/2026', hours: 6 },
      { label: '2/17/2026', hours: 8 },
      { label: '2/18/2026', hours: 8 },
      { label: '2/19/2026', hours: 8 },
      { label: '2/20/2026', hours: 8 },
      { label: '2/21/2026 - 2/22/2026 (Weekend)', hours: null },
      { label: '2/23/2026 (OOO)', hours: null },
      { label: '2/24/2026 - 2/25/2026 (Vacation)', hours: null },
      { label: '2/26/2026 (Holiday)', hours: null },
      { label: '2/27/2026', hours: 8 },
      { label: '2/28/2026 (Weekend)', hours: null },
    ])
  })

  it('collapses matching weekend, holiday, and vacation ranges without merging other status combinations', () => {
    const entries = buildSheetEntries([
      dayRecord('2026-01-03', 'WEEKEND', 0),
      dayRecord('2026-01-04', 'WEEKEND', 0),
      dayRecord('2026-01-05', 'HOLIDAY', 8),
      dayRecord('2026-01-06', 'HOLIDAY', 8),
      dayRecord('2026-01-07', 'VACATION', 8),
      dayRecord('2026-01-08', 'VACATION', 8),
      dayRecord('2026-01-09', 'OOO', 8),
      dayRecord('2026-01-10', 'OOO', 8),
      dayRecord('2026-01-11', 'HOLIDAY', 8),
      dayRecord('2026-01-12', 'VACATION', 8),
      dayRecord('2026-01-13', 'HOLIDAY', 8),
    ])

    expect(entries).toEqual([
      { label: '1/3/2026 - 1/4/2026 (Weekend)', hours: null },
      { label: '1/5/2026 - 1/6/2026 (Holiday)', hours: null },
      { label: '1/7/2026 - 1/8/2026 (Vacation)', hours: null },
      { label: '1/9/2026 (OOO)', hours: null },
      { label: '1/10/2026 (OOO)', hours: null },
      { label: '1/11/2026 (Holiday)', hours: null },
      { label: '1/12/2026 (Vacation)', hours: null },
      { label: '1/13/2026 (Holiday)', hours: null },
    ])
  })

  it('keeps consecutive OOO days as separate rows', () => {
    const period = getPeriod('2026-03', '1')
    const periodDays = listPeriodDays(period.startDate, period.endDate)

    const overrides: Record<string, DayOverride> = {
      '2026-02-16': { status: 'WORK', hours: 6 },
      '2026-02-23': { status: 'OOO', hours: 8 },
      '2026-02-24': { status: 'OOO', hours: 8 },
      '2026-02-25': { status: 'VACATION', hours: 8 },
      '2026-02-26': { status: 'VACATION', hours: 8 },
      '2026-02-27': { status: 'HOLIDAY', hours: 8 },
    }

    const days = buildDayRecords(periodDays, overrides)
    const totalHours = calculateTotalWorkHours(days)
    const entries = buildSheetEntries(days)

    expect(totalHours).toBe(38)
    expect(entries).toEqual([
      { label: '2/16/2026', hours: 6 },
      { label: '2/17/2026', hours: 8 },
      { label: '2/18/2026', hours: 8 },
      { label: '2/19/2026', hours: 8 },
      { label: '2/20/2026', hours: 8 },
      { label: '2/21/2026 - 2/22/2026 (Weekend)', hours: null },
      { label: '2/23/2026 (OOO)', hours: null },
      { label: '2/24/2026 (OOO)', hours: null },
      { label: '2/25/2026 - 2/26/2026 (Vacation)', hours: null },
      { label: '2/27/2026 (Holiday)', hours: null },
      { label: '2/28/2026 (Weekend)', hours: null },
    ])
  })
})

describe('workbook generation', () => {
  const workbookWriter: WorkbookWriter = new ExcelJSWorkbookWriter()
  const lazyWorkbookWriter: WorkbookWriter = new LazyExcelJSWorkbookWriter()

  it('writes semantic contents into the Invoice sheet for the 1st template', async () => {
    const period = getPeriod('2026-03', '1')
    const periodDays = listPeriodDays(period.startDate, period.endDate)
    const template = await loadTemplate('template-1st.xlsx')
    const templateWorkbook = await loadWorkbook(template)
    const templateWorksheet = expectInvoiceWorksheet(templateWorkbook)

    const overrides: Record<string, DayOverride> = {
      '2026-02-16': { status: 'WORK', hours: 6 },
      '2026-02-23': { status: 'OOO', hours: 8 },
      '2026-02-24': { status: 'VACATION', hours: 8 },
      '2026-02-25': { status: 'VACATION', hours: 8 },
      '2026-02-26': { status: 'HOLIDAY', hours: 8 },
    }

    const days = buildDayRecords(periodDays, overrides)
    const entries = buildSheetEntries(days)
    const totalHours = calculateTotalWorkHours(days)

    const resultBuffer = await workbookWriter.generateWorkbookFromTemplate({
      template: await loadTemplate('template-1st.xlsx'),
      personName: 'Pavlo Yurchenko',
      periodLabel: '2/16/2026 - 2/28/2026',
      entries,
      totalHours,
    })

    expect(resultBuffer).toBeInstanceOf(ArrayBuffer)
    await expectFullCalcOnLoad(resultBuffer)
    const workbook = await loadWorkbook(resultBuffer)

    expect(workbook.worksheets.map((worksheet) => worksheet.name)).toEqual(['Invoice'])
    const ws = expectInvoiceWorksheet(workbook)
    expect(ws.getCell('D5').value).toBe('2/16/2026 - 2/28/2026')
    expect(ws.getCell('D8').value).toBe('Pavlo Yurchenko')
    expect(rowsBtoE(ws)).toEqual([
      ['2/16/2026', '2/16/2026', '2/16/2026', 6],
      ['2/17/2026', null, null, 8],
      ['2/18/2026', null, null, 8],
      ['2/19/2026', null, null, 8],
      ['2/20/2026', null, null, 8],
      ['2/21/2026 - 2/22/2026 (Weekend)', null, null, null],
      ['2/23/2026 (OOO)', null, null, null],
      ['2/24/2026 - 2/25/2026 (Vacation)', null, null, null],
      ['2/26/2026 (Holiday)', null, null, null],
      ['2/27/2026', null, null, 8],
      ['2/28/2026 (Weekend)', null, null, null],
      [null, null, null, null],
      [null, null, null, null],
      [null, null, null, null],
    ])
    expectTotalFormula(ws, 46)
    expectStylesPreserved(templateWorksheet, ws, ['D5', 'D8', 'B15', 'E15', 'B28', 'E28', 'E29'])
  })

  it('writes semantic contents into the Invoice sheet for the 16th template', async () => {
    const period = getPeriod('2026-02', '16')
    const periodDays = listPeriodDays(period.startDate, period.endDate)
    const template = await loadTemplate('template-16th.xlsx')
    const templateWorkbook = await loadWorkbook(template)
    const templateWorksheet = expectInvoiceWorksheet(templateWorkbook)

    const overrides: Record<string, DayOverride> = {
      '2026-02-03': { status: 'WORK', hours: 5 },
      '2026-02-10': { status: 'OOO', hours: 8 },
    }

    const days = buildDayRecords(periodDays, overrides)
    const entries = buildSheetEntries(days)
    const totalHours = calculateTotalWorkHours(days)

    const resultBuffer = await workbookWriter.generateWorkbookFromTemplate({
      template: await loadTemplate('template-16th.xlsx'),
      personName: 'Pavlo Yurchenko',
      periodLabel: '2/1/2026 - 2/15/2026',
      entries,
      totalHours,
    })

    expect(resultBuffer).toBeInstanceOf(ArrayBuffer)
    await expectFullCalcOnLoad(resultBuffer)
    const workbook = await loadWorkbook(resultBuffer)

    expect(workbook.worksheets.map((worksheet) => worksheet.name)).toEqual(['Invoice'])
    const ws = expectInvoiceWorksheet(workbook)

    expect(ws.getCell('D5').value).toBe('2/1/2026 - 2/15/2026')
    expect(ws.getCell('D8').value).toBe('Pavlo Yurchenko')
    expect(rowsBtoE(ws)).toEqual([
      ['2/1/2026 (Weekend)', '2/1/2026 (Weekend)', '2/1/2026 (Weekend)', null],
      ['2/2/2026', null, null, 8],
      ['2/3/2026', null, null, 5],
      ['2/4/2026', null, null, 8],
      ['2/5/2026', null, null, 8],
      ['2/6/2026', null, null, 8],
      ['2/7/2026 - 2/8/2026 (Weekend)', null, null, null],
      ['2/9/2026', null, null, 8],
      ['2/10/2026 (OOO)', null, null, null],
      ['2/11/2026', null, null, 8],
      ['2/12/2026', null, null, 8],
      ['2/13/2026', null, null, 8],
      ['2/14/2026 - 2/15/2026 (Weekend)', null, null, null],
      [null, null, null, null],
    ])
    expect(totalHours).toBe(69)
    expectTotalFormula(ws, 69)
    expectStylesPreserved(templateWorksheet, ws, ['D5', 'D8', 'B15', 'E15', 'B28', 'E28', 'E29'])
  })

  it('keeps lazy workbook generation output semantically identical to the direct writer', async () => {
    const params = {
      template: await loadTemplate('template-16th.xlsx'),
      personName: 'Pavlo Yurchenko',
      periodLabel: '2/1/2026 - 2/15/2026',
      entries: [
        { label: '2/1/2026 (Weekend)', hours: null },
        { label: '2/2/2026', hours: 7.46 },
        { label: '2/3/2026', hours: 8 },
      ],
      totalHours: 15.46,
    }

    const directBuffer = await workbookWriter.generateWorkbookFromTemplate(params)
    const lazyBuffer = await lazyWorkbookWriter.generateWorkbookFromTemplate(params)

    await expectFullCalcOnLoad(lazyBuffer)
    await expect(workbookSemanticFingerprint(lazyBuffer)).resolves.toEqual(
      await workbookSemanticFingerprint(directBuffer),
    )
  })

  it('falls back to the first worksheet when the Invoice worksheet is absent', async () => {
    const templateWorkbook = new Workbook()
    const templateWorksheet = templateWorkbook.addWorksheet('Timesheet')
    templateWorksheet.getCell('D5').value = 'stale period'
    templateWorksheet.getCell('D8').value = 'stale person'
    templateWorksheet.getCell('B15').value = 'stale row'
    templateWorksheet.getCell('E15').value = 999
    templateWorksheet.getCell('B16').value = 'stale row'
    templateWorksheet.getCell('E16').value = 999

    const resultBuffer = await workbookWriter.generateWorkbookFromTemplate({
      template: await writeWorkbook(templateWorkbook),
      personName: 'Fallback User',
      periodLabel: '3/1/2026 - 3/15/2026',
      entries: [
        { label: '3/1/2026 (Weekend)', hours: null },
        { label: '3/2/2026', hours: 7.46 },
      ],
      totalHours: 7.456,
    })

    expect(resultBuffer).toBeInstanceOf(ArrayBuffer)
    await expectFullCalcOnLoad(resultBuffer)
    const workbook = await loadWorkbook(resultBuffer)

    expect(workbook.worksheets.map((worksheet) => worksheet.name)).toEqual(['Timesheet'])
    const worksheet = workbook.worksheets[0]
    expect(worksheet.getCell('D5').value).toBe('3/1/2026 - 3/15/2026')
    expect(worksheet.getCell('D8').value).toBe('Fallback User')
    expect(rowsBtoE(worksheet).slice(0, 3)).toEqual([
      ['3/1/2026 (Weekend)', null, null, null],
      ['3/2/2026', null, null, 7.46],
      [null, null, null, null],
    ])
    expectTotalFormula(worksheet, 7.46)
  })

  it('keeps the missing worksheet error for empty templates', async () => {
    const emptyWorkbook = new Workbook()

    await expect(
      workbookWriter.generateWorkbookFromTemplate({
        template: await writeWorkbook(emptyWorkbook),
        personName: 'Pavlo Yurchenko',
        periodLabel: '2/1/2026 - 2/15/2026',
        entries: [],
        totalHours: 0,
      }),
    ).rejects.toThrow('Template worksheet is missing')
  })
})
