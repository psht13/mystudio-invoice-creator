import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { Workbook } from 'exceljs'
import { describe, expect, it } from 'vitest'
import {
  buildDayRecords,
  buildSheetEntries,
  calculateTotalWorkHours,
  generateWorkbookFromTemplate,
  getDefaultInvoiceSelection,
  getPeriod,
  listPeriodDays,
  type DayOverride,
} from '../src/invoice'

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

function toArrayBuffer(buffer: Buffer): ArrayBuffer {
  const copy = new Uint8Array(buffer.byteLength)
  copy.set(buffer)
  return copy.buffer
}

async function loadTemplate(name: string): Promise<ArrayBuffer> {
  const templatePath = path.resolve(__dirname, '../public/templates', name)
  const buffer = await readFile(templatePath)
  return toArrayBuffer(buffer)
}

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
})

describe('sheet entry generation and totals', () => {
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
})

describe('workbook generation', () => {
  it('writes updated rows and stores the correct total formula result', async () => {
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
    const entries = buildSheetEntries(days)
    const totalHours = calculateTotalWorkHours(days)

    const resultBuffer = await generateWorkbookFromTemplate({
      template: await loadTemplate('template-1st.xlsx'),
      personName: 'Pavlo Yurchenko',
      periodLabel: '2/16/2026 - 2/28/2026',
      entries,
      totalHours,
    })

    const workbook = new Workbook()
    await workbook.xlsx.load(resultBuffer)

    const sheet = workbook.getWorksheet('Invoice')
    expect(sheet).toBeDefined()

    const ws = sheet!
    expect(ws.getCell('D5').value).toBe('2/16/2026 - 2/28/2026')
    expect(ws.getCell('D8').value).toBe('Pavlo Yurchenko')

    expect(ws.getCell('B15').value).toBe('2/16/2026')
    expect(ws.getCell('E15').value).toBe(6)
    expect(ws.getCell('B19').value).toBe('2/20/2026')
    expect(ws.getCell('E19').value).toBe(8)
    expect(ws.getCell('B20').value).toBe('2/21/2026 - 2/22/2026 (Weekend)')
    expect(ws.getCell('E20').value).toBeNull()

    const totalCell = ws.getCell('E29').value as { formula?: string; result?: number }
    expect(totalCell.formula).toBe('SUM(E15:E28)')
    expect(totalCell.result).toBe(46)

    expect(ws.getCell('B26').value).toBeNull()
    expect(ws.getCell('E26').value).toBeNull()
    expect(ws.getCell('B27').value).toBeNull()
    expect(ws.getCell('E27').value).toBeNull()
    expect(ws.getCell('B28').value).toBeNull()
    expect(ws.getCell('E28').value).toBeNull()
  })

  it('stores the correct total formula result for the 16th template', async () => {
    const period = getPeriod('2026-02', '16')
    const periodDays = listPeriodDays(period.startDate, period.endDate)

    const overrides: Record<string, DayOverride> = {
      '2026-02-03': { status: 'WORK', hours: 5 },
      '2026-02-10': { status: 'OOO', hours: 8 },
    }

    const days = buildDayRecords(periodDays, overrides)
    const entries = buildSheetEntries(days)
    const totalHours = calculateTotalWorkHours(days)

    const resultBuffer = await generateWorkbookFromTemplate({
      template: await loadTemplate('template-16th.xlsx'),
      personName: 'Pavlo Yurchenko',
      periodLabel: '2/1/2026 - 2/15/2026',
      entries,
      totalHours,
    })

    const workbook = new Workbook()
    await workbook.xlsx.load(resultBuffer)

    const ws = workbook.getWorksheet('Invoice')!
    const totalCell = ws.getCell('E29').value as { formula?: string; result?: number }

    expect(totalHours).toBe(69)
    expect(totalCell.formula).toBe('SUM(E15:E28)')
    expect(totalCell.result).toBe(69)
  })
})
