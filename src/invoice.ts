import { Workbook } from 'exceljs'

export type RunDay = '1' | '16'
export type EditableStatus = 'WORK' | 'HOLIDAY' | 'VACATION' | 'OOO'
export type DayStatus = EditableStatus | 'WEEKEND'

export type DayOverride = {
  status: EditableStatus
  hours: number
}

export type DayRecord = {
  iso: string
  date: Date
  status: DayStatus
  hours: number
}

export type SheetEntry = {
  label: string
  hours: number | null
}

export type Period = {
  invoiceDate: Date
  startDate: Date
  endDate: Date
}

export type DefaultSelection = {
  monthValue: string
  runDay: RunDay
}

export const DEFAULT_HOURS = 8
export const DATA_START_ROW = 15
export const DATA_END_ROW = 28
export const DATA_ROW_CAPACITY = DATA_END_ROW - DATA_START_ROW + 1

const DAY_MS = 24 * 60 * 60 * 1000
const TOTAL_FORMULA = 'SUM(E15:E28)'

export const STATUS_LABEL: Record<DayStatus, string> = {
  WORK: 'Work',
  WEEKEND: 'Weekend',
  HOLIDAY: 'Holiday',
  VACATION: 'Vacation',
  OOO: 'OOO',
}

function shouldCollapseNonWorkRange(status: DayStatus): boolean {
  return status === 'WEEKEND' || status === 'HOLIDAY' || status === 'VACATION'
}

function createUtcDate(year: number, monthIndex: number, day: number): Date {
  return new Date(Date.UTC(year, monthIndex, day))
}

function addUtcDays(date: Date, days: number): Date {
  return new Date(date.getTime() + days * DAY_MS)
}

function formatMonthInputValue(year: number, monthIndex: number): string {
  const month = String(monthIndex + 1).padStart(2, '0')
  return `${year}-${month}`
}

function toArrayBuffer(value: ArrayBuffer | Uint8Array): ArrayBuffer {
  if (value instanceof ArrayBuffer) {
    return value
  }

  const copy = new Uint8Array(value.byteLength)
  copy.set(value)
  return copy.buffer
}

export function toIsoDate(date: Date): string {
  const month = String(date.getUTCMonth() + 1).padStart(2, '0')
  const day = String(date.getUTCDate()).padStart(2, '0')
  return `${date.getUTCFullYear()}-${month}-${day}`
}

export function formatDate(date: Date): string {
  return `${date.getUTCMonth() + 1}/${date.getUTCDate()}/${date.getUTCFullYear()}`
}

export function formatFileDate(date: Date): string {
  return `${date.getUTCMonth() + 1}_${date.getUTCDate()}_${date.getUTCFullYear()}`
}

export function formatHours(value: number): string {
  if (Number.isInteger(value)) {
    return `${value}`
  }

  return value.toFixed(2)
}

export function sanitizeHours(value: number): number {
  if (!Number.isFinite(value) || value < 0) {
    return 0
  }

  return Math.round(value * 100) / 100
}

export function isWeekend(date: Date): boolean {
  const day = date.getUTCDay()
  return day === 0 || day === 6
}

export function getDefaultInvoiceSelection(baseDate: Date = new Date()): DefaultSelection {
  const year = baseDate.getFullYear()
  const monthIndex = baseDate.getMonth()
  const dayOfMonth = baseDate.getDate()

  if (dayOfMonth < 16) {
    return {
      monthValue: formatMonthInputValue(year, monthIndex),
      runDay: '16',
    }
  }

  const nextMonthIndex = monthIndex === 11 ? 0 : monthIndex + 1
  const nextMonthYear = monthIndex === 11 ? year + 1 : year

  return {
    monthValue: formatMonthInputValue(nextMonthYear, nextMonthIndex),
    runDay: '1',
  }
}

export function parseMonthValue(monthValue: string): { year: number; monthIndex: number } {
  const [yearText, monthText] = monthValue.split('-')
  const year = Number(yearText)
  const monthIndex = Number(monthText) - 1

  if (!Number.isInteger(year) || !Number.isInteger(monthIndex) || monthIndex < 0 || monthIndex > 11) {
    const now = new Date()
    return { year: now.getFullYear(), monthIndex: now.getMonth() }
  }

  return { year, monthIndex }
}

export function getPeriod(monthValue: string, runDay: RunDay): Period {
  const { year, monthIndex } = parseMonthValue(monthValue)

  if (runDay === '16') {
    return {
      invoiceDate: createUtcDate(year, monthIndex, 16),
      startDate: createUtcDate(year, monthIndex, 1),
      endDate: createUtcDate(year, monthIndex, 15),
    }
  }

  const previousMonthIndex = monthIndex === 0 ? 11 : monthIndex - 1
  const previousMonthYear = monthIndex === 0 ? year - 1 : year

  return {
    invoiceDate: createUtcDate(year, monthIndex, 1),
    startDate: createUtcDate(previousMonthYear, previousMonthIndex, 16),
    endDate: createUtcDate(year, monthIndex, 0),
  }
}

export function listPeriodDays(startDate: Date, endDate: Date): Date[] {
  const days: Date[] = []
  let cursor = startDate

  while (cursor.getTime() <= endDate.getTime()) {
    days.push(cursor)
    cursor = addUtcDays(cursor, 1)
  }

  return days
}

export function buildDayRecords(periodDays: Date[], overrides: Record<string, DayOverride>): DayRecord[] {
  return periodDays.map((date) => {
    const iso = toIsoDate(date)

    if (isWeekend(date)) {
      return {
        iso,
        date,
        status: 'WEEKEND',
        hours: 0,
      }
    }

    const override = overrides[iso]

    return {
      iso,
      date,
      status: override?.status ?? 'WORK',
      hours: override?.hours ?? DEFAULT_HOURS,
    }
  })
}

export function calculateTotalWorkHours(days: DayRecord[]): number {
  return days.reduce((total, day) => {
    if (day.status !== 'WORK') {
      return total
    }

    return total + sanitizeHours(day.hours)
  }, 0)
}

function isSameUtcDate(a: Date, b: Date): boolean {
  return (
    a.getUTCFullYear() === b.getUTCFullYear() &&
    a.getUTCMonth() === b.getUTCMonth() &&
    a.getUTCDate() === b.getUTCDate()
  )
}

export function buildSheetEntries(days: DayRecord[]): SheetEntry[] {
  const entries: SheetEntry[] = []

  for (let index = 0; index < days.length; index += 1) {
    const current = days[index]

    if (current.status === 'WORK') {
      entries.push({
        label: formatDate(current.date),
        hours: sanitizeHours(current.hours),
      })
      continue
    }

    let rangeEnd = current.date

    if (shouldCollapseNonWorkRange(current.status)) {
      while (index + 1 < days.length) {
        const next = days[index + 1]
        const isConsecutive = next.date.getTime() - rangeEnd.getTime() === DAY_MS
        if (next.status !== current.status || !isConsecutive) {
          break
        }

        rangeEnd = next.date
        index += 1
      }
    }

    const dateLabel = isSameUtcDate(current.date, rangeEnd)
      ? formatDate(current.date)
      : `${formatDate(current.date)} - ${formatDate(rangeEnd)}`

    entries.push({
      label: `${dateLabel} (${STATUS_LABEL[current.status]})`,
      hours: null,
    })
  }

  return entries
}

export function sanitizeFileName(value: string): string {
  const cleaned = value.replace(/[<>:"/\\|?*]/g, '').trim()
  return cleaned || 'Employee'
}

export function buildInvoiceFileName(personName: string, invoiceDate: Date): string {
  const safeName = sanitizeFileName(personName)
  return `${safeName} Time Tracking - ${formatFileDate(invoiceDate)}.xlsx`
}

export async function generateWorkbookFromTemplate(params: {
  template: ArrayBuffer
  personName: string
  periodLabel: string
  entries: SheetEntry[]
  totalHours: number
}): Promise<ArrayBuffer> {
  const workbook = new Workbook()
  await workbook.xlsx.load(params.template)

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
