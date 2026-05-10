import { addUtcDays, createUtcDate, formatMonthInputValue } from './dates'
import type { DefaultSelection, Period, RunDay } from './types'
import { parseRunDayValue, validateMonthValueParsing } from './validation'

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
  const monthValidation = validateMonthValueParsing(monthValue)

  if (!monthValidation.ok) {
    return monthValidation.fallback
  }

  return {
    year: monthValidation.year,
    monthIndex: monthValidation.monthIndex,
  }
}

export function getPeriod(monthValue: string, runDay: RunDay): Period {
  const { year, monthIndex } = parseMonthValue(monthValue)
  const parsedRunDay = parseRunDayValue(runDay)

  if (parsedRunDay === '16') {
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
