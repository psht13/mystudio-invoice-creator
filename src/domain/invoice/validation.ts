import { DATA_ROW_CAPACITY } from './constants'
import type { RunDay } from './types'

export type PersonNamePresenceValidation =
  | {
      ok: true
      trimmedName: string
    }
  | {
      ok: false
      reason: 'missing-person-name'
    }

export type MonthValueParsingValidation =
  | {
      ok: true
      year: number
      monthIndex: number
    }
  | {
      ok: false
      reason: 'invalid-month-value'
      fallback: {
        year: number
        monthIndex: number
      }
    }

export type HoursSanitizationValidation = {
  value: number
  sanitized: boolean
}

export type SheetRowCapacityValidation =
  | {
      ok: true
      rowCount: number
      capacity: number
    }
  | {
      ok: false
      reason: 'sheet-row-capacity-exceeded'
      rowCount: number
      capacity: number
    }

export function validatePersonNamePresence(personName: string): PersonNamePresenceValidation {
  const trimmedName = personName.trim()

  if (!trimmedName) {
    return {
      ok: false,
      reason: 'missing-person-name',
    }
  }

  return {
    ok: true,
    trimmedName,
  }
}

export function isRunDayValue(value: unknown): value is RunDay {
  return value === '1' || value === '16'
}

export function parseRunDayValue(value: unknown): RunDay {
  return value === '16' ? '16' : '1'
}

export function validateMonthValueParsing(monthValue: string, fallbackDate?: Date): MonthValueParsingValidation {
  const [yearText, monthText] = monthValue.split('-')
  const year = Number(yearText)
  const monthIndex = Number(monthText) - 1

  if (!Number.isInteger(year) || !Number.isInteger(monthIndex) || monthIndex < 0 || monthIndex > 11) {
    const fallback = fallbackDate ?? new Date()

    return {
      ok: false,
      reason: 'invalid-month-value',
      fallback: {
        year: fallback.getFullYear(),
        monthIndex: fallback.getMonth(),
      },
    }
  }

  return {
    ok: true,
    year,
    monthIndex,
  }
}

export function validateHoursSanitization(value: number): HoursSanitizationValidation {
  if (!Number.isFinite(value) || value < 0) {
    return {
      value: 0,
      sanitized: value !== 0,
    }
  }

  const sanitizedValue = Math.round(value * 100) / 100

  return {
    value: sanitizedValue,
    sanitized: sanitizedValue !== value,
  }
}

export function validateSheetRowCapacity(
  rowCount: number,
  capacity: number = DATA_ROW_CAPACITY,
): SheetRowCapacityValidation {
  if (rowCount > capacity) {
    return {
      ok: false,
      reason: 'sheet-row-capacity-exceeded',
      rowCount,
      capacity,
    }
  }

  return {
    ok: true,
    rowCount,
    capacity,
  }
}
