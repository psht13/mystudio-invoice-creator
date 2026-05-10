import type { DayRecord } from './types'
import { validateHoursSanitization } from './validation'

export function formatHours(value: number): string {
  if (Number.isInteger(value)) {
    return `${value}`
  }

  return value.toFixed(2)
}

export function sanitizeHours(value: number): number {
  return validateHoursSanitization(value).value
}

export function calculateTotalWorkHours(days: DayRecord[]): number {
  return days.reduce((total, day) => {
    if (day.status !== 'WORK') {
      return total
    }

    return total + sanitizeHours(day.hours)
  }, 0)
}
