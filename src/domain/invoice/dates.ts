import { DAY_MS } from './constants'

export function createUtcDate(year: number, monthIndex: number, day: number): Date {
  return new Date(Date.UTC(year, monthIndex, day))
}

export function addUtcDays(date: Date, days: number): Date {
  return new Date(date.getTime() + days * DAY_MS)
}

export function formatMonthInputValue(year: number, monthIndex: number): string {
  const month = String(monthIndex + 1).padStart(2, '0')
  return `${year}-${month}`
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

export function isWeekend(date: Date): boolean {
  const day = date.getUTCDay()
  return day === 0 || day === 6
}

export function isSameUtcDate(a: Date, b: Date): boolean {
  return (
    a.getUTCFullYear() === b.getUTCFullYear() &&
    a.getUTCMonth() === b.getUTCMonth() &&
    a.getUTCDate() === b.getUTCDate()
  )
}
