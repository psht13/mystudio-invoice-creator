import type { DayStatus } from './types'

export const DEFAULT_HOURS = 8
export const DATA_START_ROW = 15
export const DATA_END_ROW = 28
export const DATA_ROW_CAPACITY = DATA_END_ROW - DATA_START_ROW + 1

export const DAY_MS = 24 * 60 * 60 * 1000
export const TOTAL_FORMULA = 'SUM(E15:E28)'

export const STATUS_LABEL: Record<DayStatus, string> = {
  WORK: 'Work',
  WEEKEND: 'Weekend',
  HOLIDAY: 'Holiday',
  VACATION: 'Vacation',
  OOO: 'OOO',
}
