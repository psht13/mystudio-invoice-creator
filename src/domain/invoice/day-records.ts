import { DEFAULT_HOURS } from './constants'
import { isWeekend, toIsoDate } from './dates'
import type { DayOverride, DayRecord } from './types'

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
