import { DAY_MS, STATUS_LABEL } from './constants'
import { formatDate, isSameUtcDate } from './dates'
import { sanitizeHours } from './hours'
import type { DayRecord, DayStatus, SheetEntry } from './types'

export type SheetEntrySummary = {
  entries: SheetEntry[]
  totalHours: number
}

function shouldCollapseNonWorkRange(status: DayStatus): boolean {
  return status === 'WEEKEND' || status === 'HOLIDAY' || status === 'VACATION'
}

export function buildSheetEntriesWithTotal(days: DayRecord[]): SheetEntrySummary {
  const entries: SheetEntry[] = []
  let totalHours = 0

  for (let index = 0; index < days.length; index += 1) {
    const current = days[index]

    if (current.status === 'WORK') {
      const hours = sanitizeHours(current.hours)
      totalHours += hours

      entries.push({
        label: formatDate(current.date),
        hours,
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

  return {
    entries,
    totalHours,
  }
}

export function buildSheetEntries(days: DayRecord[]): SheetEntry[] {
  return buildSheetEntriesWithTotal(days).entries
}
