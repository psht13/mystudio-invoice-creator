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
