import { useMemo, useState } from 'react'
import { Workbook } from 'exceljs'
import './App.css'

type RunDay = '1' | '16'
type EditableStatus = 'WORK' | 'HOLIDAY' | 'VACATION' | 'OOO'
type DayStatus = EditableStatus | 'WEEKEND'

type Notice = {
  type: 'success' | 'error'
  text: string
}

type DayOverride = {
  status: EditableStatus
  hours: number
}

type DayRecord = {
  iso: string
  date: Date
  status: DayStatus
  hours: number
}

type SheetEntry = {
  label: string
  hours: number | null
}

const DAY_MS = 24 * 60 * 60 * 1000
const DEFAULT_HOURS = 8
const DATA_START_ROW = 15
const DATA_END_ROW = 28
const DATA_ROW_CAPACITY = DATA_END_ROW - DATA_START_ROW + 1

const STATUS_LABEL: Record<DayStatus, string> = {
  WORK: 'Work',
  WEEKEND: 'Weekend',
  HOLIDAY: 'Holiday',
  VACATION: 'Vacation',
  OOO: 'OOO',
}

const EDITABLE_STATUS_OPTIONS: EditableStatus[] = ['WORK', 'HOLIDAY', 'VACATION', 'OOO']

function createUtcDate(year: number, monthIndex: number, day: number): Date {
  return new Date(Date.UTC(year, monthIndex, day))
}

function addUtcDays(date: Date, days: number): Date {
  return new Date(date.getTime() + days * DAY_MS)
}

function toIsoDate(date: Date): string {
  const month = String(date.getUTCMonth() + 1).padStart(2, '0')
  const day = String(date.getUTCDate()).padStart(2, '0')
  return `${date.getUTCFullYear()}-${month}-${day}`
}

function isSameUtcDate(a: Date, b: Date): boolean {
  return (
    a.getUTCFullYear() === b.getUTCFullYear() &&
    a.getUTCMonth() === b.getUTCMonth() &&
    a.getUTCDate() === b.getUTCDate()
  )
}

function formatDate(date: Date): string {
  return `${date.getUTCMonth() + 1}/${date.getUTCDate()}/${date.getUTCFullYear()}`
}

function formatFileDate(date: Date): string {
  return `${date.getUTCMonth() + 1}_${date.getUTCDate()}_${date.getUTCFullYear()}`
}

function formatHours(value: number): string {
  if (Number.isInteger(value)) {
    return `${value}`
  }

  return value.toFixed(2)
}

function sanitizeHours(value: number): number {
  if (!Number.isFinite(value) || value < 0) {
    return 0
  }

  return Math.round(value * 100) / 100
}

function isWeekend(date: Date): boolean {
  const day = date.getUTCDay()
  return day === 0 || day === 6
}

function getCurrentMonthValue(): string {
  const today = new Date()
  const month = String(today.getMonth() + 1).padStart(2, '0')
  return `${today.getFullYear()}-${month}`
}

function parseMonthValue(monthValue: string): { year: number; monthIndex: number } {
  const [yearText, monthText] = monthValue.split('-')
  const year = Number(yearText)
  const monthIndex = Number(monthText) - 1

  if (!Number.isInteger(year) || !Number.isInteger(monthIndex) || monthIndex < 0 || monthIndex > 11) {
    const now = new Date()
    return { year: now.getFullYear(), monthIndex: now.getMonth() }
  }

  return { year, monthIndex }
}

function getPeriod(monthValue: string, runDay: RunDay): {
  invoiceDate: Date
  startDate: Date
  endDate: Date
} {
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

function listPeriodDays(startDate: Date, endDate: Date): Date[] {
  const days: Date[] = []
  let cursor = startDate

  while (cursor.getTime() <= endDate.getTime()) {
    days.push(cursor)
    cursor = addUtcDays(cursor, 1)
  }

  return days
}

function buildSheetEntries(days: DayRecord[]): SheetEntry[] {
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

    while (index + 1 < days.length) {
      const next = days[index + 1]
      const isConsecutive = next.date.getTime() - rangeEnd.getTime() === DAY_MS
      if (next.status !== current.status || !isConsecutive) {
        break
      }

      rangeEnd = next.date
      index += 1
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

function sanitizeFileName(value: string): string {
  const cleaned = value.replace(/[<>:"/\\|?*]/g, '').trim()
  return cleaned || 'Employee'
}

async function generateWorkbook(params: {
  runDay: RunDay
  personName: string
  periodLabel: string
  entries: SheetEntry[]
}): Promise<BlobPart> {
  const templatePath = params.runDay === '1' ? '/templates/template-1st.xlsx' : '/templates/template-16th.xlsx'
  const templateResponse = await fetch(templatePath)

  if (!templateResponse.ok) {
    throw new Error(`Failed to load template file: ${templatePath}`)
  }

  const workbook = new Workbook()
  await workbook.xlsx.load(await templateResponse.arrayBuffer())

  const worksheet = workbook.getWorksheet('Invoice') ?? workbook.worksheets[0]
  if (!worksheet) {
    throw new Error('Template worksheet is missing')
  }

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

  return (await workbook.xlsx.writeBuffer()) as BlobPart
}

function App() {
  const [personName, setPersonName] = useState('Pavlo Yurchenko')
  const [monthValue, setMonthValue] = useState(getCurrentMonthValue)
  const [runDay, setRunDay] = useState<RunDay>('16')
  const [overrides, setOverrides] = useState<Record<string, DayOverride>>({})
  const [isGenerating, setIsGenerating] = useState(false)
  const [notice, setNotice] = useState<Notice | null>(null)

  const period = useMemo(() => getPeriod(monthValue, runDay), [monthValue, runDay])

  const periodDays = useMemo(
    () => listPeriodDays(period.startDate, period.endDate),
    [period.startDate, period.endDate],
  )

  const days = useMemo<DayRecord[]>(() => {
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
  }, [periodDays, overrides])

  const totalHours = useMemo(() => {
    return days.reduce((total, day) => {
      if (day.status !== 'WORK') {
        return total
      }

      return total + sanitizeHours(day.hours)
    }, 0)
  }, [days])

  function updateDayStatus(iso: string, status: EditableStatus) {
    setOverrides((current) => {
      const existing = current[iso]
      return {
        ...current,
        [iso]: {
          status,
          hours: existing?.hours ?? DEFAULT_HOURS,
        },
      }
    })
  }

  function updateDayHours(iso: string, value: string) {
    const parsed = Number(value.replace(',', '.'))
    const nextHours = sanitizeHours(parsed)

    setOverrides((current) => {
      const existing = current[iso]
      return {
        ...current,
        [iso]: {
          status: existing?.status ?? 'WORK',
          hours: nextHours,
        },
      }
    })
  }

  function resetCurrentPeriod() {
    setOverrides((current) => {
      const nextOverrides = { ...current }

      for (const date of periodDays) {
        delete nextOverrides[toIsoDate(date)]
      }

      return nextOverrides
    })

    setNotice(null)
  }

  async function handleGenerate() {
    if (!personName.trim()) {
      setNotice({
        type: 'error',
        text: 'Name is required before generating the invoice.',
      })
      return
    }

    const entries = buildSheetEntries(days)

    if (entries.length > DATA_ROW_CAPACITY) {
      setNotice({
        type: 'error',
        text: `Template limit reached: ${entries.length} rows needed, ${DATA_ROW_CAPACITY} rows available.`,
      })
      return
    }

    setIsGenerating(true)
    setNotice(null)

    try {
      const periodLabel = `${formatDate(period.startDate)} - ${formatDate(period.endDate)}`
      const safeName = sanitizeFileName(personName)
      const fileName = `${safeName} Time Tracking - ${formatFileDate(period.invoiceDate)}.xlsx`

      const workbookBuffer = await generateWorkbook({
        runDay,
        personName: personName.trim(),
        periodLabel,
        entries,
      })

      const blob = new Blob([workbookBuffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      })

      const url = URL.createObjectURL(blob)
      const link = document.createElement('a')
      link.href = url
      link.download = fileName
      document.body.appendChild(link)
      link.click()
      link.remove()
      URL.revokeObjectURL(url)

      setNotice({
        type: 'success',
        text: `Downloaded ${fileName}`,
      })
    } catch (error) {
      setNotice({
        type: 'error',
        text: error instanceof Error ? error.message : 'Workbook generation failed.',
      })
    } finally {
      setIsGenerating(false)
    }
  }

  return (
    <div className="app-shell">
      <div className="ambient ambient-one" aria-hidden="true" />
      <div className="ambient ambient-two" aria-hidden="true" />

      <main className="panel">
        <header className="headline">
          <p className="kicker">Invoice Builder</p>
          <h1>Automatic Time-Tracking Excel Generator</h1>
          <p>
            Uses your original spreadsheet templates and keeps the same workbook format while auto-filling
            weekends, work hours, and custom day statuses.
          </p>
        </header>

        <section className="controls-grid">
          <label className="field">
            <span>Name</span>
            <input
              type="text"
              value={personName}
              onChange={(event) => {
                setPersonName(event.target.value)
                setNotice(null)
              }}
              placeholder="Pavlo Yurchenko"
            />
          </label>

          <label className="field">
            <span>Invoice month</span>
            <input
              type="month"
              value={monthValue}
              onChange={(event) => {
                setMonthValue(event.target.value)
                setNotice(null)
              }}
            />
          </label>

          <div className="field">
            <span>Invoice day</span>
            <div className="segmented">
              <button
                type="button"
                className={runDay === '1' ? 'active' : ''}
                onClick={() => {
                  setRunDay('1')
                  setNotice(null)
                }}
              >
                1st
              </button>
              <button
                type="button"
                className={runDay === '16' ? 'active' : ''}
                onClick={() => {
                  setRunDay('16')
                  setNotice(null)
                }}
              >
                16th
              </button>
            </div>
          </div>
        </section>

        <section className="summary-grid">
          <article className="summary-card">
            <h2>Invoice date</h2>
            <p>{formatDate(period.invoiceDate)}</p>
          </article>
          <article className="summary-card">
            <h2>Pay period</h2>
            <p>
              {formatDate(period.startDate)} - {formatDate(period.endDate)}
            </p>
          </article>
          <article className="summary-card">
            <h2>Total work hours</h2>
            <p>{formatHours(totalHours)}</p>
          </article>
        </section>

        <section className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Date</th>
                <th>Status</th>
                <th>Hours</th>
              </tr>
            </thead>
            <tbody>
              {days.map((day) => (
                <tr key={day.iso}>
                  <td>{formatDate(day.date)}</td>
                  <td>
                    {day.status === 'WEEKEND' ? (
                      <span className="status-badge weekend">Weekend</span>
                    ) : (
                      <select
                        value={day.status}
                        onChange={(event) => updateDayStatus(day.iso, event.target.value as EditableStatus)}
                      >
                        {EDITABLE_STATUS_OPTIONS.map((status) => (
                          <option key={status} value={status}>
                            {STATUS_LABEL[status]}
                          </option>
                        ))}
                      </select>
                    )}
                  </td>
                  <td>
                    {day.status === 'WORK' ? (
                      <input
                        className="hours-input"
                        type="number"
                        min="0"
                        step="0.25"
                        value={day.hours}
                        onChange={(event) => updateDayHours(day.iso, event.target.value)}
                      />
                    ) : (
                      <span className="hours-placeholder">-</span>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </section>

        <footer className="actions">
          <button type="button" className="ghost" onClick={resetCurrentPeriod}>
            Reset period
          </button>
          <button type="button" className="primary" onClick={handleGenerate} disabled={isGenerating}>
            {isGenerating ? 'Generating...' : 'Generate & Download XLSX'}
          </button>
        </footer>

        {notice && <p className={`notice ${notice.type}`}>{notice.text}</p>}
      </main>
    </div>
  )
}

export default App
