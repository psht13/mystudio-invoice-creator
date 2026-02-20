import { useMemo, useState } from 'react'
import './App.css'
import {
  DATA_ROW_CAPACITY,
  STATUS_LABEL,
  buildDayRecords,
  buildInvoiceFileName,
  buildSheetEntries,
  calculateTotalWorkHours,
  formatDate,
  formatHours,
  generateWorkbookFromTemplate,
  getDefaultInvoiceSelection,
  getPeriod,
  listPeriodDays,
  sanitizeHours,
  toIsoDate,
  type DayOverride,
  type EditableStatus,
  type RunDay,
} from './invoice'

type Notice = {
  type: 'success' | 'error'
  text: string
}

const DEFAULT_HOURS = 8
const EDITABLE_STATUS_OPTIONS: EditableStatus[] = ['WORK', 'HOLIDAY', 'VACATION', 'OOO']

function App() {
  const defaultSelection = useMemo(() => getDefaultInvoiceSelection(), [])
  const [personName, setPersonName] = useState('Pavlo Yurchenko')
  const [monthValue, setMonthValue] = useState(defaultSelection.monthValue)
  const [runDay, setRunDay] = useState<RunDay>(defaultSelection.runDay)
  const [overrides, setOverrides] = useState<Record<string, DayOverride>>({})
  const [isGenerating, setIsGenerating] = useState(false)
  const [notice, setNotice] = useState<Notice | null>(null)

  const period = useMemo(() => getPeriod(monthValue, runDay), [monthValue, runDay])

  const periodDays = useMemo(
    () => listPeriodDays(period.startDate, period.endDate),
    [period.startDate, period.endDate],
  )

  const days = useMemo(() => buildDayRecords(periodDays, overrides), [periodDays, overrides])

  const totalHours = useMemo(() => calculateTotalWorkHours(days), [days])

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
      const templatePath = runDay === '1' ? '/templates/template-1st.xlsx' : '/templates/template-16th.xlsx'
      const templateResponse = await fetch(templatePath)

      if (!templateResponse.ok) {
        throw new Error(`Failed to load template file: ${templatePath}`)
      }

      const periodLabel = `${formatDate(period.startDate)} - ${formatDate(period.endDate)}`
      const fileName = buildInvoiceFileName(personName, period.invoiceDate)
      const workbookBuffer = await generateWorkbookFromTemplate({
        template: await templateResponse.arrayBuffer(),
        personName: personName.trim(),
        periodLabel,
        entries,
        totalHours,
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
