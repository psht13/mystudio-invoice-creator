import { useMemo, useState } from 'react'
import {
  InvoiceGenerationService,
  type InvoiceGenerationNotice,
} from '../../../application/services/invoice-generation-service'
import {
  DEFAULT_HOURS,
  buildDayRecords,
  calculateTotalWorkHours,
  getDefaultInvoiceSelection,
  getPeriod,
  listPeriodDays,
  sanitizeHours,
  toIsoDate,
  type DayOverride,
  type EditableStatus,
  type RunDay,
} from '../../../domain/invoice'
import { BrowserFileDownloader } from '../../../infrastructure/browser/browser-file-downloader'
import { FetchTemplateLoader } from '../../../infrastructure/browser/fetch-template-loader'
import { LazyExcelJSWorkbookWriter } from '../../../infrastructure/excel/lazy-exceljs-workbook-writer'

export type InvoiceFormNotice = InvoiceGenerationNotice

const DEFAULT_PERSON_NAME = 'Pavlo Yurchenko'
const invoiceGenerationService = new InvoiceGenerationService({
  templateLoader: new FetchTemplateLoader(),
  workbookWriter: new LazyExcelJSWorkbookWriter(),
  fileDownloader: new BrowserFileDownloader(),
})

function getInitialPersonName() {
  if (typeof window === 'undefined') {
    return DEFAULT_PERSON_NAME
  }

  const requestedName = new URLSearchParams(window.location.search).get('name')?.trim()

  return requestedName || DEFAULT_PERSON_NAME
}

export function useInvoiceForm() {
  const defaultSelection = useMemo(() => getDefaultInvoiceSelection(), [])
  const [personName, setPersonName] = useState(() => getInitialPersonName())
  const [monthValue, setMonthValue] = useState(defaultSelection.monthValue)
  const [runDay, setRunDay] = useState<RunDay>(defaultSelection.runDay)
  const [overrides, setOverrides] = useState<Record<string, DayOverride>>({})
  const [isGenerating, setIsGenerating] = useState(false)
  const [notice, setNotice] = useState<InvoiceFormNotice | null>(null)

  const invoicePreview = useMemo(() => {
    const period = getPeriod(monthValue, runDay)
    const periodDays = listPeriodDays(period.startDate, period.endDate)
    const days = buildDayRecords(periodDays, overrides)
    const totalHours = calculateTotalWorkHours(days)

    return {
      period,
      periodDays,
      days,
      totalHours,
    }
  }, [monthValue, runDay, overrides])

  function changePersonName(value: string) {
    setPersonName(value)
    setNotice(null)
  }

  function changeMonthValue(value: string) {
    setMonthValue(value)
    setNotice(null)
  }

  function changeRunDay(value: RunDay) {
    setRunDay(value)
    setNotice(null)
  }

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

      for (const date of invoicePreview.periodDays) {
        delete nextOverrides[toIsoDate(date)]
      }

      return nextOverrides
    })

    setNotice(null)
  }

  async function handleGenerate() {
    setIsGenerating(true)
    setNotice(null)

    try {
      const result = await invoiceGenerationService.generateInvoice({
        personName,
        monthValue,
        runDay,
        overrides,
      })

      setNotice(result.notice)
    } catch (error) {
      setNotice({
        type: 'error',
        text: error instanceof Error ? error.message : 'Workbook generation failed.',
      })
    } finally {
      setIsGenerating(false)
    }
  }

  return {
    personName,
    monthValue,
    runDay,
    period: invoicePreview.period,
    days: invoicePreview.days,
    totalHours: invoicePreview.totalHours,
    isGenerating,
    notice,
    changePersonName,
    changeMonthValue,
    changeRunDay,
    updateDayStatus,
    updateDayHours,
    resetCurrentPeriod,
    handleGenerate,
  }
}
