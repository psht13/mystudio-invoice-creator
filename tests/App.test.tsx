import { cleanup, fireEvent, render, screen, waitFor, within } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { createElement } from 'react'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

const workbookWriterMock = vi.hoisted(() => ({
  constructor: vi.fn(),
  generateWorkbookFromTemplate: vi.fn(async () => new ArrayBuffer(8)),
}))

vi.mock('../src/infrastructure/excel/exceljs-workbook-writer', () => {
  class MockExcelJSWorkbookWriter {
    constructor() {
      workbookWriterMock.constructor()
    }

    generateWorkbookFromTemplate = workbookWriterMock.generateWorkbookFromTemplate
  }

  return {
    ExcelJSWorkbookWriter: MockExcelJSWorkbookWriter,
  }
})

import App from '../src/adapters/react/App'

let fetchMock: ReturnType<typeof vi.fn>
let createObjectURLMock: ReturnType<typeof vi.fn>
let revokeObjectURLMock: ReturnType<typeof vi.fn>
let linkClickMock: ReturnType<typeof vi.spyOn>

function renderApp(search = '') {
  window.history.pushState({}, '', `/${search}`)
  return render(createElement(App))
}

function setupUser() {
  return userEvent.setup()
}

function nameInput() {
  return screen.getByLabelText('Name') as HTMLInputElement
}

function monthInput() {
  return screen.getByLabelText('Invoice month') as HTMLInputElement
}

function getSummaryCard(label: string) {
  const heading = screen.getByRole('heading', { name: label })
  const card = heading.closest('article')

  if (!(card instanceof HTMLElement)) {
    throw new Error(`Summary card not found: ${label}`)
  }

  return card
}

function expectSummaryValue(label: string, value: string) {
  expect(within(getSummaryCard(label)).getByText(value)).toBeInTheDocument()
}

function tableRowForDate(dateLabel: string) {
  const dateCell = within(screen.getByRole('table')).getByText(dateLabel)
  const row = dateCell.closest('tr')

  if (!(row instanceof HTMLTableRowElement)) {
    throw new Error(`Table row not found: ${dateLabel}`)
  }

  return row
}

function statusSelectForDate(dateLabel: string) {
  return within(tableRowForDate(dateLabel)).getByRole('combobox') as HTMLSelectElement
}

function hoursInputForDate(dateLabel: string) {
  return within(tableRowForDate(dateLabel)).getByRole('spinbutton') as HTMLInputElement
}

function setInvoiceMonth(month: string) {
  fireEvent.change(monthInput(), { target: { value: month } })
}

async function setPeriod(month: string, runDay: '1' | '16') {
  setInvoiceMonth(month)
  await setupUser().click(screen.getByRole('button', { name: runDay === '1' ? '1st' : '16th' }))
}

function templateResponse() {
  return new Response(new ArrayBuffer(8), { status: 200 })
}

beforeEach(() => {
  window.history.pushState({}, '', '/')

  fetchMock = vi.fn()
  createObjectURLMock = vi.fn(() => 'blob:invoice')
  revokeObjectURLMock = vi.fn()
  linkClickMock = vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => undefined)

  vi.stubGlobal('fetch', fetchMock)
  Object.defineProperty(URL, 'createObjectURL', {
    configurable: true,
    value: createObjectURLMock,
  })
  Object.defineProperty(URL, 'revokeObjectURL', {
    configurable: true,
    value: revokeObjectURLMock,
  })

  workbookWriterMock.generateWorkbookFromTemplate.mockReset()
  workbookWriterMock.generateWorkbookFromTemplate.mockResolvedValue(new ArrayBuffer(8))
  workbookWriterMock.constructor.mockReset()
})

afterEach(() => {
  cleanup()
  vi.unstubAllGlobals()
  vi.restoreAllMocks()
})

describe('App defaults and period controls', () => {
  it('uses Pavlo Yurchenko as the default name', () => {
    renderApp()

    expect(nameInput()).toHaveValue('Pavlo Yurchenko')
  })

  it('prefills the name from the name query parameter', () => {
    renderApp('?name=Jane%20Doe')

    expect(nameInput()).toHaveValue('Jane Doe')
  })

  it('falls back to the default name when the name query parameter is blank', () => {
    renderApp('?name=%20%20%20')

    expect(nameInput()).toHaveValue('Pavlo Yurchenko')
  })

  it('updates summary cards when the invoice month changes', async () => {
    renderApp()

    await setPeriod('2026-04', '16')

    expectSummaryValue('Invoice date', '4/16/2026')
    expectSummaryValue('Pay period', '4/1/2026 - 4/15/2026')
  })

  it('updates the invoice date and pay period when run day buttons change', async () => {
    const user = setupUser()
    renderApp()

    setInvoiceMonth('2026-03')
    await user.click(screen.getByRole('button', { name: '1st' }))

    expectSummaryValue('Invoice date', '3/1/2026')
    expectSummaryValue('Pay period', '2/16/2026 - 2/28/2026')

    await user.click(screen.getByRole('button', { name: '16th' }))

    expectSummaryValue('Invoice date', '3/16/2026')
    expectSummaryValue('Pay period', '3/1/2026 - 3/15/2026')
  })
})

describe('App editable day behavior', () => {
  it('changes a workday status to holiday, vacation, and OOO', async () => {
    const user = setupUser()
    renderApp()
    await setPeriod('2026-02', '16')

    await user.selectOptions(statusSelectForDate('2/2/2026'), 'HOLIDAY')
    expect(statusSelectForDate('2/2/2026')).toHaveDisplayValue('Holiday')
    expect(within(tableRowForDate('2/2/2026')).getByText('-')).toBeInTheDocument()
    expectSummaryValue('Total work hours', '72')

    await user.selectOptions(statusSelectForDate('2/2/2026'), 'VACATION')
    expect(statusSelectForDate('2/2/2026')).toHaveDisplayValue('Vacation')

    await user.selectOptions(statusSelectForDate('2/2/2026'), 'OOO')
    expect(statusSelectForDate('2/2/2026')).toHaveDisplayValue('OOO')
  })

  it('updates total work hours when a workday hours input changes', async () => {
    renderApp()
    await setPeriod('2026-02', '16')

    fireEvent.change(hoursInputForDate('2/2/2026'), { target: { value: '6.5' } })

    expect(hoursInputForDate('2/2/2026')).toHaveValue(6.5)
    expectSummaryValue('Total work hours', '78.50')
  })

  it('sanitizes negative work hours to 0', async () => {
    renderApp()
    await setPeriod('2026-02', '16')

    fireEvent.change(hoursInputForDate('2/2/2026'), { target: { value: '-2' } })

    expect(hoursInputForDate('2/2/2026')).toHaveValue(0)
    expectSummaryValue('Total work hours', '72')
  })

  it('resets overrides only for the currently selected period', async () => {
    const user = setupUser()
    renderApp()
    await setPeriod('2026-02', '16')
    fireEvent.change(hoursInputForDate('2/2/2026'), { target: { value: '6' } })
    expectSummaryValue('Total work hours', '78')

    setInvoiceMonth('2026-03')
    await user.click(screen.getByRole('button', { name: '16th' }))
    fireEvent.change(hoursInputForDate('3/2/2026'), { target: { value: '5' } })
    expectSummaryValue('Total work hours', '77')

    await user.click(screen.getByRole('button', { name: 'Reset period' }))

    expect(hoursInputForDate('3/2/2026')).toHaveValue(8)
    expectSummaryValue('Total work hours', '80')

    setInvoiceMonth('2026-02')
    await user.click(screen.getByRole('button', { name: '16th' }))

    expect(hoursInputForDate('2/2/2026')).toHaveValue(6)
    expectSummaryValue('Total work hours', '78')
  })
})

describe('App generation behavior', () => {
  it('blocks generation with the existing error when name is empty', async () => {
    const user = setupUser()
    renderApp()

    await user.clear(nameInput())
    await user.click(screen.getByRole('button', { name: 'Generate & Download XLSX' }))

    expect(screen.getByText('Name is required before generating the invoice.')).toBeInTheDocument()
    expect(fetchMock).not.toHaveBeenCalled()
  })

  it('shows Generating... while generation is in progress and triggers the download behavior', async () => {
    const user = setupUser()
    let resolveFetch: (response: Response) => void
    fetchMock.mockReturnValue(
      new Promise<Response>((resolve) => {
        resolveFetch = resolve
      }),
    )
    renderApp()
    await setPeriod('2026-02', '16')

    expect(workbookWriterMock.constructor).not.toHaveBeenCalled()

    await user.click(screen.getByRole('button', { name: 'Generate & Download XLSX' }))

    expect(screen.getByRole('button', { name: 'Generating...' })).toBeDisabled()
    expect(workbookWriterMock.constructor).not.toHaveBeenCalled()

    resolveFetch(templateResponse())

    expect(await screen.findByText('Downloaded Pavlo Yurchenko Time Tracking - 2_16_2026.xlsx')).toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Generate & Download XLSX' })).toBeEnabled()
    expect(fetchMock).toHaveBeenCalledWith('/templates/template-16th.xlsx')
    expect(workbookWriterMock.constructor).toHaveBeenCalledTimes(1)
    expect(workbookWriterMock.generateWorkbookFromTemplate).toHaveBeenCalledTimes(1)
    expect(createObjectURLMock).toHaveBeenCalledTimes(1)
    expect(linkClickMock).toHaveBeenCalledTimes(1)
    expect(revokeObjectURLMock).toHaveBeenCalledWith('blob:invoice')
  })

  it('keeps the loading state until the lazy workbook writer finishes', async () => {
    const user = setupUser()
    let resolveWorkbook: (buffer: ArrayBuffer) => void
    fetchMock.mockResolvedValue(templateResponse())
    workbookWriterMock.generateWorkbookFromTemplate.mockReturnValue(
      new Promise<ArrayBuffer>((resolve) => {
        resolveWorkbook = resolve
      }),
    )
    renderApp()
    await setPeriod('2026-02', '16')

    await user.click(screen.getByRole('button', { name: 'Generate & Download XLSX' }))

    await waitFor(() => {
      expect(workbookWriterMock.generateWorkbookFromTemplate).toHaveBeenCalledTimes(1)
    })
    expect(screen.getByRole('button', { name: 'Generating...' })).toBeDisabled()

    resolveWorkbook(new ArrayBuffer(8))

    expect(await screen.findByText('Downloaded Pavlo Yurchenko Time Tracking - 2_16_2026.xlsx')).toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Generate & Download XLSX' })).toBeEnabled()
  })

  it('shows the existing template fetch failure error behavior', async () => {
    const user = setupUser()
    fetchMock.mockResolvedValue(new Response(null, { status: 404 }))
    renderApp()
    await setPeriod('2026-02', '16')

    await user.click(screen.getByRole('button', { name: 'Generate & Download XLSX' }))

    await waitFor(() => {
      expect(screen.getByText('Failed to load template file: /templates/template-16th.xlsx')).toBeInTheDocument()
    })
    expect(linkClickMock).not.toHaveBeenCalled()
    expect(createObjectURLMock).not.toHaveBeenCalled()
  })

  it('shows the existing workbook writer failure error behavior', async () => {
    const user = setupUser()
    fetchMock.mockResolvedValue(templateResponse())
    workbookWriterMock.generateWorkbookFromTemplate.mockRejectedValue(new Error('Workbook writer failed.'))
    renderApp()
    await setPeriod('2026-02', '16')

    await user.click(screen.getByRole('button', { name: 'Generate & Download XLSX' }))

    await waitFor(() => {
      expect(screen.getByText('Workbook writer failed.')).toBeInTheDocument()
    })
    expect(screen.getByRole('button', { name: 'Generate & Download XLSX' })).toBeEnabled()
    expect(linkClickMock).not.toHaveBeenCalled()
    expect(createObjectURLMock).not.toHaveBeenCalled()
  })
})
