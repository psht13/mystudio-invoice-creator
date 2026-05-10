import '../../App.css'
import { ActionsBar } from './components/ActionsBar'
import { DayTable } from './components/DayTable'
import { InvoiceControls } from './components/InvoiceControls'
import { Notice } from './components/Notice'
import { SummaryCards } from './components/SummaryCards'
import { useInvoiceForm } from './hooks/useInvoiceForm'

function App() {
  const invoiceForm = useInvoiceForm()

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

        <InvoiceControls
          personName={invoiceForm.personName}
          monthValue={invoiceForm.monthValue}
          runDay={invoiceForm.runDay}
          onPersonNameChange={invoiceForm.changePersonName}
          onMonthValueChange={invoiceForm.changeMonthValue}
          onRunDayChange={invoiceForm.changeRunDay}
        />

        <SummaryCards period={invoiceForm.period} totalHours={invoiceForm.totalHours} />

        <DayTable
          days={invoiceForm.days}
          onDayStatusChange={invoiceForm.updateDayStatus}
          onDayHoursChange={invoiceForm.updateDayHours}
        />

        <ActionsBar
          isGenerating={invoiceForm.isGenerating}
          onResetCurrentPeriod={invoiceForm.resetCurrentPeriod}
          onGenerate={invoiceForm.handleGenerate}
        />

        {invoiceForm.notice && <Notice notice={invoiceForm.notice} />}
      </main>
    </div>
  )
}

export default App
