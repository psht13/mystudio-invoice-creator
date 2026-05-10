import type { RunDay } from '../../../domain/invoice'

type InvoiceControlsProps = {
  personName: string
  monthValue: string
  runDay: RunDay
  onPersonNameChange: (value: string) => void
  onMonthValueChange: (value: string) => void
  onRunDayChange: (value: RunDay) => void
}

export function InvoiceControls({
  personName,
  monthValue,
  runDay,
  onPersonNameChange,
  onMonthValueChange,
  onRunDayChange,
}: InvoiceControlsProps) {
  return (
    <section className="controls-grid">
      <label className="field">
        <span>Name</span>
        <input
          type="text"
          value={personName}
          onChange={(event) => onPersonNameChange(event.target.value)}
          placeholder="Pavlo Yurchenko"
        />
      </label>

      <label className="field">
        <span>Invoice month</span>
        <input type="month" value={monthValue} onChange={(event) => onMonthValueChange(event.target.value)} />
      </label>

      <div className="field">
        <span>Invoice day</span>
        <div className="segmented">
          <button type="button" className={runDay === '1' ? 'active' : ''} onClick={() => onRunDayChange('1')}>
            1st
          </button>
          <button type="button" className={runDay === '16' ? 'active' : ''} onClick={() => onRunDayChange('16')}>
            16th
          </button>
        </div>
      </div>
    </section>
  )
}
