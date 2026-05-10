import { STATUS_LABEL, formatDate, type DayRecord, type EditableStatus } from '../../../domain/invoice'

type DayTableProps = {
  days: DayRecord[]
  onDayStatusChange: (iso: string, status: EditableStatus) => void
  onDayHoursChange: (iso: string, value: string) => void
}

const EDITABLE_STATUS_OPTIONS: EditableStatus[] = ['WORK', 'HOLIDAY', 'VACATION', 'OOO']

export function DayTable({ days, onDayStatusChange, onDayHoursChange }: DayTableProps) {
  return (
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
                    onChange={(event) => onDayStatusChange(day.iso, event.target.value as EditableStatus)}
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
                    onChange={(event) => onDayHoursChange(day.iso, event.target.value)}
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
  )
}
