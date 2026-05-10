import { formatDate, formatHours, type Period } from '../../../domain/invoice'

type SummaryCardsProps = {
  period: Period
  totalHours: number
}

export function SummaryCards({ period, totalHours }: SummaryCardsProps) {
  return (
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
  )
}
