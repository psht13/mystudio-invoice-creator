import { formatFileDate } from './dates'

export function sanitizeFileName(value: string): string {
  const cleaned = value.replace(/[<>:"/\\|?*]/g, '').trim()
  return cleaned || 'Employee'
}

export function buildInvoiceFileName(personName: string, invoiceDate: Date): string {
  const safeName = sanitizeFileName(personName)
  return `${safeName} Time Tracking - ${formatFileDate(invoiceDate)}.xlsx`
}
