import type { InvoiceFormNotice } from '../hooks/useInvoiceForm'

type NoticeProps = {
  notice: InvoiceFormNotice
}

export function Notice({ notice }: NoticeProps) {
  return <p className={`notice ${notice.type}`}>{notice.text}</p>
}
