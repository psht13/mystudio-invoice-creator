import {
  validatePersonNamePresence,
  validateSheetRowCapacity,
  type SheetEntry,
} from '../../domain/invoice'

export type InvoiceGenerationValidationNotice = {
  type: 'error'
  text: string
}

export type InvoicePersonNameValidation =
  | {
      ok: true
      trimmedPersonName: string
    }
  | {
      ok: false
      notice: InvoiceGenerationValidationNotice
    }

export type InvoiceSheetRowCapacityValidation =
  | {
      ok: true
      rowCount: number
      capacity: number
    }
  | {
      ok: false
      notice: InvoiceGenerationValidationNotice
      rowCount: number
      capacity: number
    }

export function validateInvoicePersonName(personName: string): InvoicePersonNameValidation {
  const personNameValidation = validatePersonNamePresence(personName)

  if (!personNameValidation.ok) {
    return {
      ok: false,
      notice: {
        type: 'error',
        text: 'Name is required before generating the invoice.',
      },
    }
  }

  return {
    ok: true,
    trimmedPersonName: personNameValidation.trimmedName,
  }
}

export function validateInvoiceSheetRowCapacity(entries: readonly SheetEntry[]): InvoiceSheetRowCapacityValidation {
  const capacityValidation = validateSheetRowCapacity(entries.length)

  if (!capacityValidation.ok) {
    return {
      ok: false,
      notice: {
        type: 'error',
        text: `Template limit reached: ${capacityValidation.rowCount} rows needed, ${capacityValidation.capacity} rows available.`,
      },
      rowCount: capacityValidation.rowCount,
      capacity: capacityValidation.capacity,
    }
  }

  return {
    ok: true,
    rowCount: capacityValidation.rowCount,
    capacity: capacityValidation.capacity,
  }
}
