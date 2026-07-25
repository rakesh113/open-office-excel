import type { UniverAPI } from './setupUniver'

const DATE_FORMAT = 'yyyy-mm-dd'
const DATE_MIN = new Date(1900, 0, 1)
const DATE_MAX = new Date(2100, 11, 31)

type RangeWithValidation = {
  setDataValidation: (rule: unknown) => unknown
  setNumberFormat?: (pattern: string) => unknown
}

/**
 * Apply Univer date data-validation (calendar dropdown) + date number format
 * to the current selection.
 */
export function applyDatePickerToSelection(univerAPI: UniverAPI): { ok: true } | { ok: false; error: string } {
  try {
    const workbook = univerAPI.getActiveWorkbook()
    if (!workbook) {
      return { ok: false, error: 'No workbook is open.' }
    }

    const sheet = workbook.getActiveSheet()
    const range = (sheet.getActiveRange() ?? workbook.getActiveRange()) as RangeWithValidation | null
    if (!range) {
      return { ok: false, error: 'Select one or more cells first.' }
    }

    if (typeof univerAPI.newDataValidation !== 'function') {
      return { ok: false, error: 'Date validation is not available.' }
    }

    const rule = univerAPI
      .newDataValidation()
      .requireDateBetween(DATE_MIN, DATE_MAX)
      .setOptions({
        showErrorMessage: true,
        error: 'Please pick a valid date.',
        showDropDown: true
      })
      .build()

    range.setDataValidation(rule)

    if (typeof range.setNumberFormat === 'function') {
      range.setNumberFormat(DATE_FORMAT)
    }

    return { ok: true }
  } catch (error) {
    return {
      ok: false,
      error: error instanceof Error ? error.message : String(error)
    }
  }
}

/**
 * Clear date validation from the current selection.
 */
export function clearDatePickerFromSelection(univerAPI: UniverAPI): { ok: true } | { ok: false; error: string } {
  try {
    const workbook = univerAPI.getActiveWorkbook()
    if (!workbook) {
      return { ok: false, error: 'No workbook is open.' }
    }

    const range = (workbook.getActiveSheet().getActiveRange() ??
      workbook.getActiveRange()) as RangeWithValidation | null
    if (!range) {
      return { ok: false, error: 'Select one or more cells first.' }
    }

    range.setDataValidation(null)
    return { ok: true }
  } catch (error) {
    return {
      ok: false,
      error: error instanceof Error ? error.message : String(error)
    }
  }
}
