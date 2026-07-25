import type { UniverAPI } from './setupUniver'

type ActionResult = { ok: true; message?: string } | { ok: false; error: string }

type FilterableRange = {
  createFilter: () => { remove: () => boolean } | null
  getA1Notation?: () => string
}

type CellRange = {
  getRow: () => number
  getColumn: () => number
}

type FreezeSheet = {
  getFilter: () => { remove: () => boolean } | null
  getActiveRange: () => (FilterableRange & Partial<CellRange>) | null
  getDataRange: () => FilterableRange
  setFrozenRows: (rows: number) => unknown
  setFrozenColumns?: (cols: number) => unknown
  setFreeze: (freeze: {
    startRow: number
    startColumn: number
    ySplit: number
    xSplit: number
  }) => unknown
  getFrozenRows?: () => number
  cancelFreeze?: () => unknown
}

function getActiveSheet(univerAPI: UniverAPI): FreezeSheet | null {
  const workbook = univerAPI.getActiveWorkbook()
  if (!workbook) return null
  return workbook.getActiveSheet() as unknown as FreezeSheet
}

function getActiveCell(univerAPI: UniverAPI): { row: number; column: number } | null {
  const workbook = univerAPI.getActiveWorkbook()
  if (!workbook) return null
  const sheet = workbook.getActiveSheet() as unknown as FreezeSheet
  const range =
    sheet.getActiveRange?.() ??
    (workbook.getActiveRange?.() as { getRow: () => number; getColumn: () => number } | null)
  if (!range || typeof range.getRow !== 'function') return null
  return {
    row: range.getRow(),
    column: typeof range.getColumn === 'function' ? range.getColumn() : 0
  }
}

/**
 * Apply a clean row freeze. Avoids Univer's scroll-offset ySplit bug in
 * sheet.command.set-selection-frozen / "Freeze to this row".
 *
 * Freezes rows ABOVE `row` (0-based). The selected row stays the first scrollable row.
 * If row is 0, freezes the first row (header) instead.
 */
export function applyRowFreeze(univerAPI: UniverAPI, row: number): ActionResult {
  try {
    const sheet = getActiveSheet(univerAPI)
    if (!sheet) return { ok: false, error: 'No workbook is open.' }

    const freezeRow = Math.max(0, Math.floor(row))

    if (freezeRow <= 0) {
      // On row 1: freeze that header row (common expectation).
      sheet.setFreeze({
        startRow: 1,
        ySplit: 1,
        startColumn: -1,
        xSplit: 0
      })
      return { ok: true, message: 'Row 1 frozen as header' }
    }

    sheet.setFreeze({
      startRow: freezeRow,
      ySplit: freezeRow,
      startColumn: -1,
      xSplit: 0
    })
    return {
      ok: true,
      message: `Frozen through row ${freezeRow} (rows 1–${freezeRow} stay visible while scrolling)`
    }
  } catch (error) {
    return {
      ok: false,
      error: error instanceof Error ? error.message : String(error)
    }
  }
}

/** Create AutoFilter on the selection, or the sheet data range if selection is tiny. */
export function toggleFilterOnSelection(univerAPI: UniverAPI): ActionResult {
  try {
    const sheet = getActiveSheet(univerAPI)
    if (!sheet) return { ok: false, error: 'No workbook is open.' }

    const existing = sheet.getFilter?.()
    if (existing) {
      existing.remove()
      return { ok: true, message: 'Filter removed' }
    }

    const range: FilterableRange = sheet.getActiveRange() ?? sheet.getDataRange()

    let filter = range.createFilter()
    if (!filter) {
      sheet.getFilter()?.remove()
      filter = range.createFilter()
    }

    if (!filter) {
      return { ok: false, error: 'Could not create a filter on the selection.' }
    }

    return { ok: true, message: 'Filter enabled — use column dropdowns' }
  } catch (error) {
    return {
      ok: false,
      error: error instanceof Error ? error.message : String(error)
    }
  }
}

export function removeFilter(univerAPI: UniverAPI): ActionResult {
  try {
    const sheet = getActiveSheet(univerAPI)
    if (!sheet) return { ok: false, error: 'No workbook is open.' }

    const filter = sheet.getFilter?.()
    if (!filter) {
      return { ok: false, error: 'No filter is active on this sheet.' }
    }
    filter.remove()
    return { ok: true, message: 'Filter removed' }
  } catch (error) {
    return {
      ok: false,
      error: error instanceof Error ? error.message : String(error)
    }
  }
}

/** Freeze the first N rows as a header (default 1). */
export function freezeHeaderRows(univerAPI: UniverAPI, rows = 1): ActionResult {
  try {
    const sheet = getActiveSheet(univerAPI)
    if (!sheet) return { ok: false, error: 'No workbook is open.' }

    const count = Math.max(1, Math.floor(rows))
    sheet.setFreeze({
      startRow: count,
      ySplit: count,
      startColumn: -1,
      xSplit: 0
    })
    return { ok: true, message: count === 1 ? 'First row frozen as header' : `First ${count} rows frozen` }
  } catch (error) {
    return {
      ok: false,
      error: error instanceof Error ? error.message : String(error)
    }
  }
}

/** Freeze rows above the active cell — Excel / “Freeze to this row”. */
export function freezeRowsAboveSelection(univerAPI: UniverAPI): ActionResult {
  const cell = getActiveCell(univerAPI)
  if (!cell) {
    return { ok: false, error: 'Select a cell or row first, then freeze.' }
  }
  return applyRowFreeze(univerAPI, cell.row)
}

export function unfreezePanes(univerAPI: UniverAPI): ActionResult {
  try {
    const sheet = getActiveSheet(univerAPI)
    if (!sheet) return { ok: false, error: 'No workbook is open.' }

    if (typeof sheet.cancelFreeze === 'function') {
      sheet.cancelFreeze()
    } else {
      sheet.setFreeze({
        startRow: -1,
        startColumn: -1,
        ySplit: 0,
        xSplit: 0
      })
    }
    return { ok: true, message: 'Panes unfrozen' }
  } catch (error) {
    return {
      ok: false,
      error: error instanceof Error ? error.message : String(error)
    }
  }
}

/**
 * Intercept Univer's built-in freeze commands (context menu "Freeze to this row")
 * and apply our scroll-safe freeze instead.
 *
 * Univer's built-in path computes ySplit as (row - scrollOffset), which often
 * freezes the wrong rows or appears to do nothing after scrolling.
 */
export function installFreezeCommandFix(univerAPI: UniverAPI): () => void {
  let applyingFix = false

  const disposable = univerAPI.addEvent(univerAPI.Event.BeforeCommandExecute, (event) => {
    if (applyingFix) return

    const id = String(event.id ?? '')
    const params = event.params as { type?: number } | undefined

    const isRowFreeze =
      id === 'sheet.command.set-row-frozen' ||
      (id === 'sheet.command.set-selection-frozen' && params?.type === 1)

    if (!isRowFreeze) return

    const cell = getActiveCell(univerAPI)
    if (!cell) return

    event.cancel = true

    applyingFix = true
    try {
      applyRowFreeze(univerAPI, cell.row)
    } finally {
      applyingFix = false
    }
  })

  return () => disposable.dispose()
}
