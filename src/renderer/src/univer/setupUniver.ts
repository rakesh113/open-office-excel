import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import UniverPresetSheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US'
import { UniverSheetsDataValidationPreset } from '@univerjs/preset-sheets-data-validation'
import UniverPresetSheetsDataValidationEnUS from '@univerjs/preset-sheets-data-validation/locales/en-US'
import { UniverSheetsFilterPreset } from '@univerjs/preset-sheets-filter'
import UniverPresetSheetsFilterEnUS from '@univerjs/preset-sheets-filter/locales/en-US'
import { createUniver, LocaleType, mergeLocales } from '@univerjs/presets'
import type { IWorkbookData } from '../../../shared/workbook'
import { applyDatePickerToSelection, clearDatePickerFromSelection } from './datePicker'
import {
  freezeHeaderRows,
  freezeRowsAboveSelection,
  installFreezeCommandFix,
  removeFilter,
  toggleFilterOnSelection,
  unfreezePanes
} from './sheetActions'

import '@univerjs/preset-sheets-core/lib/index.css'
import '@univerjs/preset-sheets-data-validation/lib/index.css'
import '@univerjs/preset-sheets-filter/lib/index.css'

export type UniverAPI = ReturnType<typeof createUniver>['univerAPI']

export type SheetActionResult = { ok: true } | { ok: false; error: string }

export function mountUniver(container: HTMLElement): {
  univerAPI: UniverAPI
  createOrReplaceWorkbook: (data: IWorkbookData) => void
  getSnapshot: () => IWorkbookData | null
  applyDatePicker: () => SheetActionResult
  clearDatePicker: () => SheetActionResult
  toggleFilter: () => SheetActionResult
  clearFilter: () => SheetActionResult
  freezeHeader: () => SheetActionResult
  freezeAboveSelection: () => SheetActionResult
  unfreeze: () => SheetActionResult
  onCommandExecuted: (cb: () => void) => () => void
  dispose: () => void
} {
  const { univerAPI } = createUniver({
    locale: LocaleType.EN_US,
    locales: {
      [LocaleType.EN_US]: mergeLocales(
        UniverPresetSheetsCoreEnUS,
        UniverPresetSheetsDataValidationEnUS,
        UniverPresetSheetsFilterEnUS
      )
    },
    presets: [
      UniverSheetsCorePreset({
        container,
        toolbar: true,
        formulaBar: true,
        footer: {
          sheetBar: true,
          statisticBar: true,
          menus: true,
          zoomSlider: true
        },
        contextMenu: true
      }),
      UniverSheetsDataValidationPreset({
        showEditOnDropdown: true
      }),
      UniverSheetsFilterPreset()
    ]
  })

  const createOrReplaceWorkbook = (data: IWorkbookData): void => {
    const existing = univerAPI.getActiveWorkbook()
    if (existing) {
      try {
        const unitId = existing.getId()
        univerAPI.disposeUnit(unitId)
      } catch {
        // ignore dispose errors and create on top
      }
    }
    univerAPI.createWorkbook(data as never)
  }

  const getSnapshot = (): IWorkbookData | null => {
    const wb = univerAPI.getActiveWorkbook()
    if (!wb) return null
    return wb.save() as IWorkbookData
  }

  const uninstallFreezeFix = installFreezeCommandFix(univerAPI)

  const onCommandExecuted = (cb: () => void): (() => void) => {
    const disposable = univerAPI.onCommandExecuted((command) => {
      const id = String((command as { id?: string }).id ?? '')
      if (
        /selection|scroll|focus|activate|hover|skeleton|viewport|render|lifecycle/i.test(id)
      ) {
        return
      }
      cb()
    })
    return () => disposable.dispose()
  }

  return {
    univerAPI,
    createOrReplaceWorkbook,
    getSnapshot,
    applyDatePicker: () => applyDatePickerToSelection(univerAPI),
    clearDatePicker: () => clearDatePickerFromSelection(univerAPI),
    toggleFilter: () => toggleFilterOnSelection(univerAPI),
    clearFilter: () => removeFilter(univerAPI),
    freezeHeader: () => freezeHeaderRows(univerAPI, 1),
    freezeAboveSelection: () => freezeRowsAboveSelection(univerAPI),
    unfreeze: () => unfreezePanes(univerAPI),
    onCommandExecuted,
    dispose: () => {
      uninstallFreezeFix()
      univerAPI.dispose()
    }
  }
}
