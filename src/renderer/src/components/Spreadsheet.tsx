import { useEffect, useRef } from 'react'
import type { IWorkbookData } from '../../../shared/workbook'
import { mountUniver, type SheetActionResult, type UniverAPI } from '../univer/setupUniver'

export interface SpreadsheetHandle {
  createOrReplaceWorkbook: (data: IWorkbookData) => void
  getSnapshot: () => IWorkbookData | null
  getAPI: () => UniverAPI | null
  applyDatePicker: () => SheetActionResult
  clearDatePicker: () => SheetActionResult
  toggleFilter: () => SheetActionResult
  clearFilter: () => SheetActionResult
  freezeHeader: () => SheetActionResult
  freezeAboveSelection: () => SheetActionResult
  unfreeze: () => SheetActionResult
}

interface SpreadsheetProps {
  initialData: IWorkbookData
  onReady: (handle: SpreadsheetHandle) => void
  onDirty: () => void
}

export function Spreadsheet({ initialData, onReady, onDirty }: SpreadsheetProps): JSX.Element {
  const containerRef = useRef<HTMLDivElement>(null)
  const handleRef = useRef<SpreadsheetHandle | null>(null)
  const onReadyRef = useRef(onReady)
  const onDirtyRef = useRef(onDirty)
  const initialDataRef = useRef(initialData)

  onReadyRef.current = onReady
  onDirtyRef.current = onDirty

  useEffect(() => {
    const el = containerRef.current
    if (!el) return

    const instance = mountUniver(el)
    instance.createOrReplaceWorkbook(initialDataRef.current)

    const handle: SpreadsheetHandle = {
      createOrReplaceWorkbook: instance.createOrReplaceWorkbook,
      getSnapshot: instance.getSnapshot,
      getAPI: () => instance.univerAPI,
      applyDatePicker: instance.applyDatePicker,
      clearDatePicker: instance.clearDatePicker,
      toggleFilter: instance.toggleFilter,
      clearFilter: instance.clearFilter,
      freezeHeader: instance.freezeHeader,
      freezeAboveSelection: instance.freezeAboveSelection,
      unfreeze: instance.unfreeze
    }
    handleRef.current = handle
    onReadyRef.current(handle)

    const unsubscribe = instance.onCommandExecuted(() => {
      onDirtyRef.current()
    })

    return () => {
      unsubscribe()
      instance.dispose()
      handleRef.current = null
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  return <div className="spreadsheet-host" ref={containerRef} />
}
