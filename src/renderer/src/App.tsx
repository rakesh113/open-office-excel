import { useCallback, useEffect, useRef, useState } from 'react'
import { createBlankWorkbookData, type IWorkbookData } from '../../shared/workbook'
import { AppToolbar } from './components/AppToolbar'
import { Spreadsheet, type SpreadsheetHandle } from './components/Spreadsheet'
import { StatusBar } from './components/StatusBar'

function basename(filePath: string): string {
  return filePath.split(/[/\\]/).pop() || filePath
}

function titleFor(filePath: string | null, dirty: boolean): string {
  const name = filePath ? basename(filePath) : 'Untitled.xlsx'
  return `${dirty ? '• ' : ''}${name} — Open Office Excel`
}

export default function App(): JSX.Element {
  const [filePath, setFilePath] = useState<string | null>(null)
  const [dirty, setDirty] = useState(false)
  const [busy, setBusy] = useState(false)
  const [message, setMessage] = useState<string | null>(null)
  const [workbookKey, setWorkbookKey] = useState(0)
  const [initialData, setInitialData] = useState<IWorkbookData>(() => createBlankWorkbookData())

  const sheetRef = useRef<SpreadsheetHandle | null>(null)
  const filePathRef = useRef<string | null>(null)
  const dirtyRef = useRef(false)

  filePathRef.current = filePath
  dirtyRef.current = dirty

  const syncChrome = useCallback(async (path: string | null, isDirty: boolean) => {
    setFilePath(path)
    setDirty(isDirty)
    await window.api.setDocumentEdited(isDirty)
    await window.api.setTitle(titleFor(path, isDirty))
  }, [])

  const markDirty = useCallback(() => {
    if (dirtyRef.current) return
    void syncChrome(filePathRef.current, true)
  }, [syncChrome])

  const ensureDiscard = useCallback(async (): Promise<boolean> => {
    if (!dirtyRef.current) return true
    return window.api.confirmDiscard()
  }, [])

  const handleNew = useCallback(async () => {
    if (!(await ensureDiscard())) return
    setInitialData(createBlankWorkbookData())
    setWorkbookKey((k) => k + 1)
    setMessage('New workbook')
    await syncChrome(null, false)
  }, [ensureDiscard, syncChrome])

  const handleOpen = useCallback(async () => {
    if (!(await ensureDiscard())) return
    setBusy(true)
    setMessage('Opening…')
    try {
      const result = await window.api.openFile()
      if (result.canceled) {
        setMessage('Ready')
        return
      }
      setInitialData(result.workbook)
      setWorkbookKey((k) => k + 1)
      await syncChrome(result.path, false)
      setMessage(`Opened ${basename(result.path)}`)
    } catch (error) {
      const msg = error instanceof Error ? error.message : String(error)
      await window.api.showError('Open failed', msg)
      setMessage('Open failed')
    } finally {
      setBusy(false)
    }
  }, [ensureDiscard, syncChrome])

  const writeSnapshot = useCallback(
    async (targetPath: string): Promise<boolean> => {
      const snapshot = sheetRef.current?.getSnapshot()
      if (!snapshot) {
        await window.api.showError('Save failed', 'Could not read workbook data from the editor.')
        return false
      }
      try {
        const result = await window.api.writeWorkbook(targetPath, snapshot)
        if (!result.ok) {
          await window.api.showError('Save failed', result.error)
          return false
        }
        await syncChrome(targetPath, false)
        setMessage(`Saved ${basename(targetPath)}`)
        return true
      } catch (error) {
        const msg = error instanceof Error ? error.message : String(error)
        await window.api.showError('Save failed', msg)
        setMessage('Save failed')
        return false
      }
    },
    [syncChrome]
  )

  const handleSave = useCallback(async () => {
    setBusy(true)
    try {
      if (filePathRef.current) {
        await writeSnapshot(filePathRef.current)
      } else {
        const pick = await window.api.saveFileAs('Untitled.xlsx')
        if (!pick.canceled) {
          await writeSnapshot(pick.path)
        }
      }
    } finally {
      setBusy(false)
    }
  }, [writeSnapshot])

  const handleSaveAs = useCallback(async () => {
    setBusy(true)
    try {
      const defaultName = filePathRef.current ? basename(filePathRef.current) : 'Untitled.xlsx'
      const pick = await window.api.saveFileAs(defaultName)
      if (!pick.canceled) {
        await writeSnapshot(pick.path)
      }
    } finally {
      setBusy(false)
    }
  }, [writeSnapshot])

  const handleDatePicker = useCallback(() => {
    const result = sheetRef.current?.applyDatePicker()
    if (!result) {
      setMessage('Spreadsheet not ready')
      return
    }
    if (!result.ok) {
      setMessage(result.error)
      void window.api.showError('Date picker', result.error)
      return
    }
    markDirty()
    setMessage('Calendar enabled — click the cell dropdown to pick a date')
  }, [markDirty])

  const handleClearDatePicker = useCallback(() => {
    const result = sheetRef.current?.clearDatePicker()
    if (!result) {
      setMessage('Spreadsheet not ready')
      return
    }
    if (!result.ok) {
      setMessage(result.error)
      void window.api.showError('Clear date', result.error)
      return
    }
    markDirty()
    setMessage('Calendar removed from selection')
  }, [markDirty])

  const runSheetAction = useCallback(
    (
      action: (() => { ok: true; message?: string } | { ok: false; error: string }) | undefined,
      successMessage: string,
      errorTitle: string
    ) => {
      const result = action?.()
      if (!result) {
        setMessage('Spreadsheet not ready')
        return
      }
      if (!result.ok) {
        setMessage(result.error)
        void window.api.showError(errorTitle, result.error)
        return
      }
      markDirty()
      setMessage(result.message ?? successMessage)
    },
    [markDirty]
  )

  useEffect(() => {
    void window.api.setTitle(titleFor(null, false))
  }, [])

  useEffect(() => {
    const unsubs = [
      window.api.onMenuNew(() => {
        void handleNew()
      }),
      window.api.onMenuOpen(() => {
        void handleOpen()
      }),
      window.api.onMenuSave(() => {
        void handleSave()
      }),
      window.api.onMenuSaveAs(() => {
        void handleSaveAs()
      })
    ]
    return () => {
      unsubs.forEach((u) => u())
    }
  }, [handleNew, handleOpen, handleSave, handleSaveAs])

  return (
    <div className="app-root">
      <AppToolbar
        onNew={() => void handleNew()}
        onOpen={() => void handleOpen()}
        onSave={() => void handleSave()}
        onSaveAs={() => void handleSaveAs()}
        onDatePicker={handleDatePicker}
        onClearDatePicker={handleClearDatePicker}
        onToggleFilter={() =>
          runSheetAction(
            () => sheetRef.current?.toggleFilter() ?? { ok: false, error: 'Spreadsheet not ready' },
            'Filter toggled — use column dropdowns to filter',
            'Filter'
          )
        }
        onClearFilter={() =>
          runSheetAction(
            () => sheetRef.current?.clearFilter() ?? { ok: false, error: 'Spreadsheet not ready' },
            'Filter removed',
            'Clear filter'
          )
        }
        onFreezeHeader={() =>
          runSheetAction(
            () => sheetRef.current?.freezeHeader() ?? { ok: false, error: 'Spreadsheet not ready' },
            'First row frozen as header',
            'Freeze header'
          )
        }
        onFreezeAbove={() =>
          runSheetAction(
            () =>
              sheetRef.current?.freezeAboveSelection() ?? {
                ok: false,
                error: 'Spreadsheet not ready'
              },
            'Rows frozen to selection',
            'Freeze to this row'
          )
        }
        onUnfreeze={() =>
          runSheetAction(
            () => sheetRef.current?.unfreeze() ?? { ok: false, error: 'Spreadsheet not ready' },
            'Panes unfrozen',
            'Unfreeze'
          )
        }
        dirty={dirty}
        busy={busy}
      />
      <div className="spreadsheet-wrap">
        <Spreadsheet
          key={workbookKey}
          initialData={initialData}
          onReady={(handle) => {
            sheetRef.current = handle
          }}
          onDirty={markDirty}
        />
      </div>
      <StatusBar filePath={filePath} dirty={dirty} message={message} />
    </div>
  )
}
