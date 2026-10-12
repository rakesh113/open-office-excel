import type { IWorkbookData } from './workbook'

export const IPC = {
  OPEN_FILE: 'dialog:openFile',
  SAVE_FILE_AS: 'dialog:saveFileAs',
  WRITE_WORKBOOK: 'fs:writeWorkbook',
  SET_DOCUMENT_EDITED: 'window:setDocumentEdited',
  SET_TITLE: 'window:setTitle',
  CONFIRM_DISCARD: 'dialog:confirmDiscard',
  MENU_NEW: 'menu:new',
  MENU_OPEN: 'menu:open',
  MENU_SAVE: 'menu:save',
  MENU_SAVE_AS: 'menu:saveAs',
  SHOW_ERROR: 'dialog:showError'
} as const

export type OpenFileResult =
  | { canceled: true }
  | { canceled: false; path: string; workbook: IWorkbookData }

export type SaveFileResult =
  | { canceled: true }
  | { canceled: false; path: string }

export interface ElectronAPI {
  openFile: () => Promise<OpenFileResult>
  saveFileAs: (defaultName?: string) => Promise<SaveFileResult>
  writeWorkbook: (
    filePath: string,
    workbook: IWorkbookData
  ) => Promise<{ ok: true } | { ok: false; canceled: true } | { ok: false; error: string }>
  setDocumentEdited: (edited: boolean) => Promise<void>
  setTitle: (title: string) => Promise<void>
  confirmDiscard: () => Promise<boolean>
  showError: (title: string, message: string) => Promise<void>
  onMenuNew: (callback: () => void) => () => void
  onMenuOpen: (callback: () => void) => () => void
  onMenuSave: (callback: () => void) => () => void
  onMenuSaveAs: (callback: () => void) => () => void
}
