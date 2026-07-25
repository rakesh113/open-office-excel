import { contextBridge, ipcRenderer } from 'electron'
import { IPC, type ElectronAPI } from '../shared/ipc'
import type { IWorkbookData } from '../shared/workbook'

function subscribe(channel: string, callback: () => void): () => void {
  const listener = (): void => {
    callback()
  }
  ipcRenderer.on(channel, listener)
  return () => {
    ipcRenderer.removeListener(channel, listener)
  }
}

const api: ElectronAPI = {
  openFile: () => ipcRenderer.invoke(IPC.OPEN_FILE),
  saveFileAs: (defaultName?: string) => ipcRenderer.invoke(IPC.SAVE_FILE_AS, defaultName),
  writeWorkbook: (filePath: string, workbook: IWorkbookData) =>
    ipcRenderer.invoke(IPC.WRITE_WORKBOOK, filePath, workbook),
  setDocumentEdited: (edited: boolean) => ipcRenderer.invoke(IPC.SET_DOCUMENT_EDITED, edited),
  setTitle: (title: string) => ipcRenderer.invoke(IPC.SET_TITLE, title),
  confirmDiscard: () => ipcRenderer.invoke(IPC.CONFIRM_DISCARD),
  showError: (title: string, message: string) => ipcRenderer.invoke(IPC.SHOW_ERROR, title, message),
  onMenuNew: (callback) => subscribe(IPC.MENU_NEW, callback),
  onMenuOpen: (callback) => subscribe(IPC.MENU_OPEN, callback),
  onMenuSave: (callback) => subscribe(IPC.MENU_SAVE, callback),
  onMenuSaveAs: (callback) => subscribe(IPC.MENU_SAVE_AS, callback)
}

contextBridge.exposeInMainWorld('api', api)
