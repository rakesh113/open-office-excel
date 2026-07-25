import { app, shell, BrowserWindow, ipcMain, dialog, Menu } from 'electron'
import { join } from 'path'
import { writeFile } from 'fs/promises'
import { IPC } from '../shared/ipc'
import type { IWorkbookData } from '../shared/workbook'
import { workbookDataToXlsxBuffer, xlsxBufferToWorkbookData } from './io/exceljs-adapter'
import { readFile } from 'fs/promises'

let mainWindow: BrowserWindow | null = null

function createWindow(): void {
  mainWindow = new BrowserWindow({
    width: 1280,
    height: 800,
    minWidth: 800,
    minHeight: 500,
    show: false,
    title: 'Untitled — Open Office Excel',
    webPreferences: {
      preload: join(__dirname, '../preload/index.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false
    }
  })

  mainWindow.on('ready-to-show', () => {
    mainWindow?.show()
  })

  mainWindow.webContents.setWindowOpenHandler((details) => {
    shell.openExternal(details.url)
    return { action: 'deny' }
  })

  mainWindow.on('close', (event) => {
    if (mainWindow && !mainWindow.isDestroyed() && mainWindow.isDocumentEdited()) {
      const choice = dialog.showMessageBoxSync(mainWindow, {
        type: 'warning',
        buttons: ['Cancel', 'Discard'],
        defaultId: 0,
        cancelId: 0,
        title: 'Unsaved Changes',
        message: 'You have unsaved changes. Discard them?'
      })
      if (choice === 0) {
        event.preventDefault()
      }
    }
  })

  if (process.env.ELECTRON_RENDERER_URL) {
    mainWindow.loadURL(process.env.ELECTRON_RENDERER_URL)
  } else {
    mainWindow.loadFile(join(__dirname, '../renderer/index.html'))
  }
}

function buildMenu(): void {
  const isMac = process.platform === 'darwin'

  const template: Electron.MenuItemConstructorOptions[] = [
    ...(isMac
      ? [
          {
            label: app.name,
            submenu: [
              { role: 'about' as const },
              { type: 'separator' as const },
              { role: 'services' as const },
              { type: 'separator' as const },
              { role: 'hide' as const },
              { role: 'hideOthers' as const },
              { role: 'unhide' as const },
              { type: 'separator' as const },
              { role: 'quit' as const }
            ]
          }
        ]
      : []),
    {
      label: 'File',
      submenu: [
        {
          label: 'New',
          accelerator: 'CmdOrCtrl+N',
          click: () => mainWindow?.webContents.send(IPC.MENU_NEW)
        },
        {
          label: 'Open…',
          accelerator: 'CmdOrCtrl+O',
          click: () => mainWindow?.webContents.send(IPC.MENU_OPEN)
        },
        {
          label: 'Save',
          accelerator: 'CmdOrCtrl+S',
          click: () => mainWindow?.webContents.send(IPC.MENU_SAVE)
        },
        {
          label: 'Save As…',
          accelerator: 'CmdOrCtrl+Shift+S',
          click: () => mainWindow?.webContents.send(IPC.MENU_SAVE_AS)
        },
        { type: 'separator' },
        isMac ? { role: 'close' } : { role: 'quit' }
      ]
    },
    {
      label: 'Edit',
      submenu: [
        { role: 'undo' },
        { role: 'redo' },
        { type: 'separator' },
        { role: 'cut' },
        { role: 'copy' },
        { role: 'paste' },
        { role: 'selectAll' }
      ]
    },
    {
      label: 'View',
      submenu: [
        { role: 'reload' },
        { role: 'forceReload' },
        { role: 'toggleDevTools' },
        { type: 'separator' },
        { role: 'resetZoom' },
        { role: 'zoomIn' },
        { role: 'zoomOut' },
        { type: 'separator' },
        { role: 'togglefullscreen' }
      ]
    },
    {
      label: 'Window',
      submenu: [
        { role: 'minimize' },
        { role: 'zoom' },
        ...(isMac
          ? [{ type: 'separator' as const }, { role: 'front' as const }]
          : [{ role: 'close' as const }])
      ]
    }
  ]

  Menu.setApplicationMenu(Menu.buildFromTemplate(template))
}

function registerIpc(): void {
  ipcMain.handle(IPC.OPEN_FILE, async () => {
    if (!mainWindow) return { canceled: true as const }

    const result = await dialog.showOpenDialog(mainWindow, {
      title: 'Open Spreadsheet',
      filters: [{ name: 'Excel Workbook', extensions: ['xlsx'] }],
      properties: ['openFile']
    })

    if (result.canceled || result.filePaths.length === 0) {
      return { canceled: true as const }
    }

    const filePath = result.filePaths[0]
    try {
      const buffer = await readFile(filePath)
      const workbook = await xlsxBufferToWorkbookData(buffer, filePath.split(/[/\\]/).pop())
      return {
        canceled: false as const,
        path: filePath,
        workbook
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error)
      await dialog.showMessageBox(mainWindow, {
        type: 'error',
        title: 'Open failed',
        message
      })
      return { canceled: true as const }
    }
  })

  ipcMain.handle(IPC.SAVE_FILE_AS, async (_event, defaultName?: string) => {
    if (!mainWindow) return { canceled: true as const }

    const result = await dialog.showSaveDialog(mainWindow, {
      title: 'Save Spreadsheet',
      defaultPath: defaultName ?? 'Untitled.xlsx',
      filters: [{ name: 'Excel Workbook', extensions: ['xlsx'] }]
    })

    if (result.canceled || !result.filePath) {
      return { canceled: true as const }
    }

    let filePath = result.filePath
    if (!filePath.toLowerCase().endsWith('.xlsx')) {
      filePath = `${filePath}.xlsx`
    }

    return { canceled: false as const, path: filePath }
  })

  ipcMain.handle(IPC.WRITE_WORKBOOK, async (_event, filePath: string, workbook: IWorkbookData) => {
    try {
      const data = await workbookDataToXlsxBuffer(workbook)
      await writeFile(filePath, Buffer.from(data))
      return { ok: true as const }
    } catch (error) {
      return {
        ok: false as const,
        error: error instanceof Error ? error.message : String(error)
      }
    }
  })

  ipcMain.handle(IPC.SET_DOCUMENT_EDITED, async (_event, edited: boolean) => {
    mainWindow?.setDocumentEdited(edited)
  })

  ipcMain.handle(IPC.SET_TITLE, async (_event, title: string) => {
    if (mainWindow) {
      mainWindow.setTitle(title)
    }
  })

  ipcMain.handle(IPC.CONFIRM_DISCARD, async () => {
    if (!mainWindow) return true

    const result = await dialog.showMessageBox(mainWindow, {
      type: 'warning',
      buttons: ['Cancel', 'Discard'],
      defaultId: 0,
      cancelId: 0,
      title: 'Unsaved Changes',
      message: 'You have unsaved changes. Discard them?'
    })

    return result.response === 1
  })

  ipcMain.handle(IPC.SHOW_ERROR, async (_event, title: string, message: string) => {
    if (!mainWindow) return
    await dialog.showMessageBox(mainWindow, {
      type: 'error',
      title,
      message
    })
  })
}

app.whenReady().then(() => {
  registerIpc()
  buildMenu()
  createWindow()

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow()
    }
  })
})

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit()
  }
})
