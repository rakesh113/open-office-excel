/** Shared workbook snapshot shapes (Univer-compatible). */

export type CellValue = string | number | boolean | null

export enum CellValueType {
  STRING = 1,
  NUMBER = 2,
  BOOLEAN = 3,
  FORCE_STRING = 4
}

export interface IStyleData {
  ff?: string
  fs?: number
  it?: number
  bl?: number
  ul?: { s: number }
  st?: { s: number }
  cl?: { rgb?: string }
  bg?: { rgb?: string }
  ht?: number
  vt?: number
  n?: { pattern?: string }
}

export interface ICellData {
  v?: CellValue
  t?: CellValueType
  s?: string | IStyleData
  f?: string
}

export interface IWorksheetData {
  id: string
  name: string
  tabColor?: string
  hidden?: number
  rowCount: number
  columnCount: number
  zoomRatio?: number
  freeze?: { startRow: number; startColumn: number; ySplit: number; xSplit: number }
  scrollTop?: number
  scrollLeft?: number
  defaultColumnWidth?: number
  defaultRowHeight?: number
  mergeData?: Array<{ startRow: number; endRow: number; startColumn: number; endColumn: number }>
  cellData: Record<number, Record<number, ICellData>>
  rowData?: Record<number, { h?: number; hd?: number }>
  columnData?: Record<number, { w?: number; hd?: number }>
  showGridlines?: number
  rowHeader?: { width: number }
  columnHeader?: { height: number }
}

export interface IWorkbookData {
  id: string
  name: string
  appVersion: string
  locale: string
  styles: Record<string, IStyleData>
  sheetOrder: string[]
  sheets: Record<string, IWorksheetData>
}

const DEFAULT_ROWS = 100
const DEFAULT_COLS = 26

export function createBlankWorkbookData(name = 'Untitled'): IWorkbookData {
  const id = 'sheet-1'
  const sheet: IWorksheetData = {
    id,
    name: 'Sheet1',
    rowCount: DEFAULT_ROWS,
    columnCount: DEFAULT_COLS,
    cellData: {},
    rowData: {},
    columnData: {},
    showGridlines: 1,
    defaultColumnWidth: 88,
    defaultRowHeight: 24
  }

  return {
    id: 'workbook-1',
    name,
    appVersion: '1.0.0',
    locale: 'enUS',
    styles: {},
    sheetOrder: [id],
    sheets: { [id]: sheet }
  }
}

const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30)

export function excelSerialToDate(serial: number): Date {
  const wholeDays = Math.floor(serial)
  const ms = Math.round((serial - wholeDays) * 86400000)
  const adjusted = wholeDays > 59 ? wholeDays - 1 : wholeDays
  const utc = EXCEL_EPOCH_UTC + adjusted * 86400000 + ms
  const d = new Date(utc)
  return new Date(
    d.getUTCFullYear(),
    d.getUTCMonth(),
    d.getUTCDate(),
    d.getUTCHours(),
    d.getUTCMinutes(),
    d.getUTCSeconds()
  )
}

export function dateToExcelSerial(date: Date): number {
  const utc = Date.UTC(
    date.getFullYear(),
    date.getMonth(),
    date.getDate(),
    date.getHours(),
    date.getMinutes(),
    date.getSeconds()
  )
  let days = Math.floor((utc - EXCEL_EPOCH_UTC) / 86400000)
  if (days >= 60) days += 1
  const dayStart = Date.UTC(date.getFullYear(), date.getMonth(), date.getDate())
  const fraction = (utc - dayStart) / 86400000
  return days + fraction
}

/** Heuristic parse for typed cell entry. */
export function parseUserInput(raw: string): { v: CellValue; t: CellValueType; numFmt?: string } {
  const trimmed = raw.trim()
  if (trimmed === '') {
    return { v: '', t: CellValueType.STRING }
  }

  if (/^(true|false)$/i.test(trimmed)) {
    return { v: /^true$/i.test(trimmed), t: CellValueType.BOOLEAN }
  }

  const dateMatch = trimmed.match(/^(\d{1,4})[/-](\d{1,2})[/-](\d{1,4})$/)
  if (dateMatch) {
    const a = Number(dateMatch[1])
    const b = Number(dateMatch[2])
    const c = Number(dateMatch[3])
    let year: number
    let month: number
    let day: number
    if (String(dateMatch[1]).length === 4) {
      year = a
      month = b
      day = c
    } else if (String(dateMatch[3]).length === 4) {
      month = a
      day = b
      year = c
    } else {
      return { v: trimmed, t: CellValueType.STRING }
    }
    const date = new Date(year, month - 1, day)
    if (date.getFullYear() === year && date.getMonth() === month - 1 && date.getDate() === day) {
      return { v: dateToExcelSerial(date), t: CellValueType.NUMBER, numFmt: 'yyyy-mm-dd' }
    }
  }

  if (/^-?\d+(\.\d+)?$/.test(trimmed)) {
    return { v: Number(trimmed), t: CellValueType.NUMBER }
  }

  return { v: raw, t: CellValueType.STRING }
}
