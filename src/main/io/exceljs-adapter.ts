import { Readable } from 'node:stream'
import ExcelJS from 'exceljs'
import {
  CellValueType,
  createBlankWorkbookData,
  dateToExcelSerial,
  excelSerialToDate,
  type CellValue,
  type ICellData,
  type IStyleData,
  type IWorkbookData,
  type IWorksheetData
} from '../../shared/workbook'

const DEFAULT_ROWS = 100
const DEFAULT_COLS = 26

function sheetId(index: number): string {
  return `sheet-${index + 1}`
}

function styleId(index: number): string {
  return `s${index}`
}

function argbToRgb(argb?: string): string | undefined {
  if (!argb) return undefined
  const hex = argb.replace(/^#/, '')
  if (hex.length === 8) return `#${hex.slice(2)}`
  if (hex.length === 6) return `#${hex}`
  return undefined
}

function rgbToArgb(rgb?: string): string | undefined {
  if (!rgb) return undefined
  const hex = rgb.replace(/^#/, '')
  if (hex.length === 6) return `FF${hex.toUpperCase()}`
  if (hex.length === 8) return hex.toUpperCase()
  return undefined
}

function looksLikeDateNumFmt(numFmt?: string): boolean {
  if (!numFmt || typeof numFmt !== 'string') return false
  const fmt = numFmt.toLowerCase()
  return /[ymdhs]/i.test(fmt) && !/^0/.test(fmt.trim())
}

function excelHorizontalToUniver(align?: ExcelJS.Alignment['horizontal']): number | undefined {
  switch (align) {
    case 'left':
      return 1
    case 'center':
    case 'centerContinuous':
      return 2
    case 'right':
      return 3
    case 'justify':
      return 4
    default:
      return undefined
  }
}

function univerHorizontalToExcel(ht?: number): ExcelJS.Alignment['horizontal'] | undefined {
  switch (ht) {
    case 1:
      return 'left'
    case 2:
      return 'center'
    case 3:
      return 'right'
    case 4:
      return 'justify'
    default:
      return undefined
  }
}

function excelVerticalToUniver(align?: ExcelJS.Alignment['vertical']): number | undefined {
  switch (align) {
    case 'top':
      return 1
    case 'middle':
      return 2
    case 'bottom':
      return 3
    default:
      return undefined
  }
}

function univerVerticalToExcel(vt?: number): ExcelJS.Alignment['vertical'] | undefined {
  switch (vt) {
    case 1:
      return 'top'
    case 2:
      return 'middle'
    case 3:
      return 'bottom'
    default:
      return undefined
  }
}

function styleFromExcelCell(cell: ExcelJS.Cell): IStyleData | null {
  const style: IStyleData = {}
  let has = false

  const font = cell.font
  if (font) {
    if (font.name) {
      style.ff = font.name
      has = true
    }
    if (font.size) {
      style.fs = font.size
      has = true
    }
    if (font.bold) {
      style.bl = 1
      has = true
    }
    if (font.italic) {
      style.it = 1
      has = true
    }
    if (font.underline) {
      style.ul = { s: 1 }
      has = true
    }
    if (font.strike) {
      style.st = { s: 1 }
      has = true
    }
    const color = argbToRgb((font.color as ExcelJS.Color & { argb?: string })?.argb)
    if (color) {
      style.cl = { rgb: color }
      has = true
    }
  }

  const fill = cell.fill
  if (fill && fill.type === 'pattern' && fill.pattern !== 'none') {
    const fg = fill.fgColor as ExcelJS.Color & { argb?: string }
    const bg = argbToRgb(fg?.argb)
    if (bg) {
      style.bg = { rgb: bg }
      has = true
    }
  }

  const alignment = cell.alignment
  if (alignment) {
    const ht = excelHorizontalToUniver(alignment.horizontal)
    const vt = excelVerticalToUniver(alignment.vertical)
    if (ht !== undefined) {
      style.ht = ht
      has = true
    }
    if (vt !== undefined) {
      style.vt = vt
      has = true
    }
  }

  const numFmt = typeof cell.numFmt === 'string' ? cell.numFmt : undefined
  if (numFmt && numFmt !== 'General') {
    style.n = { pattern: numFmt }
    has = true
  }

  return has ? style : null
}

function applyStyleToExcelCell(cell: ExcelJS.Cell, style: IStyleData): void {
  if (style.ff || style.fs || style.bl || style.it || style.ul || style.st || style.cl) {
    cell.font = {
      name: style.ff,
      size: style.fs,
      bold: style.bl === 1,
      italic: style.it === 1,
      underline: style.ul?.s === 1,
      strike: style.st?.s === 1,
      color: style.cl?.rgb ? { argb: rgbToArgb(style.cl.rgb) } : undefined
    }
  }

  if (style.bg?.rgb) {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: rgbToArgb(style.bg.rgb) }
    }
  }

  const horizontal = univerHorizontalToExcel(style.ht)
  const vertical = univerVerticalToExcel(style.vt)
  if (horizontal || vertical) {
    cell.alignment = {
      horizontal,
      vertical
    }
  }

  if (style.n?.pattern) {
    cell.numFmt = style.n.pattern
  }
}

function cellFromExcel(cell: ExcelJS.Cell): { data: ICellData; style: IStyleData | null } | null {
  const style = styleFromExcelCell(cell)
  const data: ICellData = {}

  if (cell.formula) {
    data.f = cell.formula.startsWith('=') ? cell.formula : `=${cell.formula}`
  }

  const value = cell.value
  if (value === null || value === undefined) {
    if (!data.f && !style) return null
    if (style) data.s = style
    return { data, style }
  }

  if (value instanceof Date) {
    data.v = dateToExcelSerial(value)
    data.t = CellValueType.NUMBER
    if (!style?.n?.pattern) {
      const s = style ?? {}
      s.n = { pattern: 'yyyy-mm-dd' }
      data.s = s
      return { data, style: s }
    }
  } else if (typeof value === 'number') {
    data.v = value
    data.t = CellValueType.NUMBER
  } else if (typeof value === 'boolean') {
    data.v = value
    data.t = CellValueType.BOOLEAN
  } else if (typeof value === 'string') {
    data.v = value
    data.t = CellValueType.STRING
  } else if (typeof value === 'object' && value !== null) {
    const rich = value as {
      richText?: Array<{ text?: string }>
      text?: string
      result?: CellValue | Date
      formula?: string
    }
    if (rich.formula && !data.f) {
      data.f = rich.formula.startsWith('=') ? rich.formula : `=${rich.formula}`
    }
    if (rich.result !== undefined && rich.result !== null) {
      if (rich.result instanceof Date) {
        data.v = dateToExcelSerial(rich.result)
        data.t = CellValueType.NUMBER
      } else if (typeof rich.result === 'number') {
        data.v = rich.result
        data.t = CellValueType.NUMBER
      } else if (typeof rich.result === 'boolean') {
        data.v = rich.result
        data.t = CellValueType.BOOLEAN
      } else {
        data.v = String(rich.result)
        data.t = CellValueType.STRING
      }
    } else if (rich.richText) {
      data.v = rich.richText.map((t) => t.text ?? '').join('')
      data.t = CellValueType.STRING
    } else if (rich.text !== undefined) {
      data.v = rich.text
      data.t = CellValueType.STRING
    } else {
      data.v = String(value)
      data.t = CellValueType.STRING
    }
  }

  if (style) {
    data.s = style
  }

  return { data, style }
}

function decodeRange(
  range: string
): { startRow: number; endRow: number; startColumn: number; endColumn: number } | null {
  const match = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i.exec(range)
  if (!match) return null
  return {
    startColumn: colLettersToIndex(match[1]),
    startRow: Number(match[2]) - 1,
    endColumn: colLettersToIndex(match[3]),
    endRow: Number(match[4]) - 1
  }
}

function colLettersToIndex(letters: string): number {
  let n = 0
  const upper = letters.toUpperCase()
  for (let i = 0; i < upper.length; i++) {
    n = n * 26 + (upper.charCodeAt(i) - 64)
  }
  return n - 1
}

function indexToColLetters(index: number): string {
  let n = index + 1
  let s = ''
  while (n > 0) {
    const rem = (n - 1) % 26
    s = String.fromCharCode(65 + rem) + s
    n = Math.floor((n - 1) / 26)
  }
  return s
}

type FrozenView = {
  state?: string
  xSplit?: number
  ySplit?: number
  topLeftCell?: string
  activeCell?: string
  showGridLines?: boolean
}

function freezeFromExcelViews(
  views: ExcelJS.Worksheet['views'] | undefined
): IWorksheetData['freeze'] | undefined {
  const view = (views as FrozenView[] | undefined)?.find(
    (v) => v.state === 'frozen' || (v.ySplit ?? 0) > 0 || (v.xSplit ?? 0) > 0
  )
  if (!view) return undefined

  const ySplit = Number(view.ySplit ?? 0)
  const xSplit = Number(view.xSplit ?? 0)
  if (ySplit <= 0 && xSplit <= 0) return undefined

  return {
    ySplit,
    xSplit,
    startRow: ySplit,
    startColumn: xSplit
  }
}

function applyFreezeToExcelWorksheet(worksheet: ExcelJS.Worksheet, freeze?: IWorksheetData['freeze']): void {
  if (!freeze || (freeze.ySplit <= 0 && freeze.xSplit <= 0)) {
    return
  }

  const topLeft = `${indexToColLetters(freeze.startColumn ?? freeze.xSplit)}${(freeze.startRow ?? freeze.ySplit) + 1}`
  const view: FrozenView = {
    state: 'frozen',
    xSplit: freeze.xSplit,
    ySplit: freeze.ySplit,
    topLeftCell: topLeft,
    activeCell: topLeft,
    showGridLines: true
  }
  worksheet.views = [view as ExcelJS.WorksheetView]
}

function resolveStyle(styles: Record<string, IStyleData>, cell: ICellData): IStyleData | undefined {
  if (!cell.s) return undefined
  if (typeof cell.s === 'string') return styles[cell.s]
  return cell.s
}

function isDateStyle(style?: IStyleData): boolean {
  return looksLikeDateNumFmt(style?.n?.pattern)
}

function workbookNameFromFile(fileName?: string): string {
  if (!fileName) return 'Untitled'
  return fileName.replace(/\.(xlsx|csv)$/i, '') || 'Untitled'
}

function stripUtf8Bom(buffer: Buffer): Buffer {
  if (buffer.length >= 3 && buffer[0] === 0xef && buffer[1] === 0xbb && buffer[2] === 0xbf) {
    return buffer.subarray(3)
  }
  return buffer
}

const CSV_DATE_FORMATS = ['YYYY-MM-DD', 'M/D/YYYY', 'MM/DD/YYYY', 'YYYY/MM/DD']

function excelWorkbookToWorkbookData(workbook: ExcelJS.Workbook, fileName?: string): IWorkbookData {
  const styles: Record<string, IStyleData> = {}
  const styleMap = new Map<string, string>()
  let styleCounter = 0

  const internStyle = (style: IStyleData | null): string | undefined => {
    if (!style) return undefined
    const key = JSON.stringify(style)
    const existing = styleMap.get(key)
    if (existing) return existing
    const id = styleId(styleCounter++)
    styleMap.set(key, id)
    styles[id] = style
    return id
  }

  const sheets: Record<string, IWorksheetData> = {}
  const sheetOrder: string[] = []

  workbook.eachSheet((worksheet, sheetIndex) => {
    const id = sheetId(sheetIndex - 1)
    sheetOrder.push(id)

    const cellData: Record<number, Record<number, ICellData>> = {}
    let maxRow = DEFAULT_ROWS
    let maxCol = DEFAULT_COLS

    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      maxRow = Math.max(maxRow, rowNumber + 20)
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        maxCol = Math.max(maxCol, colNumber + 5)
        const parsed = cellFromExcel(cell)
        if (!parsed) return
        const sid = internStyle(parsed.style)
        const cellOut: ICellData = { ...parsed.data }
        if (sid) {
          cellOut.s = sid
        } else {
          delete cellOut.s
        }
        const r = rowNumber - 1
        const c = colNumber - 1
        if (!cellData[r]) cellData[r] = {}
        cellData[r][c] = cellOut
      })
    })

    const columnData: Record<number, { w?: number; hd?: number }> = {}
    worksheet.columns?.forEach((col, index) => {
      if (col.width) {
        columnData[index] = { w: Math.round(col.width * 8) }
      }
    })

    const rowData: Record<number, { h?: number; hd?: number }> = {}
    worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      if (row.height) {
        rowData[rowNumber - 1] = { h: row.height }
      }
    })

    const merges = (worksheet.model as { merges?: string[] }).merges ?? []
    const mergeData = merges.map((range) => decodeRange(range)).filter(Boolean) as NonNullable<
      IWorksheetData['mergeData']
    >

    const freeze = freezeFromExcelViews(worksheet.views)

    sheets[id] = {
      id,
      name: worksheet.name,
      rowCount: Math.max(maxRow, worksheet.rowCount || DEFAULT_ROWS),
      columnCount: Math.max(maxCol, worksheet.columnCount || DEFAULT_COLS),
      cellData,
      rowData,
      columnData,
      mergeData,
      freeze,
      showGridlines: 1,
      defaultColumnWidth: 88,
      defaultRowHeight: 24
    }
  })

  if (sheetOrder.length === 0) {
    return createBlankWorkbookData(workbookNameFromFile(fileName))
  }

  return {
    id: 'workbook-1',
    name: workbookNameFromFile(fileName) || workbook.title || 'Workbook',
    appVersion: '1.0.0',
    locale: 'enUS',
    styles,
    sheetOrder,
    sheets
  }
}

export async function xlsxBufferToWorkbookData(
  buffer: Buffer | Uint8Array,
  fileName?: string
): Promise<IWorkbookData> {
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(Buffer.from(buffer) as unknown as ExcelJS.Buffer)
  return excelWorkbookToWorkbookData(workbook, fileName)
}

export async function csvBufferToWorkbookData(
  buffer: Buffer | Uint8Array,
  fileName?: string
): Promise<IWorkbookData> {
  const workbook = new ExcelJS.Workbook()
  const bytes = stripUtf8Bom(Buffer.from(buffer))
  await workbook.csv.read(Readable.from(bytes), {
    sheetName: 'Sheet1',
    dateFormats: CSV_DATE_FORMATS
  })
  return excelWorkbookToWorkbookData(workbook, fileName)
}

function workbookDataToExcelWorkbook(data: IWorkbookData): ExcelJS.Workbook {
  const workbook = new ExcelJS.Workbook()
  workbook.creator = 'Open Office Excel'
  workbook.created = new Date()

  for (const sheetIdKey of data.sheetOrder) {
    const sheetData = data.sheets[sheetIdKey]
    if (!sheetData) continue

    const worksheet = workbook.addWorksheet(sheetData.name || 'Sheet')
    applyFreezeToExcelWorksheet(worksheet, sheetData.freeze)

    if (sheetData.columnData) {
      Object.entries(sheetData.columnData).forEach(([key, value]) => {
        const index = Number(key)
        if (value.w) {
          worksheet.getColumn(index + 1).width = Math.max(4, value.w / 8)
        }
      })
    }

    if (sheetData.rowData) {
      Object.entries(sheetData.rowData).forEach(([key, value]) => {
        const index = Number(key)
        if (value.h) {
          worksheet.getRow(index + 1).height = value.h
        }
      })
    }

    Object.entries(sheetData.cellData || {}).forEach(([rowKey, row]) => {
      const r = Number(rowKey)
      Object.entries(row).forEach(([colKey, cell]) => {
        const c = Number(colKey)
        const excelCell = worksheet.getCell(r + 1, c + 1)
        const style = resolveStyle(data.styles || {}, cell)

        if (cell.f) {
          const formula = cell.f.startsWith('=') ? cell.f.slice(1) : cell.f
          excelCell.value = { formula, result: cell.v as string | number | boolean | undefined }
        } else if (cell.v !== undefined && cell.v !== null) {
          if (cell.t === CellValueType.NUMBER || typeof cell.v === 'number') {
            if (isDateStyle(style) && typeof cell.v === 'number') {
              excelCell.value = excelSerialToDate(cell.v)
            } else {
              excelCell.value = cell.v as number
            }
          } else if (cell.t === CellValueType.BOOLEAN || typeof cell.v === 'boolean') {
            excelCell.value = Boolean(cell.v)
          } else {
            excelCell.value = String(cell.v)
          }
        }

        if (style) {
          applyStyleToExcelCell(excelCell, style)
        }
      })
    })

    if (sheetData.mergeData?.length) {
      for (const merge of sheetData.mergeData) {
        const start = `${indexToColLetters(merge.startColumn)}${merge.startRow + 1}`
        const end = `${indexToColLetters(merge.endColumn)}${merge.endRow + 1}`
        worksheet.mergeCells(`${start}:${end}`)
      }
    }
  }

  return workbook
}

export async function workbookDataToXlsxBuffer(data: IWorkbookData): Promise<Uint8Array> {
  const workbook = workbookDataToExcelWorkbook(data)
  const buffer = await workbook.xlsx.writeBuffer()
  return new Uint8Array(buffer)
}

export async function workbookDataToCsvBuffer(data: IWorkbookData): Promise<Uint8Array> {
  const firstId = data.sheetOrder[0]
  const firstSheet = firstId ? data.sheets[firstId] : undefined
  const singleSheet: IWorkbookData = firstSheet
    ? { ...data, sheetOrder: [firstId], sheets: { [firstId]: firstSheet } }
    : data
  const workbook = workbookDataToExcelWorkbook(singleSheet)
  const sheetName = firstSheet?.name || 'Sheet1'
  const csv = await workbook.csv.writeBuffer({
    sheetName,
    dateFormat: 'YYYY-MM-DD',
    formatterOptions: { rowDelimiter: '\r\n', writeBOM: true }
  })
  return new Uint8Array(csv)
}

export { createBlankWorkbookData }
