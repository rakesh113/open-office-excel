import { existsSync, readFileSync, writeFileSync } from 'fs'
import { createRequire } from 'module'
import { join, dirname } from 'path'
import { fileURLToPath } from 'url'
import { tmpdir } from 'os'
import esbuild from 'esbuild'
import ExcelJS from 'exceljs'

const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const samplePath = join(root, 'fixtures/sample.xlsx')

if (!existsSync(samplePath)) {
  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet('Data')
  ws.getCell('A1').value = 'Hello'
  ws.getCell('A1').font = { bold: true, size: 14 }
  ws.getCell('B1').value = 1234.5
  ws.getCell('B1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } }
  ws.getCell('C1').value = new Date(2026, 6, 25)
  ws.getCell('C1').numFmt = 'yyyy-mm-dd'
  ws.getColumn(1).width = 20
  ws.getRow(1).height = 30
  writeFileSync(samplePath, Buffer.from(await wb.xlsx.writeBuffer()))
}

const outfile = join(tmpdir(), 'exceljs-adapter-test.cjs')
await esbuild.build({
  entryPoints: [join(root, 'src/main/io/exceljs-adapter.ts')],
  outfile,
  bundle: true,
  platform: 'node',
  format: 'cjs'
})

const require = createRequire(import.meta.url)
const { xlsxBufferToWorkbookData, workbookDataToXlsxBuffer, createBlankWorkbookData } = require(outfile)

const buf = readFileSync(samplePath)
const data = await xlsxBufferToWorkbookData(buf, 'sample.xlsx')
console.log('sheets', data.sheetOrder)
const sheet = data.sheets[data.sheetOrder[0]]
console.log('A1', sheet.cellData[0]?.[0])
console.log('B1', sheet.cellData[0]?.[1])
console.log('C1', sheet.cellData[0]?.[2])
console.log('col0 width', sheet.columnData?.[0])
console.log('row0 height', sheet.rowData?.[0])

const out = await workbookDataToXlsxBuffer(data)
writeFileSync(join(tmpdir(), 'sample-ooe-roundtrip.xlsx'), out)
console.log('roundtrip bytes', out.length)

const blankOut = await workbookDataToXlsxBuffer(createBlankWorkbookData())
writeFileSync(join(tmpdir(), 'blank.xlsx'), blankOut)
console.log('blank ok', blankOut.length)
