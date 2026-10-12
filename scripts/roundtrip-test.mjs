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
const {
  xlsxBufferToWorkbookData,
  workbookDataToXlsxBuffer,
  csvBufferToWorkbookData,
  workbookDataToCsvBuffer,
  createBlankWorkbookData
} = require(outfile)

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

const csvPath = join(root, 'fixtures/sample.csv')
const csvData = await csvBufferToWorkbookData(readFileSync(csvPath), 'sample.csv')
const csvSheet = csvData.sheets[csvData.sheetOrder[0]]
const csvA2 = csvSheet.cellData[1]?.[0]?.v
const csvB2 = csvSheet.cellData[1]?.[1]?.v
const csvC2 = csvSheet.cellData[1]?.[2]?.v
const csvD2 = csvSheet.cellData[1]?.[3]?.v
const csvE2 = csvSheet.cellData[1]?.[4]?.v
console.log('csv', { csvA2, csvB2, csvC2, csvD2, csvE2 })
if (csvA2 !== 'Hello' || csvB2 !== 1234.5 || csvC2 !== 'Quoted, value' || csvE2 !== true) {
  throw new Error(`CSV open mismatch: ${JSON.stringify({ csvA2, csvB2, csvC2, csvD2, csvE2 })}`)
}
if (typeof csvD2 !== 'number') {
  throw new Error(`CSV date was not parsed as a number: ${csvD2}`)
}

const csvOut = Buffer.from(await workbookDataToCsvBuffer(csvData)).toString('utf8')
if (
  !csvOut.startsWith('\uFEFF') ||
  !csvOut.includes('Hello') ||
  !csvOut.includes('Quoted, value') ||
  !csvOut.includes('1234.5') ||
  !csvOut.includes('2026-07-25')
) {
  throw new Error(`CSV save mismatch:\n${csvOut}`)
}
writeFileSync(join(tmpdir(), 'sample-ooe-roundtrip.csv'), csvOut)
console.log('csv roundtrip ok')
