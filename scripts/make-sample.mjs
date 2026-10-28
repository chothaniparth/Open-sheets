// Generates samples/Sales Demo.xlsx — a small workbook for trying out formulas, charts and the AI assistant.
// Usage: node scripts/make-sample.mjs "samples/Sales Demo.xlsx"
import ExcelJS from 'exceljs'
const out = process.argv[2]
const wb = new ExcelJS.Workbook()
const ws = wb.addWorksheet('Sales', { views: [{ state: 'frozen', ySplit: 1 }] })
ws.columns = [
  { header: 'Month', key: 'm', width: 12 },
  { header: 'Region', key: 'r', width: 12 },
  { header: 'Product', key: 'p', width: 16 },
  { header: 'Units', key: 'u', width: 10 },
  { header: 'Price', key: 'pr', width: 10 },
  { header: 'Revenue', key: 'rev', width: 14 },
  { header: 'Date', key: 'd', width: 14 }
]
const regions = ['North', 'South', 'East', 'West']
const products = ['Notebook', 'Pen set', 'Desk lamp']
const prices = [12.5, 8, 39]
let row = 2
for (let m = 0; m < 6; m++) {
  for (const region of regions) {
    for (let p = 0; p < products.length; p++) {
      const units = 20 + ((m * 7 + p * 13 + region.length * 5) % 60)
      ws.addRow({ m: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'][m], r: region, p: products[p], u: units, pr: prices[p], d: new Date(Date.UTC(2026, m, 15)) })
      ws.getCell(`F${row}`).value = { formula: `D${row}*E${row}` }
      row++
    }
  }
}
ws.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } }
ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2A78D6' } }
ws.getColumn('E').numFmt = '$#,##0.00'
ws.getColumn('F').numFmt = '$#,##0.00'
ws.getColumn('G').numFmt = 'yyyy-mm-dd'
const s = wb.addWorksheet('Summary')
s.getCell('A1').value = 'Region'
s.getCell('B1').value = 'Revenue'
s.getCell('C1').value = 'Units'
regions.forEach((r, i) => {
  s.getCell(`A${i + 2}`).value = r
  s.getCell(`B${i + 2}`).value = { formula: `SUMIFS(Sales!F:F,Sales!B:B,A${i + 2})` }
  s.getCell(`C${i + 2}`).value = { formula: `SUMIFS(Sales!D:D,Sales!B:B,A${i + 2})` }
})
s.getCell('A6').value = 'Total'
s.getCell('B6').value = { formula: 'SUM(B2:B5)' }
s.getCell('C6').value = { formula: 'SUM(C2:C5)' }
s.getRow(1).font = { bold: true }
s.getRow(6).font = { bold: true }
s.getColumn('B').numFmt = '$#,##0'
s.getColumn('A').width = 14
s.getColumn('B').width = 14
await wb.xlsx.writeFile(out)
console.log('wrote', out)
