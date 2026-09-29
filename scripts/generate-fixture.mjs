// 테스트 픽스처를 만드는 1회성 스크립트. `node scripts/generate-fixture.mjs`로 실행해서
// src/xlsx/__fixtures__/sample.xlsx를 (다시) 만든다. 테스트 실행 중에는 안 돈다 —
// 커밋된 바이너리 픽스처를 그대로 쓴다.
import ExcelJS from 'exceljs'

const wb = new ExcelJS.Workbook()
const ws = wb.addWorksheet('Data')

ws.getCell('A1').value = 'Name'
ws.getCell('B1').value = 'Amount'
ws.getRow(1).font = { bold: true }
ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEFEFEF' } }

ws.getCell('A2').value = 'Alpha'
ws.getCell('B2').value = 100
ws.getCell('B2').numFmt = '#,##0'

ws.getCell('A3').value = 'Beta'
ws.getCell('B3').value = { formula: 'B2*2', result: 200 }

ws.getCell('A4').value = 'ToDelete'
ws.getCell('B4').value = 40

ws.mergeCells('A6:B6')
ws.getCell('A6').value = 'Merged banner'
ws.getCell('A6').alignment = { horizontal: 'center' }

ws.getColumn(1).width = 18
ws.getColumn(2).width = 12

await wb.xlsx.writeFile(new URL('../src/xlsx/__fixtures__/sample.xlsx', import.meta.url).pathname.replace(/^\/([A-Za-z]:)/, '$1'))
console.log('wrote src/xlsx/__fixtures__/sample.xlsx')
