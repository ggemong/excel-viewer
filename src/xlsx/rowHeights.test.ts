import ExcelJS from 'exceljs'
import { describe, expect, it } from 'vitest'
import { readDefaultFontPt } from './defaultFont'
import { customHeightRowsOf, readCustomHeightRows } from './rowHeights'
import { ZipArchive } from './zipReader'

describe('customHeightRowsOf', () => {
  it('높이를 직접 정한 행(customHeight)만 모은다 — 저장만 된 높이(ht)는 자동 높이일 수 있어 제외', () => {
    const xml = `<sheetData>
      <row r="3" spans="1:14"><c r="A3"/></row>
      <row r="4" ht="24" spans="1:14"><c r="A4"/></row>
      <row r="46" ht="72.75" customHeight="1" spans="1:14"><c r="A46"/></row>
      <row r="47" customHeight="true" ht="30"><c r="A47"/></row>
      <row r="48" customHeight="0" ht="30"><c r="A48"/></row>
    </sheetData>`
    expect([...customHeightRowsOf(xml)]).toEqual([46, 47])
  })

  it('속성 순서와 무관하게 행 번호를 찾는다', () => {
    expect([...customHeightRowsOf('<row ht="20" customHeight="1" r="9"><c r="A9"/></row>')]).toEqual([9])
  })

  it('행이 없으면 빈 집합', () => {
    expect(customHeightRowsOf('<sheetData/>').size).toBe(0)
  })
})

async function makeZip(setup: (ws: ExcelJS.Worksheet) => void): Promise<ZipArchive> {
  const wb = new ExcelJS.Workbook()
  setup(wb.addWorksheet('S'))
  const out = new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer)
  return ZipArchive.open(out.buffer.slice(out.byteOffset, out.byteOffset + out.byteLength) as ArrayBuffer)
}

describe('readCustomHeightRows', () => {
  it('실제 xlsx의 시트 XML에서 직접 정한 행을 읽는다', async () => {
    const zip = await makeZip((ws) => {
      ws.getCell('A1').value = 'x'
      ws.getCell('A2').value = 'y'
      ws.getRow(2).height = 40 // ExcelJS는 높이를 지정하면 customHeight="1"로 쓴다
    })
    expect([...(await readCustomHeightRows(zip, 'xl/worksheets/sheet1.xml'))]).toEqual([2])
  })

  it('시트 XML이 없으면 "직접 정한 행 없음"으로 뭉개지 않고 던진다', async () => {
    const zip = await makeZip((ws) => {
      ws.getCell('A1').value = 'x'
    })
    await expect(readCustomHeightRows(zip, 'xl/worksheets/nope.xml')).rejects.toThrow()
  })
})

describe('readDefaultFontPt', () => {
  it('기본 서식(cellXfs[0])이 가리키는 글꼴의 크기를 읽는다', async () => {
    const zip = await makeZip((ws) => {
      ws.getCell('A1').value = 'x'
      ws.getCell('A1').font = { size: 9 } // 칸에 쓴 글꼴이 기본 글꼴 크기로 오해되지 않아야 한다
    })
    // ExcelJS가 쓰는 기본 글꼴은 Calibri 11
    expect(await readDefaultFontPt(zip)).toBe(11)
  })
})
