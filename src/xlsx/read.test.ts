import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { readWorkbookDrawings } from './drawing'
import type { PictureNode, ShapeNode } from './drawingTypes'
import { readWorkbook } from './read'
import { ZipArchive } from './zipReader'

/** 1x1 투명 PNG */
const PNG_1X1 = Uint8Array.from(
  atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='),
  (c) => c.charCodeAt(0),
)

async function toFile(wb: ExcelJS.Workbook, patch?: (zip: JSZip) => Promise<void>): Promise<File> {
  const buf = await wb.xlsx.writeBuffer()
  let bytes: Uint8Array = new Uint8Array(buf as ArrayBuffer)
  if (patch) {
    const zip = await JSZip.loadAsync(bytes)
    await patch(zip)
    bytes = await zip.generateAsync({ type: 'uint8array' })
  }
  const file = new File([bytes as BlobPart], 't.xlsx')
  // jsdom의 File에는 arrayBuffer()가 없을 수 있어 실제 브라우저와 같게 맞춰 준다.
  if (typeof file.arrayBuffer !== 'function') {
    ;(file as unknown as { arrayBuffer: () => Promise<ArrayBuffer> }).arrayBuffer = async () =>
      bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer
  }
  return file
}

describe('readWorkbook (xlsx)', () => {
  it('숨김/완전숨김 시트는 제외한다', async () => {
    const wb = new ExcelJS.Workbook()
    wb.addWorksheet('보임').getCell('A1').value = 1
    wb.addWorksheet('숨김').state = 'hidden'
    wb.addWorksheet('완전숨김').state = 'veryHidden'
    const model = await readWorkbook(await toFile(wb))
    expect(model.sheets.map((s) => s.name)).toEqual(['보임'])
  })

  it('공유 수식의 따라가는 칸도 계산값이 읽힌다', async () => {
    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet('S')
    ws.getCell('A1').value = 1
    ws.getCell('A2').value = 5
    ws.getCell('B1').value = { formula: 'A1*2', result: 2 }
    const model = await readWorkbook(
      await toFile(wb, async (zip) => {
        const path = 'xl/worksheets/sheet1.xml'
        let xml = await zip.file(path)!.async('string')
        // Excel이 채우기로 저장하는 모양: 마스터(ref+si) + 따라가는 칸(si만)
        xml = xml.replace(/<c r="B1"[^>]*>.*?<\/c>/, '<c r="B1"><f t="shared" ref="B1:B2" si="0">A1*2</f><v>2</v></c>')
        xml = xml.replace(/(<row r="2"[^>]*>.*?)(<\/row>)/, '$1<c r="B2"><f t="shared" si="0"/><v>10</v></c>$2')
        zip.file(path, xml)
      }),
    )
    const sheet = model.sheets[0]
    expect(sheet.rows[0][1]?.value).toBe(2)
    expect(sheet.rows[1][1]?.value).toBe(10)
  })

  it('명시 안 된 행높이/열너비는 파일의 sheetFormatPr 기본값을 쓴다(15pt/8.43 고정이 아님)', async () => {
    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet('S')
    ws.properties.defaultRowHeight = 16.5
    ws.getCell('A1').value = 'x'
    ws.getCell('A2').value = 'y'
    const model = await readWorkbook(await toFile(wb))
    expect(model.sheets[0].rowHeights[1]).toBeCloseTo((16.5 * 96) / 72, 5)
  })

  it('그림: 위치(앵커)와 데이터를 읽고, 데이터 영역 밖까지 걸치면 시트 크기를 늘린다', async () => {
    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet('그림시트')
    const id = wb.addImage({ buffer: PNG_1X1 as unknown as ExcelJS.Buffer, extension: 'png' })
    ws.addImage(id, 'B3:H30') // 데이터는 하나도 없음. B3(열1,행2)~H30의 먼 모서리(열8,행30)
    const model = await readWorkbook(await toFile(wb))
    const sheet = model.sheets[0]
    expect(model.warnings).toEqual([])
    expect(sheet.drawings).toHaveLength(1)
    const pic = sheet.drawings[0].node as PictureNode
    expect(pic.kind).toBe('picture')
    expect(pic.blob).toBeInstanceOf(Blob)
    expect(sheet.drawings[0].anchor).toMatchObject({ kind: 'twoCell', from: { col: 1, row: 2 }, to: { col: 8, row: 30 } })
    // 데이터 0행이어도 그림이 걸친 30행/8열까지는 빈 칸으로 이어진다
    expect(sheet.rowCount).toBeGreaterThanOrEqual(30)
    expect(sheet.colCount).toBeGreaterThanOrEqual(8)
    expect(sheet.rowHeights).toHaveLength(sheet.rowCount)
    expect(sheet.colWidths).toHaveLength(sheet.colCount)
  })

  it('ExcelJS가 버리는 도형(sp)도 원본 XML에서 읽는다', async () => {
    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet('S')
    const id = wb.addImage({ buffer: PNG_1X1 as unknown as ExcelJS.Buffer, extension: 'png' })
    ws.addImage(id, { tl: { col: 0, row: 0 }, ext: { width: 10, height: 10 } })
    const model = await readWorkbook(
      await toFile(wb, async (zip) => {
        const path = 'xl/drawings/drawing1.xml'
        const xml = await zip.file(path)!.async('string')
        const shape =
          '<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>' +
          '<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="9" name="r"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></xdr:spPr></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>'
        zip.file(path, xml.replace('</xdr:wsDr>', `${shape}</xdr:wsDr>`))
      }),
    )
    const kinds = model.sheets[0].drawings.map((d) => d.node.kind).sort()
    expect(kinds).toEqual(['picture', 'shape'])
    const shape = model.sheets[0].drawings.find((d) => d.node.kind === 'shape')!.node as ShapeNode
    expect(shape.fill).toBe('#C00000')
  })

  it('자동 필터: 범위와, ExcelJS가 버리는 조건이 걸린 열을 읽는다', async () => {
    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet('S')
    for (let r = 1; r <= 5; r++) {
      ws.getCell(r, 1).value = r % 2 ? 'A' : 'B'
      ws.getCell(r, 2).value = r
    }
    ws.autoFilter = 'A1:C5'
    const model = await readWorkbook(
      await toFile(wb, async (zip) => {
        const path = 'xl/worksheets/sheet1.xml'
        const xml = await zip.file(path)!.async('string')
        // 2번째 열(colId=1)에만 조건이 걸린 상태를 Excel이 저장하는 모양 그대로
        zip.file(path, xml.replace('<autoFilter ref="A1:C5"/>', '<autoFilter ref="A1:C5"><filterColumn colId="1"><filters><filter val="1"/></filters></filterColumn></autoFilter>'))
      }),
    )
    expect(model.sheets[0].filters).toEqual([{ headerRow: 1, lastRow: 5, firstCol: 1, lastCol: 3, activeCols: [2], hiddenButtonCols: [] }])
  })

  it('자동 필터가 없는 시트는 filters가 비어 있다', async () => {
    const wb = new ExcelJS.Workbook()
    wb.addWorksheet('S').getCell('A1').value = 1
    expect((await readWorkbook(await toFile(wb))).sheets[0].filters).toEqual([])
  })

  it('한 시트의 그림 XML이 깨져도 다른 시트는 계속 읽고, 실패는 경고로 남긴다', async () => {
    const wb = new ExcelJS.Workbook()
    const id = wb.addImage({ buffer: PNG_1X1 as unknown as ExcelJS.Buffer, extension: 'png' })
    wb.addWorksheet('정상').addImage(id, { tl: { col: 0, row: 0 }, ext: { width: 10, height: 10 } })
    wb.addWorksheet('깨짐').addImage(id, { tl: { col: 0, row: 0 }, ext: { width: 10, height: 10 } })

    const zip = await JSZip.loadAsync(new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer))
    zip.file('xl/drawings/drawing2.xml', '<xdr:wsDr') // 두 번째 시트의 drawing만 망가뜨린다
    const bytes = await zip.generateAsync({ type: 'uint8array' })

    const archive = ZipArchive.open(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer)
    const { bySheet, warnings } = await readWorkbookDrawings(archive, null)
    expect(bySheet.get('정상')?.items).toHaveLength(1)
    expect(bySheet.has('깨짐')).toBe(false)
    expect(warnings).toHaveLength(1)
    expect(warnings[0]).toContain('깨짐')
  })
})
