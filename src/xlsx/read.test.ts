import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { readWorkbookDrawings } from './drawing'
import type { PictureNode, ShapeNode } from './drawingTypes'
import { EXCELJS_IGNORED_NODES, mergeRangesOf, readWorkbook } from './read'
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

  it('값이 있는 칸의 메모를 글자로 읽고, 메모 없는 칸에는 note가 없다', async () => {
    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet('S')
    ws.getCell('A1').value = '값 있는 칸'
    ws.getCell('A1').note = '이 칸 메모'
    ws.getCell('B1').value = '메모 없는 칸'
    const model = await readWorkbook(await toFile(wb))
    const rows = model.sheets[0].rows
    expect(rows[0][0]).toMatchObject({ value: '값 있는 칸', note: '이 칸 메모' })
    expect(rows[0][1]?.note).toBeUndefined()
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

describe('데이터 유효성 검사 건너뛰기', () => {
  // ExcelJS의 타입 선언에는 dataValidations가 빠져 있다(런타임에는 있다).
  const validationsOf = (ws: ExcelJS.Worksheet) => (ws as unknown as { dataValidations: { add(address: string, rule: unknown): void; model: Record<string, unknown> } }).dataValidations

  /** 열 전체에 드롭다운이 걸린 시트 — ExcelJS가 읽을 때 100만 칸으로 펼치는 실제 파일의 모양. */
  async function wholeColumnValidationBuffer(): Promise<ArrayBuffer> {
    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet('S')
    ws.getCell('A1').value = '보임'
    validationsOf(ws).add('B4:B1048576', { type: 'list', allowBlank: true, formulae: ['"가,나"'] })
    const out = new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer)
    return out.buffer.slice(out.byteOffset, out.byteOffset + out.byteLength) as ArrayBuffer
  }

  it('옵션을 주면 유효성 규칙을 펼치지 않는다(옵션이 실제로 먹는지 — 안 주면 펼쳐진다는 전제도 함께 확인)', async () => {
    const buffer = await wholeColumnValidationBuffer()

    const skipped = new ExcelJS.Workbook()
    await skipped.xlsx.load(buffer.slice(0), { ignoreNodes: EXCELJS_IGNORED_NODES })
    expect(Object.keys(validationsOf(skipped.getWorksheet('S')!).model)).toHaveLength(0)

    const expanded = new ExcelJS.Workbook()
    await expanded.xlsx.load(buffer.slice(0))
    expect(Object.keys(validationsOf(expanded.getWorksheet('S')!).model).length).toBeGreaterThan(1000)
  })

  it('건너뛰어도 칸 값은 그대로 읽힌다', async () => {
    const buffer = await wholeColumnValidationBuffer()
    const file = new File([buffer], 'v.xlsx')
    if (typeof file.arrayBuffer !== 'function') (file as unknown as { arrayBuffer: () => Promise<ArrayBuffer> }).arrayBuffer = async () => buffer
    const model = await readWorkbook(file)
    expect(model.sheets[0].rows[0][0]?.value).toBe('보임')
  })
})

describe('테마 색', () => {
  it('ZIP의 테마 파일에서 읽어 테마 색 채우기를 실제 색으로 해석한다', async () => {
    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet('S')
    ws.getCell('A1').value = '강조색'
    // theme 4 = accent1. ExcelJS가 쓰는 기본 Office 테마의 accent1은 4F81BD
    ws.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { theme: 4 } }
    const model = await readWorkbook(await toFile(wb))
    expect(model.sheets[0].rows[0][0]?.style?.bg?.toLowerCase()).toBe('#4f81bd')
    expect(model.warnings).toEqual([])
  })
})

describe('mergeRangesOf', () => {
  it('ExcelJS가 model로 주는 병합 범위와 같은 값을 돌려준다(느린 getter를 피하는 지름길이 어긋나지 않았는지)', async () => {
    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet('S')
    ws.getCell('B2').value = '제목'
    ws.mergeCells('B2:C3')
    ws.mergeCells('E5:F5')
    const loaded = new ExcelJS.Workbook()
    await loaded.xlsx.load((await wb.xlsx.writeBuffer()) as ArrayBuffer)
    const sheet = loaded.getWorksheet('S')!
    expect(mergeRangesOf(sheet).sort()).toEqual(['B2:C3', 'E5:F5'])
    expect(mergeRangesOf(sheet).sort()).toEqual([...(sheet.model.merges ?? [])].sort())
  })

  it('병합이 없으면 빈 목록', () => {
    const ws = new ExcelJS.Workbook().addWorksheet('S')
    expect(mergeRangesOf(ws)).toEqual([])
  })

  it('내부 구조를 못 읽으면 model로 되돌아간다', () => {
    const fake = { model: { merges: ['A1:B1'] } } as unknown as ExcelJS.Worksheet
    expect(mergeRangesOf(fake)).toEqual(['A1:B1'])
  })
})

describe('틀 고정', () => {
  async function frozenOf(views: Partial<ExcelJS.WorksheetView>[] | null) {
    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet('S')
    ws.getCell('A1').value = 1
    // ExcelJS의 타입은 보기 설정 전체 필드를 요구하지만 실제로는 일부만 줘도 된다.
    if (views) ws.views = views as ExcelJS.WorksheetView[]
    return (await readWorkbook(await toFile(wb))).sheets[0].frozen
  }

  it('행과 열을 함께 고정하면 둘 다 읽는다', async () => {
    expect(await frozenOf([{ state: 'frozen', xSplit: 2, ySplit: 1 }])).toEqual({ rows: 1, cols: 2 })
  })

  it('열만 고정해도 고정으로 읽는다(행 수 0)', async () => {
    expect(await frozenOf([{ state: 'frozen', xSplit: 1, ySplit: 0 }])).toEqual({ rows: 0, cols: 1 })
  })

  it('행만 고정하면 열 수 0', async () => {
    expect(await frozenOf([{ state: 'frozen', xSplit: 0, ySplit: 3 }])).toEqual({ rows: 3, cols: 0 })
  })

  it('고정이 없으면 null', async () => {
    expect(await frozenOf(null)).toBeNull()
  })
})

describe('차트 읽기(드로잉 → 차트 부품)', () => {
  const C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
  const barChartXml = (group: string) =>
    `<?xml version="1.0"?><c:chartSpace xmlns:c="${C_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>분기별</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea>${group}</c:plotArea></c:chart></c:chartSpace>`
  const bar = '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/><c:cat><c:strRef><c:f>S!$A$1</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>가</c:v></c:pt><c:pt idx="1"><c:v>나</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>S!$B$1</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>7</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:barChart>'
  const area = '<c:areaChart><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numRef><c:f>S!$B$1</c:f><c:numCache><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:areaChart>'

  async function withCharts(charts: string[]) {
    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet('S')
    const id = wb.addImage({ buffer: PNG_1X1 as unknown as ExcelJS.Buffer, extension: 'png' })
    ws.addImage(id, { tl: { col: 0, row: 0 }, ext: { width: 10, height: 10 } })
    return readWorkbook(
      await toFile(wb, async (zip) => {
        const drawingPath = 'xl/drawings/drawing1.xml'
        const relsPath = 'xl/drawings/_rels/drawing1.xml.rels'
        let drawing = await zip.file(drawingPath)!.async('string')
        let rels = await zip.file(relsPath)!.async('string')
        charts.forEach((xml, i) => {
          zip.file(`xl/charts/chart${i + 1}.xml`, xml)
          rels = rels.replace('</Relationships>', `<Relationship Id="rIdChart${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart${i + 1}.xml"/></Relationships>`)
          const frame =
            `<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${2 + i * 10}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${10 + i * 10}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>` +
            `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${20 + i}" name="차트"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>` +
            `<a:graphic><a:graphicData uri="${C_NS}"><c:chart xmlns:c="${C_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdChart${i + 1}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`
          drawing = drawing.replace('</xdr:wsDr>', `${frame}</xdr:wsDr>`)
        })
        zip.file(drawingPath, drawing)
        zip.file(relsPath, rels)
      }),
    )
  }

  it('차트 부품을 따라가 모델로 읽고 앵커 위치도 갖는다', async () => {
    const model = await withCharts([barChartXml(bar)])
    const sheet = model.sheets[0]
    const chart = sheet.drawings.find((d) => d.node.kind === 'chart')
    expect(chart).toBeDefined()
    const node = chart!.node as import('./chartTypes').ChartNode
    expect(node).toMatchObject({ family: 'bar', categories: ['가', '나'] })
    expect(node.title?.text).toBe('분기별')
    expect(node.series[0].values).toEqual([3, 7])
    expect(chart!.anchor).toMatchObject({ kind: 'twoCell', from: { col: 1, row: 2 }, to: { col: 6, row: 10 } })
    expect(sheet.skippedDrawings).toEqual({})
    expect(model.warnings).toEqual([])
  })

  it('못 그리는 차트는 사유별로 세어 알리고, 그릴 수 있는 차트는 그대로 읽는다', async () => {
    const model = await withCharts([barChartXml(bar), barChartXml(area)])
    const sheet = model.sheets[0]
    expect(sheet.drawings.filter((d) => d.node.kind === 'chart')).toHaveLength(1)
    expect(sheet.skippedDrawings).toEqual({ 'chart:area': 1 })
  })

  it('차트가 시트 크기를 늘린다(차트가 걸친 영역까지 빈 칸으로 이어진다)', async () => {
    const model = await withCharts([barChartXml(bar)])
    expect(model.sheets[0].rowCount).toBeGreaterThanOrEqual(10)
    expect(model.sheets[0].colCount).toBeGreaterThanOrEqual(6)
  })
})

describe('수식으로 만든 링크(HYPERLINK)', () => {
  async function linkCells() {
    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet('목록')
    wb.addWorksheet('참조 1')
    ws.getCell('A1').value = { formula: `HYPERLINK("#'참조 1'!B3","참조 1")`, result: '참조 1' }
    ws.getCell('A2').value = { formula: `HYPERLINK("https://a.example/x","열기")`, result: '열기' }
    ws.getCell('A3').value = { formula: `HYPERLINK("javascript:alert(1)","나쁨")`, result: '나쁨' }
    ws.getCell('A4').value = { formula: `HYPERLINK("#'"&A1&"'!A1","계산된 대상")`, result: '계산된 대상' }
    ws.getCell('A5').value = { text: '직접 링크', hyperlink: 'https://direct.example' }
    return (await readWorkbook(await toFile(wb))).sheets[0].rows
  }

  it('시트 이동 링크 수식은 대상 시트와 칸을 갖고, 보이는 글자는 수식 결과다', async () => {
    const cell = (await linkCells())[0][0]!
    expect(cell.value).toBe('참조 1')
    expect(cell.internalLink).toEqual({ sheet: '참조 1', row: 3, col: 2 })
    expect(cell.hyperlink).toBeNull()
  })

  it('외부 주소 수식은 허용된 스킴만 링크가 된다', async () => {
    const rows = await linkCells()
    expect(rows[1][0]!.hyperlink).toBe('https://a.example/x')
    expect(rows[2][0]!.hyperlink).toBeNull()
    expect(rows[2][0]!.internalLink).toBeUndefined()
  })

  it('대상이 계산으로 만들어지는 수식은 지금 가리킬 곳을 모르므로 "눌렀을 때 계산"으로만 표시한다', async () => {
    const cell = (await linkCells())[3][0]!
    expect(cell.internalLink).toBeUndefined()
    expect(cell.hyperlink).toBeNull()
    expect(cell.computedLink).toBe(true)
  })

  it('허용되지 않은 주소의 리터럴 수식은 계산 링크도 아니다(링크가 아예 아니다)', async () => {
    const cell = (await linkCells())[2][0]!
    expect(cell.computedLink).toBeUndefined()
  })

  it('계산 결과가 비어 보이는 링크 수식 칸(IFERROR가 빈 글자를 낸 경우)은 링크로 만들지 않는다', async () => {
    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet('S')
    ws.getCell('A1').value = { formula: `IFERROR(HYPERLINK("#"&B1&"!A1","이동"),"")`, result: '' }
    const cell = (await readWorkbook(await toFile(wb))).sheets[0].rows[0][0]
    expect(cell?.computedLink).toBeUndefined()
  })

  it('칸에 직접 건 외부 링크는 그대로 읽힌다', async () => {
    expect((await linkCells())[4][0]!.hyperlink).toBe('https://direct.example')
  })
})
