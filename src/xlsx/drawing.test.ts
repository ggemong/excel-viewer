import { describe, expect, it } from 'vitest'
import { parseDrawingXml, type DrawingParseContext } from './drawing'
import type { PictureNode, ShapeNode, GroupNode } from './drawingTypes'
import type { ThemeColors } from './themeColor'

const NS =
  'xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'

const theme: ThemeColors = {
  slots: ['#FFFFFF', '#000000', '#EEECE1', '#1F497D', '#4F81BD', '#C0504D', '', '', '', '', '', ''],
  lineWidthsEmu: [9525, 25400, 38100],
}

const blob = new Blob(['x'], { type: 'image/png' })
const ctx: DrawingParseContext = {
  theme,
  media: (id) => (id === 'rId1' ? { blob, unsupportedFormat: null } : id === 'rId9' ? { blob: null, unsupportedFormat: 'emf' } : undefined),
}

const wrap = (inner: string) => `<xdr:wsDr ${NS}>${inner}</xdr:wsDr>`
const from = (col: number, row: number) =>
  `<xdr:from><xdr:col>${col}</xdr:col><xdr:colOff>9525</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>19050</xdr:rowOff></xdr:from>`
const to = (col: number, row: number) =>
  `<xdr:to><xdr:col>${col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>`

describe('그림', () => {
  const pic = (rid: string, extra = '', xfrm = '') =>
    `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="그림 1" descr="화면"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="${rid}"/>${extra}<a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr>${xfrm}<a:prstGeom prst="rect"/></xdr:spPr></xdr:pic>`

  it('두 셀 앵커 그림: 위치(셀+px 오프셋)와 데이터를 읽는다', () => {
    const { items } = parseDrawingXml(wrap(`<xdr:twoCellAnchor editAs="oneCell">${from(1, 2)}${to(11, 28)}${pic('rId1')}<xdr:clientData/></xdr:twoCellAnchor>`), ctx)
    expect(items).toHaveLength(1)
    expect(items[0].anchor).toEqual({
      kind: 'twoCell',
      from: { col: 1, row: 2, colOffset: 1, rowOffset: 2 },
      to: { col: 11, row: 28, colOffset: 0, rowOffset: 0 },
    })
    const node = items[0].node as PictureNode
    expect(node.kind).toBe('picture')
    expect(node.blob).toBe(blob)
    expect(node.alt).toBe('화면')
  })

  it('자르기(srcRect)는 비율로 읽는다 — ExcelJS가 버리던 정보', () => {
    const { items } = parseDrawingXml(
      wrap(`<xdr:twoCellAnchor>${from(0, 0)}${to(2, 2)}${pic('rId1', '<a:srcRect l="25000" r="10000" b="50000"/>')}<xdr:clientData/></xdr:twoCellAnchor>`),
      ctx,
    )
    expect((items[0].node as PictureNode).crop).toEqual({ l: 0.25, t: 0, r: 0.1, b: 0.5 })
  })

  it('회전/뒤집기', () => {
    const xfrm = '<a:xfrm rot="5400000" flipH="1"><a:off x="0" y="0"/><a:ext cx="10" cy="10"/></a:xfrm>'
    const { items } = parseDrawingXml(wrap(`<xdr:twoCellAnchor>${from(0, 0)}${to(1, 1)}${pic('rId1', '', xfrm)}<xdr:clientData/></xdr:twoCellAnchor>`), ctx)
    const node = items[0].node as PictureNode
    expect(node.rotation).toBe(90)
    expect(node.flipH).toBe(true)
    expect(node.flipV).toBe(false)
  })

  it('한 셀 앵커(크기 고정)와 절대 앵커', () => {
    const one = `<xdr:oneCellAnchor>${from(3, 1)}<xdr:ext cx="1905000" cy="952500"/>${pic('rId1')}<xdr:clientData/></xdr:oneCellAnchor>`
    const abs = `<xdr:absoluteAnchor><xdr:pos x="4762500" y="2857500"/><xdr:ext cx="1905000" cy="952500"/>${pic('rId1')}<xdr:clientData/></xdr:absoluteAnchor>`
    const { items } = parseDrawingXml(wrap(one + abs), ctx)
    expect(items[0].anchor).toMatchObject({ kind: 'oneCell', width: 200, height: 100 })
    expect(items[1].anchor).toEqual({ kind: 'absolute', x: 500, y: 300, width: 200, height: 100 })
  })

  it('브라우저가 못 그리는 형식(EMF)은 blob 없이 형식만 알린다', () => {
    const { items } = parseDrawingXml(wrap(`<xdr:twoCellAnchor>${from(0, 0)}${to(1, 1)}${pic('rId9')}<xdr:clientData/></xdr:twoCellAnchor>`), ctx)
    const node = items[0].node as PictureNode
    expect(node.blob).toBeNull()
    expect(node.unsupportedFormat).toBe('emf')
  })

  it('관계 ID가 깨진 그림은 건너뛰고 skipped에 센다', () => {
    const { items, skipped } = parseDrawingXml(wrap(`<xdr:twoCellAnchor>${from(0, 0)}${to(1, 1)}${pic('rId404')}<xdr:clientData/></xdr:twoCellAnchor>`), ctx)
    expect(items).toHaveLength(0)
    expect(skipped.picture).toBe(1)
  })
})

describe('도형', () => {
  const anchorWith = (sp: string) => wrap(`<xdr:twoCellAnchor>${from(0, 0)}${to(5, 5)}${sp}<xdr:clientData/></xdr:twoCellAnchor>`)
  const style =
    '<xdr:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="15000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></xdr:style>'

  it('빨간 점선 테두리 + 8% 투명 채우기 사각형 — 실제 요청서의 강조 상자', () => {
    const sp = `<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="7" name="사각형 6"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm flipV="1"><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="C00000"><a:alpha val="8000"/></a:srgbClr></a:solidFill><a:ln w="28575"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill><a:prstDash val="sysDash"/></a:ln></xdr:spPr>${style}</xdr:sp>`
    const node = parseDrawingXml(anchorWith(sp), ctx).items[0].node as ShapeNode
    expect(node.preset).toBe('rect')
    expect(node.fill).toBe('rgba(192, 0, 0, 0.08)')
    expect(node.line).toEqual({ color: '#C00000', width: 3, dash: 'dashed' })
    expect(node.flipV).toBe(true)
    expect(node.text).toBeNull()
  })

  it('명시된 채우기/선이 없으면 스타일 참조(accent1)를 따른다', () => {
    const sp = `<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="1" name="a"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"/></xdr:spPr>${style}</xdr:sp>`
    const node = parseDrawingXml(anchorWith(sp), ctx).items[0].node as ShapeNode
    expect(node.fill).toBe('#4F81BD')
    // lnRef idx=2 -> 테마 선 두께 2번째(25400 EMU ≈ 2.67px), shade로 어두워진 accent1
    expect(node.line?.width).toBeCloseTo(25400 / 9525, 3)
    expect(node.line?.color).not.toBe('#4F81BD')
  })

  it('noFill 선/채우기는 그리지 않는다', () => {
    const sp = `<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="1" name="a"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"/><a:noFill/><a:ln><a:noFill/></a:ln></xdr:spPr>${style}</xdr:sp>`
    const node = parseDrawingXml(anchorWith(sp), ctx).items[0].node as ShapeNode
    expect(node.fill).toBeNull()
    expect(node.line).toBeNull()
  })

  it('말풍선: 프리셋 이름과 조절값(adj)을 읽는다', () => {
    const sp = `<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="8" name="말풍선"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="wedgeRoundRectCallout"><a:avLst><a:gd name="adj1" fmla="val -144644"/><a:gd name="adj2" fmla="val 170620"/><a:gd name="adj3" fmla="val 16667"/></a:avLst></a:prstGeom></xdr:spPr>${style}<xdr:txBody><a:bodyPr anchor="t"/><a:p><a:pPr algn="l"/><a:r><a:rPr lang="ko-KR" sz="1100"/><a:t>수정</a:t></a:r></a:p></xdr:txBody></xdr:sp>`
    const node = parseDrawingXml(anchorWith(sp), ctx).items[0].node as ShapeNode
    expect(node.preset).toBe('wedgeRoundRectCallout')
    expect(node.adjust).toEqual({ adj1: -144644, adj2: 170620, adj3: 16667 })
  })

  it('텍스트: 크기·굵기·정렬·기본 글자색(스타일 fontRef lt1=흰색)', () => {
    const sp = `<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="8" name="t"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"/></xdr:spPr>${style}<xdr:txBody><a:bodyPr anchor="ctr" wrap="none" lIns="0"/><a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="1400" b="1"/><a:t>제목</a:t></a:r><a:br/><a:r><a:rPr sz="1100"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><a:t>빨강</a:t></a:r></a:p></xdr:txBody></xdr:sp>`
    const node = parseDrawingXml(anchorWith(sp), ctx).items[0].node as ShapeNode
    const text = node.text!
    expect(text.vAlign).toBe('middle')
    expect(text.wrap).toBe(false)
    expect(text.insets.l).toBe(0)
    expect(text.paragraphs[0].align).toBe('center')
    const [title, br, red] = text.paragraphs[0].runs
    expect(title).toMatchObject({ text: '제목', bold: true, color: '#FFFFFF' })
    expect(title.sizePx).toBeCloseTo((14 * 96) / 72, 3)
    expect(br.text).toBe('\n')
    expect(red.color).toBe('#FF0000')
  })

  it('글자가 하나도 없는 텍스트 본문은 text 없음으로', () => {
    const sp = `<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="8" name="t"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"/></xdr:spPr><xdr:txBody><a:bodyPr/><a:p><a:endParaRPr sz="1100"/></a:p></xdr:txBody></xdr:sp>`
    expect((parseDrawingXml(anchorWith(sp), ctx).items[0].node as ShapeNode).text).toBeNull()
  })

  it('회전한 화살표', () => {
    const sp = `<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="14" name="a"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm rot="13240301"><a:off x="0" y="0"/><a:ext cx="261456" cy="1199619"/></a:xfrm><a:prstGeom prst="downArrow"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:ln><a:noFill/></a:ln></xdr:spPr></xdr:sp>`
    const node = parseDrawingXml(anchorWith(sp), ctx).items[0].node as ShapeNode
    expect(node.preset).toBe('downArrow')
    expect(node.rotation).toBeCloseTo(13240301 / 60000, 3)
    expect(node.line).toBeNull()
  })
})

describe('묶음 / 대체 콘텐츠 / 미지원', () => {
  it('묶음: 자식 위치를 묶음 상자 대비 비율로 바꾼다', () => {
    const grp = `<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="16" name="그룹"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr><a:xfrm><a:off x="1000" y="2000"/><a:ext cx="4000" cy="2000"/><a:chOff x="1000" y="2000"/><a:chExt cx="4000" cy="2000"/></a:xfrm></xdr:grpSpPr>
      <xdr:sp><xdr:nvSpPr><xdr:cNvPr id="1" name="a"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="3000" y="3000"/><a:ext cx="1000" cy="500"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></xdr:spPr></xdr:sp></xdr:grpSp>`
    const { items } = parseDrawingXml(wrap(`<xdr:twoCellAnchor>${from(0, 0)}${to(5, 5)}${grp}<xdr:clientData/></xdr:twoCellAnchor>`), ctx)
    const group = items[0].node as GroupNode
    expect(group.kind).toBe('group')
    expect(group.children).toHaveLength(1)
    expect(group.children[0].rel).toEqual({ x: 0.5, y: 0.5, w: 0.25, h: 0.25 })
  })

  it('mc:AlternateContent는 Choice 내용을 풀어서 읽는다', () => {
    const sp = `<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="1" name="a"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="ellipse"/></xdr:spPr></xdr:sp>`
    const xml = wrap(`<xdr:twoCellAnchor>${from(0, 0)}${to(1, 1)}<mc:AlternateContent><mc:Choice Requires="a14">${sp}</mc:Choice><mc:Fallback/></mc:AlternateContent><xdr:clientData/></xdr:twoCellAnchor>`)
    expect((parseDrawingXml(xml, ctx).items[0].node as ShapeNode).preset).toBe('ellipse')
  })

  it('차트(graphicFrame)는 버리되 개수는 알려준다', () => {
    const gf = `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="차트"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"/></a:graphic></xdr:graphicFrame>`
    const { items, skipped } = parseDrawingXml(wrap(`<xdr:twoCellAnchor>${from(0, 0)}${to(5, 5)}${gf}<xdr:clientData/></xdr:twoCellAnchor>`), ctx)
    expect(items).toHaveLength(0)
    expect(skipped).toEqual({ chart: 1 })
  })

  it('깨진 XML은 조용히 넘기지 않고 오류로', () => {
    expect(() => parseDrawingXml('<xdr:wsDr', ctx)).toThrow()
  })
})
