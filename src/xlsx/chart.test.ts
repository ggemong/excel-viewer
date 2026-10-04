import { describe, expect, it } from 'vitest'
import { parseChartXml } from './chart'
import type { ChartNode } from './chartTypes'
import type { ThemeColors } from './themeColor'

const NS = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

const strCache = (items: string[]) => `<c:strCache><c:ptCount val="${items.length}"/>${items.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:strCache>`
const numCache = (items: (number | null)[], fmt = 'General') =>
  `<c:numCache><c:formatCode>${fmt}</c:formatCode><c:ptCount val="${items.length}"/>${items.map((v, i) => (v === null ? '' : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)).join('')}</c:numCache>`

interface SerOpts {
  name?: string
  cats?: string[]
  values: (number | null)[]
  extra?: string
  fmt?: string
}
const ser = (i: number, o: SerOpts) =>
  `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>` +
  (o.name ? `<c:tx><c:strRef><c:f>S!$A$1</c:f>${strCache([o.name])}</c:strRef></c:tx>` : '') +
  (o.extra ?? '') +
  (o.cats ? `<c:cat><c:strRef><c:f>S!$A$2</c:f>${strCache(o.cats)}</c:strRef></c:cat>` : '') +
  `<c:val><c:numRef><c:f>S!$B$2</c:f>${numCache(o.values, o.fmt)}</c:numRef></c:val></c:ser>`

function chartXml(plot: string, extra = '', space = ''): string {
  return `<?xml version="1.0"?><c:chartSpace ${NS}><c:chart>${extra}<c:plotArea>${plot}</c:plotArea></c:chart>${space}</c:chartSpace>`
}

const AXES =
  '<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:majorGridlines/></c:valAx>'
const BAR_OPEN = '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>'

function parse(xml: string, theme: ThemeColors | null = null): ChartNode {
  const result = parseChartXml(xml, theme)
  if (!result.ok) throw new Error(`parse failed: ${result.reason}`)
  return result.node
}

describe('parseChartXml: 종류', () => {
  it('세로 막대(묶음)를 읽는다 — 제목, 카테고리, 값, 시리즈 색', () => {
    const node = parse(
      chartXml(
        `${BAR_OPEN}${ser(0, { name: '매출', cats: ['1월', '2월', '3월'], values: [10, 20, 15], extra: '<c:spPr><a:solidFill><a:srgbClr val="4F81BD"/></a:solidFill></c:spPr>' })}<c:gapWidth val="80"/><c:overlap val="10"/></c:barChart>${AXES}`,
        '<c:title><c:tx><c:rich><a:p><a:r><a:rPr sz="1800" b="1"><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:rPr><a:t>월별 매출</a:t></a:r></a:p></c:rich></c:tx></c:title>',
      ),
    )
    expect(node).toMatchObject({ family: 'bar', barDirection: 'col', grouping: 'clustered', gapWidth: 80, overlap: 10 })
    expect(node.title).toMatchObject({ text: '월별 매출', bold: true, color: '#112233' })
    expect(node.title?.sizePx).toBeCloseTo(24, 5)
    expect(node.categories).toEqual(['1월', '2월', '3월'])
    expect(node.series[0]).toMatchObject({ name: '매출', values: [10, 20, 15], color: '#4F81BD' })
  })

  it('가로 막대와 누적/100% 누적', () => {
    const horizontal = parse(chartXml(`<c:barChart><c:barDir val="bar"/><c:grouping val="stacked"/>${ser(0, { values: [1, 2] })}${ser(1, { values: [3, 4] })}</c:barChart>${AXES}`))
    expect(horizontal).toMatchObject({ barDirection: 'bar', grouping: 'stacked', overlap: 100 })
    expect(horizontal.series).toHaveLength(2)
    const percent = parse(chartXml(`<c:barChart><c:barDir val="col"/><c:grouping val="percentStacked"/>${ser(0, { values: [1, 2] })}</c:barChart>${AXES}`))
    expect(percent.grouping).toBe('percentStacked')
  })

  it('선: 부드러운 선, 표식 없음, 선 굵기와 색', () => {
    const line = ser(0, {
      cats: ['a', 'b'],
      values: [1, 2],
      extra: '<c:spPr><a:ln w="28575"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></c:spPr><c:marker><c:symbol val="none"/></c:marker>',
    }).replace('</c:ser>', '<c:smooth val="1"/></c:ser>')
    const node = parse(chartXml(`<c:lineChart><c:grouping val="standard"/>${line}<c:marker val="1"/></c:lineChart>${AXES}`))
    expect(node.family).toBe('line')
    expect(node.series[0]).toMatchObject({ color: '#FF0000', marker: false, smooth: true })
    expect(node.series[0].lineWidthPx).toBeCloseTo(3, 5)
  })

  it('선: 표식 요소가 없으면 그룹 설정(marker)을 따른다', () => {
    const withMarkers = parse(chartXml(`<c:lineChart><c:grouping val="standard"/>${ser(0, { values: [1, 2] })}<c:marker val="1"/></c:lineChart>${AXES}`))
    expect(withMarkers.series[0].marker).toBe(true)
    const without = parse(chartXml(`<c:lineChart><c:grouping val="standard"/>${ser(0, { values: [1, 2] })}<c:marker val="0"/></c:lineChart>${AXES}`))
    expect(without.series[0].marker).toBe(false)
  })

  it('원형: 첫 시리즈만, 조각 색, 시작 각도 / 도넛: 구멍 크기', () => {
    const slice = '<c:dPt><c:idx val="1"/><c:spPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></c:spPr></c:dPt>'
    const pie = parse(chartXml(`<c:pieChart><c:varyColors val="1"/>${ser(0, { cats: ['가', '나'], values: [30, 70], extra: slice })}${ser(1, { values: [1, 1] })}<c:firstSliceAng val="90"/></c:pieChart>`))
    expect(pie).toMatchObject({ family: 'pie', firstSliceAngle: 90 })
    expect(pie.series).toHaveLength(1)
    expect(pie.series[0].pointColors).toEqual({ 1: '#00FF00' })
    const donut = parse(chartXml(`<c:doughnutChart>${ser(0, { values: [1, 2, 3] })}<c:holeSize val="60"/></c:doughnutChart>`))
    expect(donut).toMatchObject({ family: 'doughnut', holeSize: 60 })
  })
})

describe('parseChartXml: 못 그리는 차트는 사유와 함께 돌려준다', () => {
  const reason = (plot: string) => {
    const result = parseChartXml(chartXml(plot), null)
    return result.ok ? 'ok' : result.reason
  }

  it('영역형, 분산형, 레이더, 3D, 콤보, 선 누적', () => {
    expect(reason(`<c:areaChart>${ser(0, { values: [1] })}</c:areaChart>`)).toBe('chart:area')
    expect(reason(`<c:scatterChart>${ser(0, { values: [1] })}</c:scatterChart>`)).toBe('chart:scatter')
    expect(reason(`<c:radarChart>${ser(0, { values: [1] })}</c:radarChart>`)).toBe('chart:radar')
    expect(reason(`<c:bar3DChart>${ser(0, { values: [1] })}</c:bar3DChart>`)).toBe('chart:3d')
    expect(reason(`<c:barChart>${ser(0, { values: [1] })}</c:barChart><c:lineChart>${ser(1, { values: [1] })}</c:lineChart>`)).toBe('chart:combo')
    expect(reason(`<c:lineChart><c:grouping val="stacked"/>${ser(0, { values: [1] })}</c:lineChart>`)).toBe('chart:other')
  })

  it('저장된 값(캐시)이 없으면 nodata — 빈 차트를 그리지 않는다', () => {
    expect(reason('<c:barChart><c:ser><c:idx val="0"/><c:val><c:numRef><c:f>S!$A$1</c:f></c:numRef></c:val></c:ser></c:barChart>')).toBe('chart:nodata')
    expect(reason(`<c:barChart>${ser(0, { values: [null, null] })}</c:barChart>`)).toBe('chart:nodata')
  })

  it('깨진 XML이나 플롯 영역이 없는 XML은 던지지 않고 chart', () => {
    expect(parseChartXml('<not xml', null)).toEqual({ ok: false, reason: 'chart' })
    expect(parseChartXml(`<c:chartSpace ${NS}><c:chart/></c:chartSpace>`, null)).toEqual({ ok: false, reason: 'chart' })
  })
})

describe('parseChartXml: 값과 카테고리', () => {
  it('빈 점은 null로 읽고, 카테고리가 없으면 1,2,3…', () => {
    const node = parse(chartXml(`<c:lineChart><c:grouping val="standard"/>${ser(0, { values: [5, null, 7] })}</c:lineChart>${AXES}`))
    expect(node.series[0].values).toEqual([5, null, 7])
    expect(node.categories).toEqual(['1', '2', '3'])
  })

  it('시리즈가 짧으면 카테고리 수에 맞춰 null로 채운다', () => {
    const node = parse(chartXml(`${BAR_OPEN}${ser(0, { cats: ['a', 'b', 'c'], values: [1] })}</c:barChart>${AXES}`))
    expect(node.series[0].values).toEqual([1, null, null])
  })

  it('숫자(날짜) 카테고리는 캐시의 표시 형식대로 글자로 만든다', () => {
    const cat = `<c:cat><c:numRef><c:f>S!$A$2</c:f>${numCache([46299], 'yyyy-mm-dd')}</c:numRef></c:cat>`
    const val = `<c:val><c:numRef><c:f>S!$B$2</c:f>${numCache([3])}</c:numRef></c:val>`
    const xml = chartXml(`<c:lineChart><c:grouping val="standard"/><c:ser><c:idx val="0"/><c:order val="0"/>${cat}${val}</c:ser></c:lineChart>${AXES}`)
    expect(parse(xml).categories).toEqual(['2026-10-04'])
  })

  it('시리즈는 order 순서로 정렬한다', () => {
    const a = ser(0, { name: 'A', values: [1] }).replace('<c:order val="0"/>', '<c:order val="1"/>')
    const b = ser(1, { name: 'B', values: [2] }).replace('<c:order val="1"/>', '<c:order val="0"/>')
    const node = parse(chartXml(`${BAR_OPEN}${a}${b}</c:barChart>${AXES}`))
    expect(node.series.map((s) => s.name)).toEqual(['B', 'A'])
  })

  it('값의 표시 형식 코드를 시리즈에 담는다', () => {
    const node = parse(chartXml(`${BAR_OPEN}${ser(0, { values: [0.5], fmt: '0.0%' })}</c:barChart>${AXES}`))
    expect(node.series[0].formatCode).toBe('0.0%')
  })
})

describe('parseChartXml: 축·범례·레이블·영역', () => {
  const base = (axes: string, extra = '', plotExtra = '') => chartXml(`${BAR_OPEN}${ser(0, { values: [1, 2] })}${plotExtra}</c:barChart>${axes}`, extra)

  it('값 축 범위·눈금 간격·가로줄·숫자 형식, 항목 축 제목', () => {
    const axes =
      '<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="maxMin"/></c:scaling><c:delete val="0"/><c:title><c:tx><c:rich><a:p><a:r><a:t>분기</a:t></a:r></a:p></c:rich></c:tx></c:title></c:catAx>' +
      '<c:valAx><c:axId val="2"/><c:scaling><c:min val="0"/><c:max val="100"/></c:scaling><c:delete val="0"/><c:majorGridlines/><c:numFmt formatCode="0%" sourceLinked="0"/><c:majorUnit val="25"/><c:crossBetween val="midCat"/></c:valAx>'
    const node = parse(base(axes))
    expect(node.valueAxis).toMatchObject({ min: 0, max: 100, majorUnit: 25, gridlines: true, formatCode: '0%', hidden: false })
    expect(node.categoryAxis).toMatchObject({ reversed: true, hidden: false })
    expect(node.categoryAxis.title?.text).toBe('분기')
    expect(node.pointsOnTicks).toBe(true)
  })

  it('sourceLinked=1이면 축 숫자 형식은 값을 따른다(null), 축이 삭제(delete=1)되면 hidden', () => {
    const axes = '<c:catAx><c:axId val="1"/><c:scaling/><c:delete val="1"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling/><c:delete val="0"/><c:numFmt formatCode="0.0" sourceLinked="1"/></c:valAx>'
    const node = parse(base(axes))
    expect(node.valueAxis.formatCode).toBeNull()
    expect(node.categoryAxis.hidden).toBe(true)
  })

  it('원형처럼 축 요소가 없는 차트는 축을 숨긴다', () => {
    const node = parse(chartXml(`<c:pieChart>${ser(0, { values: [1, 2] })}</c:pieChart>`))
    expect(node.valueAxis.hidden).toBe(true)
    expect(node.categoryAxis.hidden).toBe(true)
  })

  it('범례 위치 (기본 오른쪽), 범례가 없으면 null', () => {
    expect(parse(base(AXES, '<c:legend><c:legendPos val="b"/></c:legend>')).legend?.position).toBe('b')
    expect(parse(base(AXES, '<c:legend/>')).legend?.position).toBe('r')
    expect(parse(base(AXES)).legend).toBeNull()
  })

  it('데이터 레이블: 켜진 항목, 구분자, 형식 / 모두 꺼져 있으면 null / 시리즈 설정이 그룹 설정보다 우선', () => {
    const labels = (flags: string) => `<c:dLbls><c:numFmt formatCode="0.0" sourceLinked="0"/><c:separator> | </c:separator>${flags}</c:dLbls>`
    const on = parse(chartXml(`${BAR_OPEN}${ser(0, { values: [1] })}${labels('<c:showVal val="1"/><c:showCatName val="1"/>')}</c:barChart>${AXES}`))
    expect(on.series[0].labels).toMatchObject({ showValue: true, showCategory: true, showPercent: false, separator: ' | ', formatCode: '0.0' })

    const off = parse(chartXml(`${BAR_OPEN}${ser(0, { values: [1] })}${labels('<c:showVal val="0"/>')}</c:barChart>${AXES}`))
    expect(off.series[0].labels).toBeNull()

    const own = ser(0, { values: [1], extra: '<c:dLbls><c:showVal val="0"/></c:dLbls>' })
    const ownOff = parse(chartXml(`${BAR_OPEN}${own}${labels('<c:showVal val="1"/>')}</c:barChart>${AXES}`))
    expect(ownOff.series[0].labels).toBeNull()
  })

  it('차트 영역: 배경색(없으면 흰색, noFill이면 투명)과 테두리', () => {
    const space = '<c:spPr><a:solidFill><a:srgbClr val="F0F0F0"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr>'
    const pie = `<c:pieChart>${ser(0, { values: [1] })}</c:pieChart>`
    const filled = parse(chartXml(pie, '', space))
    expect(filled.background).toBe('#F0F0F0')
    expect(filled.border).toEqual({ color: '#D9D9D9', widthPx: 1 })
    expect(parse(chartXml(pie)).background).toBe('#FFFFFF')
    expect(parse(chartXml(pie, '', '<c:spPr><a:noFill/></c:spPr>')).background).toBeNull()
  })

  it('테마 색(schemeClr)은 파일의 테마로 해석한다', () => {
    const theme: ThemeColors = { slots: ['#FFFFFF', '#000000', '#EEEEEE', '#111111', '#C0504D', '#9BBB59', '', '', '', '', '', ''] }
    const fill = '<c:spPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></c:spPr>'
    const node = parse(chartXml(`${BAR_OPEN}${ser(0, { values: [1], extra: fill })}</c:barChart>${AXES}`), theme)
    expect(node.series[0].color).toBe('#C0504D')
  })

  it('빈 점 처리(dispBlanksAs)', () => {
    expect(parse(base(AXES, '<c:dispBlanksAs val="zero"/>')).blanksAs).toBe('zero')
    expect(parse(base(AXES, '<c:dispBlanksAs val="span"/>')).blanksAs).toBe('gap')
  })
})
