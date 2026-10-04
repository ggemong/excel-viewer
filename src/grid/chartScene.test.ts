import { describe, expect, it } from 'vitest'
import type { ChartAxis, ChartNode } from '../xlsx/chartTypes'
import { buildChartScene, type MeasureText, type SceneItem } from './chartScene'

/** 글자당 폭 = 크기의 절반인 가짜 측정기 */
const measure: MeasureText = (text, size) => text.length * size * 0.5

const axis = (over: Partial<ChartAxis> = {}): ChartAxis => ({
  hidden: true, title: null, formatCode: null, sizePx: 12, color: null, lineColor: null, gridlines: false, gridColor: null, min: null, max: null, majorUnit: null, reversed: false, ...over,
})

const PALETTE = ['#111111', '#222222', '#333333']

function chartOf(series: (number | null)[][], over: Partial<ChartNode> = {}): ChartNode {
  return {
    kind: 'chart', rotation: 0, flipH: false, flipV: false,
    family: 'bar', barDirection: 'col', grouping: 'clustered', title: null,
    categories: series[0].map((_, i) => `c${i}`),
    series: series.map((values) => ({ name: null, values, color: null, pointColors: {}, lineWidthPx: null, marker: false, smooth: false, labels: null, formatCode: null })),
    palette: PALETTE, legend: null, categoryAxis: axis(), valueAxis: axis(),
    gapWidth: 150, overlap: 0, pointsOnTicks: false, holeSize: 50, firstSliceAngle: 0, blanksAs: 'gap', background: null, border: null,
    ...over,
  }
}

const scene = (chart: ChartNode, w = 400, h = 300) => buildChartScene(chart, w, h, measure)
const rects = (items: SceneItem[], fill: string) => items.filter((i): i is Extract<SceneItem, { t: 'rect' }> => i.t === 'rect' && i.fill === fill)
const texts = (items: SceneItem[]) => items.filter((i): i is Extract<SceneItem, { t: 'text' }> => i.t === 'text')
const paths = (items: SceneItem[]) => items.filter((i): i is Extract<SceneItem, { t: 'path' }> => i.t === 'path')

describe('buildChartScene: 막대', () => {
  it('막대 높이는 값에 비례하고 바닥(기준선)이 같다', () => {
    const bars = rects(scene(chartOf([[10, 20]])).items, PALETTE[0])
    expect(bars).toHaveLength(2)
    expect(bars[1].h / bars[0].h).toBeCloseTo(2, 5)
    expect(bars[0].y + bars[0].h).toBeCloseTo(bars[1].y + bars[1].h, 5)
  })

  it('시리즈 색이 없으면 팔레트 순서, 점 색이 있으면 그 점만 다른 색', () => {
    const chart = chartOf([[1, 2], [3, 4]])
    chart.series[1].pointColors = { 1: '#abcdef' }
    const items = scene(chart).items
    expect(rects(items, PALETTE[0])).toHaveLength(2)
    expect(rects(items, PALETTE[1])).toHaveLength(1)
    expect(rects(items, '#abcdef')).toHaveLength(1)
  })

  it('묶음: 같은 항목의 시리즈 막대가 옆으로 나란히, 서로 겹치지 않는다', () => {
    const items = scene(chartOf([[5], [5]])).items
    const [a] = rects(items, PALETTE[0])
    const [b] = rects(items, PALETTE[1])
    expect(b.x).toBeGreaterThanOrEqual(a.x + a.w - 1e-6)
  })

  it('누적: 둘째 막대가 첫째 막대 위에서 시작한다', () => {
    const items = scene(chartOf([[3], [4]], { grouping: 'stacked', overlap: 100 })).items
    const [a] = rects(items, PALETTE[0])
    const [b] = rects(items, PALETTE[1])
    expect(b.y + b.h).toBeCloseTo(a.y, 5)
    expect(b.x).toBeCloseTo(a.x, 5)
  })

  it('100% 누적: 항목마다 막대 높이의 합이 그래프 높이와 같다', () => {
    const items = scene(chartOf([[1, 5], [3, 5]], { grouping: 'percentStacked', overlap: 100 })).items
    const first = [...rects(items, PALETTE[0]), ...rects(items, PALETTE[1])].filter((r) => r.x < 200)
    expect(first).toHaveLength(2)
    const total = first.reduce((sum, r) => sum + r.h, 0)
    // 그래프 높이: 영역 284 - 위 여백 4
    expect(total).toBeCloseTo(280, 3)
  })

  it('가로 막대: 너비가 값에 비례하고 첫 항목이 아래에 있다', () => {
    const bars = rects(scene(chartOf([[10, 20]], { barDirection: 'bar' })).items, PALETTE[0])
    expect(bars[1].w / bars[0].w).toBeCloseTo(2, 5)
    expect(bars[0].y).toBeGreaterThan(bars[1].y)
  })

  it('항목 순서 뒤집기(maxMin)면 가로 막대의 첫 항목이 위로 간다', () => {
    const chart = chartOf([[10, 20]], { barDirection: 'bar', categoryAxis: axis({ reversed: true }) })
    const bars = rects(scene(chart).items, PALETTE[0])
    expect(bars[0].y).toBeLessThan(bars[1].y)
  })

  it('음수 값은 기준선 아래로 내려간다', () => {
    const items = scene(chartOf([[10, -5]])).items
    const [pos, neg] = rects(items, PALETTE[0])
    expect(neg.y).toBeCloseTo(pos.y + pos.h, 5) // 음수 막대는 양수 막대의 바닥(=기준선)에서 시작
    expect(neg.h).toBeGreaterThan(0)
  })

  it('빈 점(null)은 막대를 그리지 않는다', () => {
    expect(rects(scene(chartOf([[10, null, 5]])).items, PALETTE[0])).toHaveLength(2)
  })
})

describe('buildChartScene: 선', () => {
  const line = (values: (number | null)[], over: Partial<ChartNode> = {}) => chartOf([values], { family: 'line', ...over })

  it('선은 하나의 경로, 점이 모두 이어진다', () => {
    const p = paths(scene(line([1, 3, 2])).items)
    expect(p).toHaveLength(1)
    expect(p[0].d.match(/L/g)).toHaveLength(2)
  })

  it('빈 점에서 선이 끊긴다(gap) — 한 점짜리 조각은 선이 아니다', () => {
    expect(paths(scene(line([1, null, 3, 4])).items)).toHaveLength(1)
    expect(paths(scene(line([1, 2, null, 3, 4])).items)).toHaveLength(2)
  })

  it('빈 점을 0으로 보는 설정이면 끊기지 않고 0 위치를 지난다', () => {
    expect(paths(scene(line([1, null, 3], { blanksAs: 'zero' })).items)).toHaveLength(1)
  })

  it('부드러운 선은 곡선(C), 표식이 켜져 있으면 점마다 원', () => {
    const chart = line([1, 3, 2])
    chart.series[0].smooth = true
    chart.series[0].marker = true
    const items = scene(chart).items
    expect(paths(items)[0].d).toContain('C')
    expect(items.filter((i) => i.t === 'circle')).toHaveLength(3)
  })

  it('점이 칸 가운데에 놓인다(눈금 위가 아니라). pointsOnTicks면 양 끝에 붙는다', () => {
    const between = paths(scene(line([1, 2])).items)[0].d
    const onTicks = paths(scene(line([1, 2], { pointsOnTicks: true })).items)[0].d
    const xs = (d: string) => [...d.matchAll(/[ML]([\d.-]+) /g)].map((m) => Number(m[1]))
    const [b0, b1] = xs(between)
    const [t0, t1] = xs(onTicks)
    expect(t0).toBeLessThan(b0)
    expect(t1).toBeGreaterThan(b1)
  })
})

describe('buildChartScene: 원형', () => {
  const pie = (values: number[], over: Partial<ChartNode> = {}) => chartOf([values], { family: 'pie', ...over })

  it('값이 있는 조각마다 경로 하나, 0 이하 값은 건너뛴다', () => {
    expect(paths(scene(pie([1, 0, 2, -3])).items)).toHaveLength(2)
  })

  it('첫 조각은 12시 방향에서 시작한다', () => {
    const [first] = paths(scene(pie([1, 1])).items)
    // 영역 중심 (200,150), 반지름 = min(384,284)/2 - 4 = 138 -> 12시 방향 끝점 (200, 12)
    expect(first.d.startsWith('M200 150L200 12A')).toBe(true)
  })

  it('조각 색: 점 색이 우선, 없으면 팔레트 순서', () => {
    const chart = pie([1, 1, 1])
    chart.series[0].pointColors = { 1: '#abcdef' }
    expect(paths(scene(chart).items).map((p) => p.fill)).toEqual([PALETTE[0], '#abcdef', PALETTE[2]])
  })

  it('한 조각이 100%면 원으로 그린다', () => {
    const [only] = paths(scene(pie([5])).items)
    expect(only.d).toContain('a')
  })

  it('도넛은 안쪽 구멍이 있다(안쪽 호가 있다)', () => {
    const [slice] = paths(scene(pie([1, 1], { family: 'doughnut', holeSize: 60 })).items)
    expect(slice.d).toMatch(/A[\d.]+ [\d.]+ 0 0 0/)
  })

  it('퍼센트 레이블은 전체 대비 비율', () => {
    const chart = pie([1, 3])
    chart.series[0].labels = { showValue: false, showPercent: true, showCategory: false, showSeries: false, separator: ', ', formatCode: null, position: 'ctr', sizePx: 12, color: null }
    expect(texts(scene(chart).items).map((t) => t.text)).toEqual(['25%', '75%'])
  })
})

describe('buildChartScene: 제목·범례·축', () => {
  it('제목은 위 가운데에, 그 아래로 그래프 영역이 줄어든다', () => {
    const without = rects(scene(chartOf([[10]])).items, PALETTE[0])[0]
    const withTitle = scene(chartOf([[10]], { title: { text: '제목', sizePx: 20, bold: true, color: null } }))
    const title = texts(withTitle.items).find((t) => t.text === '제목')!
    expect(title.x).toBe(200)
    expect(title.anchor).toBe('middle')
    expect(rects(withTitle.items, PALETTE[0])[0].h).toBeLessThan(without.h)
  })

  it('오른쪽 범례는 시리즈마다 한 줄, 그래프 영역을 오른쪽에서 줄인다', () => {
    const base = chartOf([[1, 2], [3, 4]])
    const legend = chartOf([[1, 2], [3, 4]], { legend: { position: 'r', sizePx: 12, color: null } })
    legend.series[0].name = '첫째'
    legend.series[1].name = '둘째'
    const items = scene(legend).items
    expect(texts(items).map((t) => t.text)).toEqual(['첫째', '둘째'])
    // 범례의 색 표시 상자(10x10)도 시리즈 색 사각형이라, 막대만 집계하려고 키 크기보다 큰 것만 센다
    const maxRight = (list: SceneItem[]) => Math.max(...rects(list, PALETTE[0]).filter((r) => r.h > 11).map((r) => r.x + r.w))
    expect(maxRight(items)).toBeLessThan(maxRight(scene(base).items))
  })

  it('원형 범례는 항목(카테고리)마다 한 줄, 이름 없는 시리즈는 계열N', () => {
    const chart = chartOf([[1, 1]], { family: 'pie', legend: { position: 'b', sizePx: 12, color: null } })
    expect(texts(scene(chart).items).map((t) => t.text)).toEqual(['c0', 'c1'])
    const bars = chartOf([[1]], { legend: { position: 't', sizePx: 12, color: null } })
    expect(texts(scene(bars).items).map((t) => t.text)).toEqual(['계열1'])
  })

  it('값 축 눈금 글자는 숫자 형식을 따르고, 숨긴 축은 글자가 없다', () => {
    const shown = chartOf([[0.25, 0.5]], { valueAxis: axis({ hidden: false, formatCode: '0%' }) })
    const labels = texts(scene(shown).items).map((t) => t.text)
    expect(labels.every((t) => t.endsWith('%'))).toBe(true)
    expect(labels).toContain('0%')
    expect(texts(scene(chartOf([[0.25, 0.5]])).items)).toHaveLength(0)
  })

  it('항목 글자가 한 칸에 안 들어가면 건너뛰며 그린다', () => {
    const many = Array.from({ length: 40 }, (_, i) => i)
    const chart = chartOf([many], { categoryAxis: axis({ hidden: false }) })
    chart.categories = many.map((i) => `항목번호${i}`)
    const cats = texts(scene(chart, 300, 200).items).filter((t) => t.text.startsWith('항목번호'))
    expect(cats.length).toBeLessThan(40)
    expect(cats.length).toBeGreaterThan(0)
  })

  it('값 축 제목은 왼쪽에 세로로 눕혀 그린다', () => {
    const chart = chartOf([[1, 2]], { valueAxis: axis({ hidden: false, title: { text: '건수', sizePx: 12, bold: false, color: null } }) })
    const title = texts(scene(chart).items).find((t) => t.text === '건수')!
    expect(title.rotate).toBe(-90)
  })

  it('배경과 테두리', () => {
    const items = scene(chartOf([[1]], { background: '#ffffff', border: { color: '#d9d9d9', widthPx: 1 } })).items
    expect(items[0]).toMatchObject({ t: 'rect', fill: '#ffffff', x: 0, y: 0, w: 400, h: 300 })
    expect(items[1]).toMatchObject({ t: 'rect', fill: 'none', stroke: '#d9d9d9' })
  })
})

describe('buildChartScene: 데이터 레이블', () => {
  const labels = (position: string | null) => ({ showValue: true, showPercent: false, showCategory: false, showSeries: false, separator: ', ', formatCode: null, position, sizePx: 12, color: null })

  it('막대 끝 바깥(outEnd)은 막대 위쪽에 값 글자', () => {
    const chart = chartOf([[10, 20]])
    chart.series[0].labels = labels('outEnd')
    const items = scene(chart).items
    const bars = rects(items, PALETTE[0])
    const t = texts(items)
    expect(t.map((x) => x.text)).toEqual(['10', '20'])
    expect(t[0].y).toBeLessThan(bars[0].y)
    expect(t[0].baseline).toBe('alphabetic')
  })

  it('음수 막대의 outEnd는 막대 아래쪽', () => {
    const chart = chartOf([[10, -5]])
    chart.series[0].labels = labels('outEnd')
    const items = scene(chart).items
    const neg = rects(items, PALETTE[0])[1]
    expect(texts(items)[1].y).toBeGreaterThan(neg.y + neg.h)
  })

  it('가운데(ctr)는 막대 한가운데, 안쪽 글자색은 막대와 대비되는 색', () => {
    const chart = chartOf([[10]], { grouping: 'stacked', overlap: 100 })
    chart.series[0].color = '#000000'
    chart.series[0].labels = labels('ctr')
    const items = scene(chart).items
    const bar = rects(items, '#000000')[0]
    const t = texts(items)[0]
    expect(t.y).toBeCloseTo(bar.y + bar.h / 2, 5)
    expect(t.fill).not.toBe('currentColor')
  })

  it('값 표시 형식은 레이블 설정 > 시리즈 형식 순서', () => {
    const chart = chartOf([[0.5]])
    chart.series[0].formatCode = '0.0%'
    chart.series[0].labels = labels('outEnd')
    expect(texts(scene(chart).items)[0].text).toBe('50.0%')
    chart.series[0].labels = { ...labels('outEnd'), formatCode: '0%' }
    expect(texts(scene(chart).items)[0].text).toBe('50%')
  })

  it('선 차트 레이블은 기본이 점 위쪽', () => {
    const chart = chartOf([[1, 2]], { family: 'line' })
    chart.series[0].labels = labels(null)
    const items = scene(chart).items
    const p = paths(items)[0]
    const ys = [...p.d.matchAll(/[ML][\d.-]+ ([\d.-]+)/g)].map((m) => Number(m[1]))
    const t = texts(items)
    expect(t[0].y).toBeLessThan(ys[0])
  })
})

describe('buildChartScene: 안정성', () => {
  const finite = (items: SceneItem[]) =>
    items.every((i) => Object.values(i).every((v) => typeof v !== 'number' || Number.isFinite(v)))

  it('모든 종류·방향에서 좌표가 유한하다(NaN/Infinity가 없다)', () => {
    const shown = axis({ hidden: false, gridlines: true })
    const charts = [
      chartOf([[1, 2, 3]], { categoryAxis: shown, valueAxis: shown }),
      chartOf([[1, 2, 3]], { barDirection: 'bar', categoryAxis: shown, valueAxis: shown }),
      chartOf([[1, 2, 3], [3, 2, 1]], { grouping: 'stacked', overlap: 100, categoryAxis: shown, valueAxis: shown }),
      chartOf([[0, 0, 0]], { categoryAxis: shown, valueAxis: shown }),
      chartOf([[5]], { family: 'line', categoryAxis: shown, valueAxis: shown, pointsOnTicks: true }),
      chartOf([[null, null]], { categoryAxis: shown, valueAxis: shown }),
      chartOf([[1, 2]], { family: 'doughnut' }),
      chartOf([[0, 0]], { family: 'pie' }),
    ]
    for (const chart of charts) expect(finite(scene(chart).items)).toBe(true)
  })

  it('상자가 아주 작아도 던지지 않는다', () => {
    expect(() => scene(chartOf([[1, 2]], { legend: { position: 'r', sizePx: 12, color: null }, title: { text: '제목', sizePx: 20, bold: false, color: null } }), 30, 20)).not.toThrow()
  })
})
