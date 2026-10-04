import { describe, expect, it } from 'vitest'
import type { ChartNode } from '../xlsx/chartTypes'
import { dataRange, niceScale } from './chartScale'

const auto = { min: null, max: null, majorUnit: null, maxTicks: 6 }

describe('niceScale', () => {
  it('양수 데이터는 0에서 시작하고 위에 5% 여유를 둔 뒤 1·2·5 간격으로 맞춘다(Excel: 최댓값 100 -> 축 120)', () => {
    const s = niceScale(10, 100, auto)
    expect(s.lo).toBe(0)
    expect(s.hi).toBe(120)
    expect(s.ticks).toEqual([0, 20, 40, 60, 80, 100, 120])
  })

  it('값이 몰려 있으면(최솟값이 최댓값의 5/6보다 크면) 0을 포함하지 않는다', () => {
    const s = niceScale(90, 100, auto)
    expect(s.lo).toBeGreaterThan(0)
    expect(s.lo).toBeLessThanOrEqual(90)
    expect(s.hi).toBeGreaterThanOrEqual(100)
  })

  it('사용자가 정한 범위와 눈금 간격은 그대로 지킨다', () => {
    const s = niceScale(3, 97, { min: 0, max: 100, majorUnit: 25, maxTicks: 6 })
    expect(s.ticks).toEqual([0, 25, 50, 75, 100])
  })

  it('음수가 섞인 데이터는 0을 지나는 범위가 된다', () => {
    const s = niceScale(-50, 30, auto)
    expect(s.lo).toBeLessThanOrEqual(-50)
    expect(s.hi).toBeGreaterThanOrEqual(30)
    expect(s.ticks).toContain(0)
  })

  it('소수 눈금에 부동소수점 오차가 남지 않는다(0.30000000000000004 같은 값)', () => {
    const s = niceScale(0, 1, { min: 0, max: 1, majorUnit: 0.1, maxTicks: 6 })
    expect(s.ticks).toHaveLength(11)
    expect(s.ticks[3]).toBe(0.3)
  })

  it('모든 값이 같아도 폭이 있는 범위를 만든다', () => {
    const s = niceScale(5, 5, auto)
    expect(s.hi).toBeGreaterThan(s.lo)
    expect(s.ticks.length).toBeGreaterThanOrEqual(2)
  })

  it('눈금 개수 상한을 넘기지 않는다', () => {
    const s = niceScale(0, 1000, { ...auto, maxTicks: 4 })
    expect(s.ticks.length).toBeLessThanOrEqual(6)
  })
})

function chartOf(over: Partial<ChartNode>, series: (number | null)[][]): ChartNode {
  const axis = { hidden: false, title: null, formatCode: null, sizePx: 12, color: null, lineColor: null, gridlines: false, gridColor: null, min: null, max: null, majorUnit: null, reversed: false }
  return {
    kind: 'chart', rotation: 0, flipH: false, flipV: false,
    family: 'bar', barDirection: 'col', grouping: 'clustered', title: null,
    categories: series[0].map((_, i) => `c${i}`),
    series: series.map((values) => ({ name: null, values, color: null, pointColors: {}, lineWidthPx: null, marker: false, smooth: false, labels: null, formatCode: null })),
    palette: ['#111111', '#222222'], legend: null, categoryAxis: axis, valueAxis: axis,
    gapWidth: 150, overlap: 0, pointsOnTicks: false, holeSize: 50, firstSliceAngle: 0, blanksAs: 'gap', background: null, border: null,
    ...over,
  }
}

describe('dataRange', () => {
  it('묶음/선은 모든 값의 최소·최대', () => {
    expect(dataRange(chartOf({}, [[3, 9], [1, 5]]))).toEqual({ min: 1, max: 9 })
  })

  it('누적 막대는 항목별 합(양수/음수 따로)과 0을 포함', () => {
    const chart = chartOf({ grouping: 'stacked' }, [[2, -1], [3, -4]])
    expect(dataRange(chart)).toEqual({ min: -5, max: 5 })
  })

  it('100% 누적은 0~1', () => {
    expect(dataRange(chartOf({ grouping: 'percentStacked' }, [[2, 3]]))).toEqual({ min: 0, max: 1 })
  })

  it('값이 하나도 없으면 0~1', () => {
    expect(dataRange(chartOf({}, [[null, null]]))).toEqual({ min: 0, max: 1 })
  })
})
