/**
 * 차트 값 축의 범위와 눈금 계산 — Excel의 "자동" 축과 같은 규칙을 따른다.
 *
 * 순수 계산만 한다(DOM·글꼴 무관). 이 눈금으로 막대 높이와 선 위치가 정해지므로, 틀리면 차트가 파일과 다른 값을 가리키는 것처럼
 * 보인다 — 그래서 별도 모듈로 떼어 테스트한다.
 */
import type { ChartNode } from '../xlsx/chartTypes'

/** Excel 자동 축: 최솟값이 최댓값의 5/6 이하일 때만 0에서 시작한다(값이 몰려 있으면 0을 포함하지 않는다). */
const ZERO_BASE_RATIO = 5 / 6
/** Excel 자동 축은 데이터 범위 위·아래에 5% 여유를 둔다. */
const HEADROOM_RATIO = 0.05
/**
 * 눈금 간격은 1·2·5·10 × 10^n 중 "가장 가까운 것"으로 고른다(Heckbert의 nice number 반올림). 올림만 하면 간격이 지나치게 커져
 * 눈금이 성겨 보이고(예: 최댓값 100이 0·50·100·150), Excel이 고르는 20 간격과 어긋난다.
 */
const NICE_THRESHOLDS: [limit: number, nice: number][] = [[1.5, 1], [3, 2], [7, 5]]
/** 부동소수점 오차(0.30000000000000004)를 눈금 값에서 없애는 유효 자릿수. */
const TICK_PRECISION = 12

export interface ValueScale {
  lo: number
  hi: number
  step: number
  ticks: number[]
}

export interface ScaleOptions {
  /** 사용자가 정한 값(없으면 null = 자동). */
  min: number | null
  max: number | null
  majorUnit: number | null
  /** 눈금 개수의 목표(그릴 수 있는 길이에서 정해진다). 상한이 아니라 간격을 고르는 기준이라 한두 개 넘칠 수 있다. */
  maxTicks: number
}

function niceStep(raw: number): number {
  const exponent = Math.floor(Math.log10(raw))
  const fraction = raw / 10 ** exponent
  const nice = NICE_THRESHOLDS.find(([limit]) => fraction < limit)?.[1] ?? 10
  return nice * 10 ** exponent
}

const cleanTick = (v: number) => Number(v.toPrecision(TICK_PRECISION))

/**
 * @param dataMin 데이터의 최솟값
 * @param dataMax 데이터의 최댓값
 * @returns 축 범위(lo~hi), 눈금 간격, 눈금 값 목록. 사용자가 정한 값은 그대로 지킨다.
 */
export function niceScale(dataMin: number, dataMax: number, opts: ScaleOptions): ValueScale {
  let lo = dataMin
  let hi = dataMax

  // Excel 자동 범위: 양수 데이터는 최솟값이 최댓값의 5/6 이하면 0부터, 음수 데이터는 거울상.
  if (opts.min === null && lo >= 0 && lo <= hi * ZERO_BASE_RATIO) lo = 0
  if (opts.max === null && hi <= 0 && hi >= lo * ZERO_BASE_RATIO) hi = 0
  if (opts.min !== null) lo = opts.min
  if (opts.max !== null) hi = opts.max
  if (hi <= lo) hi = lo + (Math.abs(lo) || 1)

  const range = hi - lo
  const headroom = range * HEADROOM_RATIO
  if (opts.max === null) hi += headroom
  if (opts.min === null && lo !== 0) lo -= headroom

  const step = opts.majorUnit && opts.majorUnit > 0 ? opts.majorUnit : niceStep((hi - lo) / Math.max(1, opts.maxTicks - 1))
  if (opts.min === null) lo = cleanTick(Math.floor(lo / step + 1e-9) * step)
  if (opts.max === null) hi = cleanTick(Math.ceil(hi / step - 1e-9) * step)
  if (hi <= lo) hi = lo + step

  const count = Math.round((hi - lo) / step)
  const ticks: number[] = []
  for (let i = 0; i <= count; i++) ticks.push(cleanTick(lo + i * step))
  return { lo, hi, step, ticks }
}

/**
 * 차트가 값 축에 그릴 데이터 범위. 묶음/선은 모든 값, 누적은 항목별 합(양수와 음수를 따로), 100% 누적은 0~1 고정.
 * 값이 하나도 없으면 0~1.
 */
export function dataRange(chart: ChartNode): { min: number; max: number } {
  if (chart.family === 'bar' && chart.grouping === 'percentStacked') return { min: 0, max: 1 }

  let min = Infinity
  let max = -Infinity
  const stacked = chart.family === 'bar' && chart.grouping === 'stacked'
  for (let i = 0; i < chart.categories.length; i++) {
    let positive = 0
    let negative = 0
    for (const series of chart.series) {
      const v = series.values[i]
      if (v === null || v === undefined) continue
      if (stacked) {
        if (v >= 0) positive += v
        else negative += v
      } else {
        min = Math.min(min, v)
        max = Math.max(max, v)
      }
    }
    if (stacked) {
      min = Math.min(min, negative, 0)
      max = Math.max(max, positive, 0)
    }
  }
  return Number.isFinite(min) && Number.isFinite(max) ? { min, max } : { min: 0, max: 1 }
}
