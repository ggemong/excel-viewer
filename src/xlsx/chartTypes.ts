/**
 * 시트 위 차트의 화면용 모델. xlsx의 `xl/charts/chartN.xml`(DrawingML Chart)을 src/xlsx/chart.ts가 읽어서 이 모양으로
 * 만든다. 읽기 전용 뷰어라 다시 쓰는 경로는 없다.
 *
 * 지원하는 범위(밖은 그리지 않고 종류별로 세어 안내한다): 선·막대(세로/가로, 묶음/누적/100% 누적)·원형·도넛, 한 종류만.
 * 값은 파일에 저장된 캐시(`numCache`/`strCache`)를 쓴다 — Excel이 저장할 때 항상 함께 적어 두므로 셀을 다시 읽어 계산할 필요가
 * 없고, 캐시가 없는 차트(일부 생성 도구)는 값이 없는 것으로 보고 지원하지 않는 차트로 센다.
 *
 * 좌표·크기 단위는 전부 px다. 그리는 쪽(src/grid/chartScene.ts)은 이 모델과 상자 크기만 있으면 된다.
 */
import type { DrawingTransform } from './drawingTypes'

export type ChartFamily = 'line' | 'bar' | 'pie' | 'doughnut'
export type BarGrouping = 'clustered' | 'stacked' | 'percentStacked'
export type LegendPosition = 'r' | 'l' | 't' | 'b'

export interface ChartText {
  /** 줄바꿈은 `\n`. */
  text: string
  sizePx: number
  bold: boolean
  /** null이면 차트 기본 글자색(차트 영역 배경에 맞춰 정해진다). */
  color: string | null
}

/** 데이터 레이블이 보여줄 항목 — Excel의 "레이블 내용" 체크박스 그대로. */
export interface ChartLabels {
  showValue: boolean
  showPercent: boolean
  showCategory: boolean
  showSeries: boolean
  separator: string
  /** 숫자 표시 형식. null이면 값의 원래 형식(시리즈 캐시)을 쓴다. */
  formatCode: string | null
  /** 레이블 위치(outEnd, inEnd, ctr, t, b, l, r, bestFit ...). null이면 차트 종류의 기본 위치. */
  position: string | null
  sizePx: number
  color: string | null
}

export interface ChartSeries {
  name: string | null
  /** 카테고리 순서의 값. 비어 있는 점은 null. */
  values: (number | null)[]
  /** null이면 테마 색 순서에서 고른다. */
  color: string | null
  /** 원형·막대에서 점(조각)마다 따로 정한 색: 점 번호 -> 색. */
  pointColors: Record<number, string>
  lineWidthPx: number | null
  /** 선 차트에서 점마다 표식을 그릴지. */
  marker: boolean
  smooth: boolean
  labels: ChartLabels | null
  /** 값의 표시 형식 코드(캐시의 formatCode). */
  formatCode: string | null
}

export interface ChartAxis {
  /** 축 자체를 그리지 않는다(`c:delete`). */
  hidden: boolean
  title: ChartText | null
  /** 눈금 글자의 숫자 형식. null이면 값의 원래 형식. */
  formatCode: string | null
  sizePx: number
  color: string | null
  lineColor: string | null
  /** 값 축의 주 눈금 가로줄. */
  gridlines: boolean
  gridColor: string | null
  /** 값 축: 사용자가 정한 범위·눈금 간격. 없으면 null(자동). */
  min: number | null
  max: number | null
  majorUnit: number | null
  /** 항목 순서를 뒤집는다(orientation maxMin). */
  reversed: boolean
}

export interface ChartLegend {
  position: LegendPosition
  sizePx: number
  color: string | null
}

export interface ChartNode extends DrawingTransform {
  kind: 'chart'
  family: ChartFamily
  /** 막대에서만 의미 있다: col = 세로 막대, bar = 가로 막대. */
  barDirection: 'col' | 'bar'
  grouping: BarGrouping
  title: ChartText | null
  categories: string[]
  series: ChartSeries[]
  /** 색을 따로 정하지 않은 시리즈·조각이 순서대로 쓰는 색(테마의 강조색 1~6, 테마를 모르면 Office 기본색). */
  palette: string[]
  legend: ChartLegend | null
  categoryAxis: ChartAxis
  valueAxis: ChartAxis
  /** 막대 사이 간격(막대 폭 대비 %, 기본 150). */
  gapWidth: number
  /** 막대 겹침(%, 묶음 0 / 누적 100). */
  overlap: number
  /** 선 차트에서 점이 칸의 가운데가 아니라 눈금 위에 놓인다(crossBetween midCat). */
  pointsOnTicks: boolean
  /** 도넛 구멍 크기(바깥 반지름 대비 %). */
  holeSize: number
  /** 원형의 첫 조각 시작 각도(도, 12시 방향부터 시계 방향). */
  firstSliceAngle: number
  /** 값이 없는 점: gap = 선을 끊는다, zero = 0으로 본다. */
  blanksAs: 'gap' | 'zero'
  /** null이면 채우기 없음(투명). */
  background: string | null
  border: { color: string; widthPx: number } | null
}

/** 차트를 못 그릴 때의 사유 — 그대로 skipped 키로 쓰여 안내 문구가 된다(src/grid/skippedDrawingNotice.ts). */
export type ChartSkipReason = 'chart' | 'chart:area' | 'chart:scatter' | 'chart:radar' | 'chart:bubble' | 'chart:stock' | 'chart:surface' | 'chart:3d' | 'chart:combo' | 'chart:other' | 'chart:nodata'

export type ChartParseResult = { ok: true; node: ChartNode } | { ok: false; reason: ChartSkipReason }
