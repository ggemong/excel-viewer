/**
 * 차트 모델(ChartNode)과 상자 크기로 "무엇을 어디에 그릴지"를 계산한다 — 제목·범례·축·막대·선·원형 조각·레이블이 전부
 * 좌표가 정해진 도형(SceneItem) 목록이 된다. DOM도 글꼴도 모르는 순수 함수라서(글자 폭은 measure로 주입) 브라우저 없이
 * 막대 높이와 눈금 위치를 숫자로 검증할 수 있다. 실제 SVG 요소로 바꾸는 일은 ChartView.tsx 몫이다.
 *
 * 값 축의 범위·눈금은 chartScale.ts(Excel 자동 축 규칙), 경로 기하는 chartGeometry.ts에 있다.
 */
import { contrastColor } from '../xlsx/cellStyle'
import type { ChartAxis, ChartLabels, ChartNode, ChartSeries, ChartText } from '../xlsx/chartTypes'
import { formatNumber } from '../xlsx/numberFormat'
import { polar, polylinePath, slicePath, smoothPath, type Point } from './chartGeometry'
import { dataRange, niceScale } from './chartScale'

export type SceneItem =
  | { t: 'rect'; x: number; y: number; w: number; h: number; fill: string; stroke?: string; strokeW?: number }
  | { t: 'line'; x1: number; y1: number; x2: number; y2: number; stroke: string; strokeW: number }
  | { t: 'path'; d: string; fill: string | null; stroke: string | null; strokeW: number; evenOdd?: boolean }
  | { t: 'circle'; cx: number; cy: number; r: number; fill: string }
  | { t: 'text'; x: number; y: number; text: string; size: number; anchor: 'start' | 'middle' | 'end'; baseline: 'alphabetic' | 'middle' | 'hanging'; bold: boolean; fill: string; rotate?: number }

export interface ChartScene {
  width: number
  height: number
  items: SceneItem[]
}

/** 글 한 줄의 폭(px)을 재는 함수. 실제로는 캔버스, 테스트에서는 가짜. */
export type MeasureText = (text: string, sizePx: number, bold: boolean) => number

interface Rect {
  x: number
  y: number
  w: number
  h: number
}

const PADDING = 8
const TITLE_GAP = 6
const LINE_HEIGHT = 1.3
const LEGEND_KEY_PX = 10
const LEGEND_LINE_KEY_PX = 18
const LEGEND_KEY_TEXT_GAP = 6
const LEGEND_ITEM_GAP = 14
const LEGEND_PLOT_GAP = 12
const LEGEND_MAX_TEXT_RATIO = 0.35
const AXIS_LABEL_GAP = 6
const AXIS_TITLE_GAP = 4
const LABEL_OFFSET = 4
const NEUTRAL_LINE = 'rgba(128, 128, 128, 0.45)'
const GRID_LINE = 'rgba(128, 128, 128, 0.3)'
const DEFAULT_LINE_WIDTH_PX = 3
const MARKER_RADIUS_PX = 3.5
const MIN_BAR_PX = 1
const TICK_SPACING_VERTICAL_PX = 36
const TICK_SPACING_HORIZONTAL_PX = 80
const MIN_TICKS = 3
const MAX_TICKS = 11
const PIE_LABEL_MARGIN_RATIO = 2.4
const PIE_OUTSIDE_LABEL_GAP = 10
const PIE_INSIDE_LABEL_RATIO = 0.62
/** 글자색을 따로 안 정하면 차트를 담은 상자의 글자색을 따른다(상자가 배경에 맞춰 정한다 — ChartView). */
const TEXT_FILL = 'currentColor'
const ELLIPSIS = '…'

const unnamedSeries = (i: number) => `계열${i + 1}`
const clamp = (v: number, lo: number, hi: number) => Math.min(Math.max(v, lo), hi)
const lineCount = (t: ChartText) => t.text.split('\n').length
const textBlock = (t: ChartText | null) => (t ? t.sizePx * LINE_HEIGHT * lineCount(t) + AXIS_TITLE_GAP : 0)

interface Ctx {
  chart: ChartNode
  items: SceneItem[]
  measure: MeasureText
}

function push(ctx: Ctx, item: SceneItem): void {
  ctx.items.push(item)
}

function drawText(ctx: Ctx, text: ChartText, x: number, y: number, anchor: 'start' | 'middle' | 'end', rotate?: number): void {
  const lines = text.text.split('\n')
  lines.forEach((line, i) => {
    push(ctx, { t: 'text', x, y: y + i * text.sizePx * LINE_HEIGHT, text: line, size: text.sizePx, anchor, baseline: 'hanging', bold: text.bold, fill: text.color ?? TEXT_FILL, rotate })
  })
}

/** maxWidth에 들어가게 뒤를 잘라 …를 붙인다. */
function fitText(ctx: Ctx, text: string, size: number, maxWidth: number): string {
  if (ctx.measure(text, size, false) <= maxWidth) return text
  let end = text.length
  while (end > 1 && ctx.measure(text.slice(0, end) + ELLIPSIS, size, false) > maxWidth) end--
  return text.slice(0, end) + ELLIPSIS
}

// ---------- 범례 ----------

interface LegendEntry {
  label: string
  color: string
  kind: 'box' | 'line'
  marker: boolean
}

function legendEntries(chart: ChartNode): LegendEntry[] {
  const palette = chart.palette
  if (chart.family === 'pie' || chart.family === 'doughnut') {
    const colors = chart.series[0]?.pointColors ?? {}
    return chart.categories.map((label, i) => ({ label, color: colors[i] ?? palette[i % palette.length], kind: 'box' as const, marker: false }))
  }
  return chart.series.map((s, i) => ({ label: s.name ?? unnamedSeries(i), color: s.color ?? palette[i % palette.length], kind: chart.family === 'line' ? ('line' as const) : ('box' as const), marker: s.marker }))
}

/** 범례를 그리고, 남은 본문 영역을 돌려준다. */
function layoutLegend(ctx: Ctx, area: Rect): Rect {
  const { legend } = ctx.chart
  if (!legend) return area
  const entries = legendEntries(ctx.chart)
  if (entries.length === 0) return area

  const size = legend.sizePx
  const rowH = size * 1.5
  const vertical = legend.position === 'r' || legend.position === 'l'
  const maxText = vertical ? area.w * LEGEND_MAX_TEXT_RATIO : area.w
  const keyW = (e: LegendEntry) => (e.kind === 'line' ? LEGEND_LINE_KEY_PX : LEGEND_KEY_PX)
  const labelOf = (e: LegendEntry) => fitText(ctx, e.label, size, maxText - keyW(e) - LEGEND_KEY_TEXT_GAP)
  const widthOf = (e: LegendEntry) => keyW(e) + LEGEND_KEY_TEXT_GAP + ctx.measure(labelOf(e), size, false)

  // 항목 줄 나누기: 세로 범례는 한 항목이 한 줄, 가로 범례는 폭이 차면 줄을 바꾼다.
  const rows: LegendEntry[][] = []
  if (vertical) entries.forEach((e) => rows.push([e]))
  else {
    let row: LegendEntry[] = []
    let rowW = 0
    for (const e of entries) {
      const w = widthOf(e)
      if (row.length > 0 && rowW + LEGEND_ITEM_GAP + w > area.w) {
        rows.push(row)
        row = []
        rowW = 0
      }
      rowW += (row.length > 0 ? LEGEND_ITEM_GAP : 0) + w
      row.push(e)
    }
    if (row.length > 0) rows.push(row)
  }
  const rowWidth = (row: LegendEntry[]) => row.reduce((sum, e, i) => sum + (i > 0 ? LEGEND_ITEM_GAP : 0) + widthOf(e), 0)
  const boxW = Math.max(...rows.map(rowWidth))
  const boxH = rows.length * rowH

  const originX = legend.position === 'r' ? area.x + area.w - boxW : area.x
  const originY = vertical ? area.y + (area.h - boxH) / 2 : legend.position === 't' ? area.y : area.y + area.h - boxH
  rows.forEach((row, r) => {
    let x = vertical ? originX : area.x + (area.w - rowWidth(row)) / 2
    const cy = originY + r * rowH + rowH / 2
    for (const e of row) {
      const kw = keyW(e)
      if (e.kind === 'line') {
        push(ctx, { t: 'line', x1: x, y1: cy, x2: x + kw, y2: cy, stroke: e.color, strokeW: 2 })
        if (e.marker) push(ctx, { t: 'circle', cx: x + kw / 2, cy, r: MARKER_RADIUS_PX - 0.5, fill: e.color })
      } else {
        push(ctx, { t: 'rect', x, y: cy - kw / 2, w: kw, h: kw, fill: e.color })
      }
      push(ctx, { t: 'text', x: x + kw + LEGEND_KEY_TEXT_GAP, y: cy, text: labelOf(e), size, anchor: 'start', baseline: 'middle', bold: false, fill: legend.color ?? TEXT_FILL })
      x += widthOf(e) + LEGEND_ITEM_GAP
    }
  })

  switch (legend.position) {
    case 'r':
      return { ...area, w: area.w - boxW - LEGEND_PLOT_GAP }
    case 'l':
      return { ...area, x: area.x + boxW + LEGEND_PLOT_GAP, w: area.w - boxW - LEGEND_PLOT_GAP }
    case 't':
      return { ...area, y: area.y + boxH + LEGEND_PLOT_GAP, h: area.h - boxH - LEGEND_PLOT_GAP }
    default:
      return { ...area, h: area.h - boxH - LEGEND_PLOT_GAP }
  }
}

// ---------- 데이터 레이블 ----------

function labelText(chart: ChartNode, series: ChartSeries, labels: ChartLabels, index: number, value: number, percent: number | null): string {
  const parts: string[] = []
  if (labels.showSeries) parts.push(series.name ?? '')
  if (labels.showCategory) parts.push(chart.categories[index] ?? '')
  if (labels.showValue) parts.push(formatNumber(value, labels.formatCode ?? series.formatCode))
  if (labels.showPercent && percent !== null) parts.push(formatNumber(percent, labels.showValue ? '0%' : (labels.formatCode ?? '0%')))
  return parts.filter((p) => p !== '').join(labels.separator)
}

/** 막대 위의 글자색: 막대 안에 놓이면 막대 색과 대비되는 색, 바깥이면 기본 글자색. */
function insideColor(labels: ChartLabels, fill: string): string {
  return labels.color ?? (/^#[0-9a-fA-F]{6}$/.test(fill) ? contrastColor(fill) : TEXT_FILL)
}

// ---------- 막대·선 차트 ----------

function drawAxisChart(ctx: Ctx, area: Rect): void {
  const { chart, measure } = ctx
  const horizontal = chart.family === 'bar' && chart.barDirection === 'bar'
  const catAxis = chart.categoryAxis
  const valAxis = chart.valueAxis
  const n = chart.categories.length
  if (n === 0) return

  // 값 축 범위·눈금
  const range = dataRange(chart)
  const approxLength = (horizontal ? area.w : area.h) * 0.8
  const spacing = horizontal ? TICK_SPACING_HORIZONTAL_PX : TICK_SPACING_VERTICAL_PX
  const maxTicks = clamp(Math.floor(approxLength / spacing) + 1, MIN_TICKS, MAX_TICKS)
  const percentStacked = chart.family === 'bar' && chart.grouping === 'percentStacked'
  // 100% 누적은 축이 0~100%로 고정이다(자동 축의 5% 여유가 붙으면 120%까지 늘어난다).
  const scale = niceScale(range.min, range.max, {
    min: valAxis.min ?? (percentStacked ? 0 : null),
    max: valAxis.max ?? (percentStacked ? 1 : null),
    majorUnit: valAxis.majorUnit,
    maxTicks,
  })
  const valueFormat = valAxis.formatCode ?? (percentStacked ? '0%' : (chart.series[0]?.formatCode ?? null))
  const tickLabels = scale.ticks.map((t) => formatNumber(t, valueFormat))

  // 축 글자가 차지하는 여백 → 그래프 영역(frame)
  const maxCatW = catAxis.hidden ? 0 : Math.max(...chart.categories.map((c) => measure(c, catAxis.sizePx, false)))
  const maxTickW = valAxis.hidden ? 0 : Math.max(...tickLabels.map((t) => measure(t, valAxis.sizePx, false)))
  const labelBand = (axis: ChartAxis) => (axis.hidden ? 0 : axis.sizePx * LINE_HEIGHT + AXIS_LABEL_GAP)
  const leftMargin = horizontal ? (catAxis.hidden ? 0 : maxCatW + AXIS_LABEL_GAP) + textBlock(catAxis.title) : (valAxis.hidden ? 0 : maxTickW + AXIS_LABEL_GAP) + textBlock(valAxis.title)
  const bottomMargin = horizontal ? labelBand(valAxis) + textBlock(valAxis.title) : labelBand(catAxis) + textBlock(catAxis.title)
  const topPad = valAxis.hidden ? 4 : valAxis.sizePx / 2 + 2
  const rightPad = Math.min(maxCatW / 2, 30) + 6
  const frame: Rect = { x: area.x + leftMargin, y: area.y + topPad, w: area.w - leftMargin - rightPad, h: area.h - bottomMargin - topPad }
  if (frame.w < 4 || frame.h < 4) return

  // 값 → 좌표 / 항목 → 좌표
  const span = scale.hi - scale.lo || 1
  const valuePos = (v: number) => {
    const t = (clamp(v, scale.lo, scale.hi) - scale.lo) / span
    return horizontal ? frame.x + t * frame.w : frame.y + frame.h - t * frame.h
  }
  const basePos = valuePos(clamp(0, scale.lo, scale.hi))
  const lengthAlong = horizontal ? frame.h : frame.w
  const onTicks = chart.family === 'line' && chart.pointsOnTicks && n > 1
  const band = lengthAlong / n
  /** 축 시작(가로면 아래, 세로면 왼쪽)에서 몇 번째 자리인가. 항목 순서를 뒤집으면 반대. */
  const slotOf = (i: number) => (catAxis.reversed ? n - 1 - i : i)
  const catCenter = (i: number) => {
    const s = slotOf(i)
    const offset = onTicks ? s * (lengthAlong / (n - 1)) : (s + 0.5) * band
    return horizontal ? frame.y + frame.h - offset : frame.x + offset
  }

  // 가로줄(눈금선)
  if (valAxis.gridlines && !valAxis.hidden) {
    for (const t of scale.ticks) {
      const p = valuePos(t)
      push(ctx, horizontal ? { t: 'line', x1: p, y1: frame.y, x2: p, y2: frame.y + frame.h, stroke: valAxis.gridColor ?? GRID_LINE, strokeW: 1 } : { t: 'line', x1: frame.x, y1: p, x2: frame.x + frame.w, y2: p, stroke: valAxis.gridColor ?? GRID_LINE, strokeW: 1 })
    }
  }

  if (chart.family === 'bar') drawBars(ctx, { frame, horizontal, band, valuePos, n, slotOf })
  else drawLines(ctx, { frame, horizontal, valuePos, catCenter, n })

  // 축 선(기준선 / 값 축 선)
  if (!catAxis.hidden) {
    const stroke = catAxis.lineColor ?? NEUTRAL_LINE
    push(ctx, horizontal ? { t: 'line', x1: basePos, y1: frame.y, x2: basePos, y2: frame.y + frame.h, stroke, strokeW: 1 } : { t: 'line', x1: frame.x, y1: basePos, x2: frame.x + frame.w, y2: basePos, stroke, strokeW: 1 })
  }
  if (!valAxis.hidden && valAxis.lineColor) {
    push(ctx, horizontal ? { t: 'line', x1: frame.x, y1: frame.y + frame.h, x2: frame.x + frame.w, y2: frame.y + frame.h, stroke: valAxis.lineColor, strokeW: 1 } : { t: 'line', x1: frame.x, y1: frame.y, x2: frame.x, y2: frame.y + frame.h, stroke: valAxis.lineColor, strokeW: 1 })
  }

  // 값 축 눈금 글자
  if (!valAxis.hidden) {
    scale.ticks.forEach((t, i) => {
      const p = valuePos(t)
      const fill = valAxis.color ?? TEXT_FILL
      push(ctx, horizontal ? { t: 'text', x: p, y: frame.y + frame.h + AXIS_LABEL_GAP, text: tickLabels[i], size: valAxis.sizePx, anchor: 'middle', baseline: 'hanging', bold: false, fill } : { t: 'text', x: frame.x - AXIS_LABEL_GAP, y: p, text: tickLabels[i], size: valAxis.sizePx, anchor: 'end', baseline: 'middle', bold: false, fill })
    })
  }

  // 항목 축 눈금 글자: 안 들어가면 건너뛰며 그린다
  if (!catAxis.hidden) {
    const labelExtent = horizontal ? catAxis.sizePx * LINE_HEIGHT : maxCatW + 6
    const every = Math.max(1, Math.ceil(labelExtent / (onTicks ? lengthAlong / Math.max(1, n - 1) : band)))
    const fill = catAxis.color ?? TEXT_FILL
    for (let i = 0; i < n; i++) {
      if (i % every !== 0) continue
      const c = catCenter(i)
      push(ctx, horizontal ? { t: 'text', x: frame.x - AXIS_LABEL_GAP, y: c, text: chart.categories[i], size: catAxis.sizePx, anchor: 'end', baseline: 'middle', bold: false, fill } : { t: 'text', x: c, y: frame.y + frame.h + AXIS_LABEL_GAP, text: chart.categories[i], size: catAxis.sizePx, anchor: 'middle', baseline: 'hanging', bold: false, fill })
    }
  }

  // 축 제목
  if (valAxis.title) {
    if (horizontal) drawText(ctx, valAxis.title, frame.x + frame.w / 2, area.y + area.h - valAxis.title.sizePx * LINE_HEIGHT * lineCount(valAxis.title), 'middle')
    else drawText(ctx, valAxis.title, area.x + valAxis.title.sizePx * 0.7, frame.y + frame.h / 2, 'middle', -90)
  }
  if (catAxis.title) {
    if (horizontal) drawText(ctx, catAxis.title, area.x + catAxis.title.sizePx * 0.7, frame.y + frame.h / 2, 'middle', -90)
    else drawText(ctx, catAxis.title, frame.x + frame.w / 2, area.y + area.h - catAxis.title.sizePx * LINE_HEIGHT * lineCount(catAxis.title), 'middle')
  }
}

interface BarGeometry {
  frame: Rect
  horizontal: boolean
  band: number
  valuePos: (v: number) => number
  n: number
  slotOf: (i: number) => number
}

function drawBars(ctx: Ctx, g: BarGeometry): void {
  const { chart } = ctx
  const { frame, horizontal, band, valuePos, n, slotOf } = g
  const clustered = chart.grouping === 'clustered'
  const percent = chart.grouping === 'percentStacked'
  const lanes = clustered ? chart.series.length : 1
  const overlap = clustered ? clamp(chart.overlap, -100, 100) / 100 : 1
  const gap = chart.gapWidth / 100
  const barW = band / (lanes - (lanes - 1) * overlap + gap)
  const clusterW = barW * (lanes - (lanes - 1) * overlap)
  const lead = (band - clusterW) / 2

  for (let i = 0; i < n; i++) {
    const s = slotOf(i)
    // 이 항목 띠의 시작선: 세로 막대는 왼쪽, 가로 막대는 아래쪽(첫 항목이 아래)
    const bandStart = horizontal ? frame.y + frame.h - s * band : frame.x + s * band
    const total = percent ? chart.series.reduce((sum, ser) => sum + Math.abs(ser.values[i] ?? 0), 0) : 0
    let positive = 0
    let negative = 0

    chart.series.forEach((series, j) => {
      const raw = series.values[i]
      if (raw === null || raw === undefined) return
      const v = percent ? (total === 0 ? 0 : raw / total) : raw
      let from = 0
      if (!clustered) {
        from = v >= 0 ? positive : negative
        if (v >= 0) positive += v
        else negative += v
      }
      const to = from + v
      const p0 = valuePos(clustered ? 0 : from)
      const p1 = valuePos(to)
      const lane = lead + (clustered ? j * barW * (1 - overlap) : 0)
      const color = series.pointColors[i] ?? series.color ?? chart.palette[j % chart.palette.length]

      const along = Math.max(Math.abs(p1 - p0), MIN_BAR_PX)
      const start = Math.min(p0, p1) - (Math.abs(p1 - p0) < MIN_BAR_PX ? (MIN_BAR_PX - Math.abs(p1 - p0)) / 2 : 0)
      const box: Rect = horizontal ? { x: start, y: bandStart - lane - barW, w: along, h: barW } : { x: bandStart + lane, y: start, w: barW, h: along }
      push(ctx, { t: 'rect', x: box.x, y: box.y, w: box.w, h: box.h, fill: color })

      if (series.labels) drawBarLabel(ctx, series, series.labels, i, raw, box, horizontal, color, clustered)
    })
  }
}

/**
 * 막대의 데이터 레이블 위치. 막대의 "끝"은 기준선에서 먼 쪽(양수: 세로는 위, 가로는 오른쪽 / 음수: 반대)이고 `dir`은
 * 끝에서 바깥으로 향하는 화면 방향이다. outEnd는 끝 밖, inEnd는 끝 안, inBase는 기준선 안, ctr은 한가운데.
 */
function drawBarLabel(ctx: Ctx, series: ChartSeries, labels: ChartLabels, index: number, value: number, box: Rect, horizontal: boolean, fill: string, clustered: boolean): void {
  const text = labelText(ctx.chart, series, labels, index, value, null)
  if (!text) return
  const positive = value >= 0
  const dir = horizontal ? (positive ? 1 : -1) : positive ? -1 : 1
  const end = horizontal ? (positive ? box.x + box.w : box.x) : positive ? box.y : box.y + box.h
  const base = horizontal ? (positive ? box.x : box.x + box.w) : positive ? box.y + box.h : box.y
  const position = labels.position ?? (clustered ? 'outEnd' : 'ctr')
  const color = position === 'outEnd' ? (labels.color ?? TEXT_FILL) : insideColor(labels, fill)
  const cx = box.x + box.w / 2
  const cy = box.y + box.h / 2
  const make = (x: number, y: number, anchor: 'start' | 'middle' | 'end', baseline: 'alphabetic' | 'middle' | 'hanging'): SceneItem => ({
    t: 'text', x, y, text, size: labels.sizePx, anchor, baseline, bold: false, fill: color,
  })

  if (position === 'ctr') return void push(ctx, make(cx, cy, 'middle', 'middle'))
  // 글이 놓일 점과, 그 점에서 글이 뻗는 방향(ext)
  const at = position === 'inEnd' ? end - dir * LABEL_OFFSET : position === 'inBase' ? base + dir * LABEL_OFFSET : end + dir * LABEL_OFFSET
  const ext = position === 'inEnd' ? -dir : dir
  if (horizontal) push(ctx, make(at, cy, ext > 0 ? 'start' : 'end', 'middle'))
  else push(ctx, make(cx, at, 'middle', ext < 0 ? 'alphabetic' : 'hanging'))
}

interface LineGeometry {
  frame: Rect
  horizontal: boolean
  valuePos: (v: number) => number
  catCenter: (i: number) => number
  n: number
}

function drawLines(ctx: Ctx, g: LineGeometry): void {
  const { chart } = ctx
  const { valuePos, catCenter, n } = g
  chart.series.forEach((series, j) => {
    const color = series.color ?? chart.palette[j % chart.palette.length]
    const points: (Point | null)[] = []
    for (let i = 0; i < n; i++) {
      const v = series.values[i] ?? (chart.blanksAs === 'zero' ? 0 : null)
      points.push(v === null ? null : { x: catCenter(i), y: valuePos(v) })
    }
    // 빈 점에서 선을 끊는다: 연속된 점끼리 한 조각
    const segments: Point[][] = []
    let current: Point[] = []
    for (const p of points) {
      if (p) current.push(p)
      else if (current.length > 0) {
        segments.push(current)
        current = []
      }
    }
    if (current.length > 0) segments.push(current)
    for (const seg of segments) {
      if (seg.length < 2) continue
      push(ctx, { t: 'path', d: series.smooth ? smoothPath(seg) : polylinePath(seg), fill: null, stroke: color, strokeW: series.lineWidthPx ?? DEFAULT_LINE_WIDTH_PX })
    }
    if (series.marker) {
      for (const p of points) if (p) push(ctx, { t: 'circle', cx: p.x, cy: p.y, r: MARKER_RADIUS_PX, fill: color })
    }
    const labels = series.labels
    if (labels) {
      points.forEach((p, i) => {
        const v = series.values[i]
        if (!p || v === null || v === undefined) return
        const text = labelText(chart, series, labels, i, v, null)
        if (!text) return
        const position = labels.position ?? 't'
        const base = { t: 'text' as const, text, size: labels.sizePx, bold: false, fill: labels.color ?? TEXT_FILL }
        const o = MARKER_RADIUS_PX + LABEL_OFFSET
        if (position === 'b') push(ctx, { ...base, x: p.x, y: p.y + o, anchor: 'middle', baseline: 'hanging' })
        else if (position === 'l') push(ctx, { ...base, x: p.x - o, y: p.y, anchor: 'end', baseline: 'middle' })
        else if (position === 'r') push(ctx, { ...base, x: p.x + o, y: p.y, anchor: 'start', baseline: 'middle' })
        else if (position === 'ctr') push(ctx, { ...base, x: p.x, y: p.y, anchor: 'middle', baseline: 'middle' })
        else push(ctx, { ...base, x: p.x, y: p.y - o, anchor: 'middle', baseline: 'alphabetic' })
      })
    }
  })
}

// ---------- 원형 / 도넛 ----------

function drawPie(ctx: Ctx, area: Rect): void {
  const { chart } = ctx
  const series = chart.series[0]
  if (!series) return
  const values = series.values.map((v) => (v !== null && v > 0 ? v : 0))
  const total = values.reduce((a, b) => a + b, 0)
  if (total <= 0) return

  const labels = series.labels
  const outerLimit = Math.min(area.w, area.h) / 2
  const outer = Math.max(1, outerLimit - (labels ? labels.sizePx * PIE_LABEL_MARGIN_RATIO : 4))
  const inner = chart.family === 'doughnut' ? outer * (chart.holeSize / 100) : 0
  const cx = area.x + area.w / 2
  const cy = area.y + area.h / 2
  const separator = chart.background ?? '#FFFFFF'

  let angle = (chart.firstSliceAngle * Math.PI) / 180
  values.forEach((v, i) => {
    if (v <= 0) return
    const sweep = (v / total) * Math.PI * 2
    const color = series.pointColors[i] ?? chart.palette[i % chart.palette.length]
    const { d, evenOdd } = slicePath(cx, cy, outer, inner, angle, angle + sweep)
    push(ctx, { t: 'path', d, fill: color, stroke: separator, strokeW: 1, evenOdd })

    if (labels) {
      const text = labelText(chart, series, labels, i, v, v / total)
      if (text) {
        const mid = angle + sweep / 2
        const inside = chart.family === 'doughnut' || labels.position === 'ctr' || labels.position === 'inEnd'
        const radius = chart.family === 'doughnut' ? (outer + inner) / 2 : labels.position === 'inEnd' ? outer * 0.8 : inside ? outer * PIE_INSIDE_LABEL_RATIO : outer + PIE_OUTSIDE_LABEL_GAP
        const p = polar(cx, cy, radius, mid)
        const rightSide = Math.sin(mid) >= 0
        push(ctx, {
          t: 'text', x: p.x, y: p.y, text, size: labels.sizePx, bold: false,
          anchor: inside ? 'middle' : rightSide ? 'start' : 'end', baseline: 'middle',
          fill: inside ? insideColor(labels, color) : (labels.color ?? TEXT_FILL),
        })
      }
    }
    angle += sweep
  })
}

// ---------- 공개 ----------

/**
 * @param chart 읽어 둔 차트 모델
 * @param width 차트 상자의 폭(px)
 * @param height 차트 상자의 높이(px)
 * @param measure 글자 폭을 재는 함수(제목·범례·축 글자가 차지할 자리를 계산하는 데 쓴다)
 */
export function buildChartScene(chart: ChartNode, width: number, height: number, measure: MeasureText): ChartScene {
  const ctx: Ctx = { chart, items: [], measure }
  if (chart.background) push(ctx, { t: 'rect', x: 0, y: 0, w: width, h: height, fill: chart.background })
  if (chart.border) {
    const bw = chart.border.widthPx
    push(ctx, { t: 'rect', x: bw / 2, y: bw / 2, w: width - bw, h: height - bw, fill: 'none', stroke: chart.border.color, strokeW: bw })
  }

  let area: Rect = { x: PADDING, y: PADDING, w: width - 2 * PADDING, h: height - 2 * PADDING }
  if (chart.title) {
    drawText(ctx, chart.title, width / 2, area.y, 'middle')
    const used = chart.title.sizePx * LINE_HEIGHT * lineCount(chart.title) + TITLE_GAP
    area = { ...area, y: area.y + used, h: area.h - used }
  }
  area = layoutLegend(ctx, area)
  if (area.w < 8 || area.h < 8) return { width, height, items: ctx.items }

  if (chart.family === 'pie' || chart.family === 'doughnut') drawPie(ctx, area)
  else drawAxisChart(ctx, area)
  return { width, height, items: ctx.items }
}
