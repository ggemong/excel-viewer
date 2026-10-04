/**
 * 차트 XML(`xl/charts/chartN.xml`)을 화면용 ChartNode로 바꾼다.
 *
 * 경계: XML -> 모델 변환까지만 한다. 어디에 놓일지(앵커)는 drawing.ts, 어떻게 그릴지는 src/grid/chartScene.ts.
 *
 * 값은 파일이 저장해 둔 캐시를 그대로 쓴다. Excel은 저장할 때 차트가 가리키는 셀 값을 캐시로 함께 적어 두므로 셀을
 * 다시 읽어 계산할 필요가 없다. 지원하지 않는 차트(영역·분산형·콤보·3D 등)나 캐시가 없는 차트는 억지로 근사해서 틀리게 그리지
 * 않고 { ok: false, reason }으로 돌려줘 호출부가 "표시하지 못해요"로 안내하게 한다.
 */
import { resolveDrawingColor } from './drawingColor'
import type {
  BarGrouping,
  ChartAxis,
  ChartFamily,
  ChartLabels,
  ChartLegend,
  ChartParseResult,
  ChartSeries,
  ChartSkipReason,
  ChartText,
  LegendPosition,
} from './chartTypes'
import { formatNumber } from './numberFormat'
import { resolveThemeColorByName, type ThemeColors } from './themeColor'
import { parseXml } from './workbookParts'
import { kid, kids } from './xmlDom'

const EMU_PER_PX = 9525
const PX_PER_PT = 96 / 72
const FONT_SIZE_UNIT = 100
/** Excel 차트의 기본 글자 크기(pt) — 파일이 크기를 안 적었을 때. */
const DEFAULT_TITLE_PT = 14
const DEFAULT_TEXT_PT = 9
const DEFAULT_GAP_WIDTH = 150
const DEFAULT_HOLE_SIZE = 50
const DEFAULT_LABEL_SEPARATOR = ', '
/** 차트 영역 채우기가 파일에 없으면 Excel은 흰 배경으로 그린다. */
const DEFAULT_BACKGROUND = '#FFFFFF'

/** 그룹 요소 이름 -> 우리가 그리는 종류. 여기 없는 `*Chart`는 지원하지 않는 종류다. */
const SUPPORTED_GROUPS: Record<string, ChartFamily> = { lineChart: 'line', barChart: 'bar', pieChart: 'pie', doughnutChart: 'doughnut' }
/** 지원하지 않는 그룹 요소의 안내용 분류. */
const UNSUPPORTED_REASONS: Record<string, ChartSkipReason> = {
  areaChart: 'chart:area',
  scatterChart: 'chart:scatter',
  radarChart: 'chart:radar',
  bubbleChart: 'chart:bubble',
  stockChart: 'chart:stock',
  surfaceChart: 'chart:surface',
}

/** 테마를 모르거나 강조색이 비었을 때 쓰는 Office 기본 강조색(accent1~6). */
const FALLBACK_PALETTE = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47']
const ACCENT_NAMES = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6']

const ptToPx = (pt: number) => pt * PX_PER_PT

/** 파일의 테마 강조색 1~6. 하나라도 비어 있으면 일부만 섞이지 않게 기본색 전체로 대신한다. */
function paletteOf(theme: ThemeColors | null): string[] {
  const colors = theme ? ACCENT_NAMES.map((name) => resolveThemeColorByName(theme, name)) : []
  return colors.length === ACCENT_NAMES.length && colors.every((c): c is string => Boolean(c)) ? colors : FALLBACK_PALETTE
}

function val(el: Element | null | undefined, name: string): string | null {
  return kid(el, name)?.getAttribute('val') ?? null
}

function numVal(el: Element | null | undefined, name: string): number | null {
  const raw = val(el, name)
  if (raw === null) return null
  const n = Number(raw)
  return Number.isFinite(n) ? n : null
}

/** `<c:smooth/>`처럼 val이 없으면 true인 불리언 요소. 요소 자체가 없으면 기본값. */
function flag(el: Element | null | undefined, name: string, fallback: boolean): boolean {
  const child = kid(el, name)
  if (!child) return fallback
  const v = child.getAttribute('val')
  return v === null ? true : v === '1' || v === 'true'
}

// ---------- 글자 ----------

interface TextProps {
  sizePx: number
  bold: boolean
  color: string | null
}

/** rPr 또는 defRPr 요소에서 크기·굵기·색. 없으면 기본값. */
function textPropsOf(props: Element | null, theme: ThemeColors | null, defaultPt: number): TextProps {
  const sz = Number(props?.getAttribute('sz'))
  const b = props?.getAttribute('b')
  return {
    sizePx: ptToPx(Number.isFinite(sz) && sz > 0 ? sz / FONT_SIZE_UNIT : defaultPt),
    bold: b === '1' || b === 'true',
    color: resolveDrawingColor(kid(props, 'solidFill'), theme),
  }
}

/** `txPr`(글자 서식만 있는 상자)의 기본 글자 속성. */
function txPrProps(txPr: Element | null, theme: ThemeColors | null, defaultPt: number): TextProps {
  return textPropsOf(kid(kid(kid(txPr, 'p'), 'pPr'), 'defRPr'), theme, defaultPt)
}

function parseTitle(titleEl: Element | null, theme: ThemeColors | null, defaultPt: number): ChartText | null {
  if (!titleEl) return null
  const tx = kid(titleEl, 'tx')
  const rich = kid(tx, 'rich')
  if (rich) {
    const paragraphs = kids(rich, 'p')
    const text = paragraphs
      .map((p) => kids(p, 'r').map((r) => kid(r, 't')?.textContent ?? '').join(''))
      .join('\n')
      .trim()
    if (!text) return null
    const firstParagraph = paragraphs[0]
    const props = textPropsOf(kid(kids(firstParagraph, 'r')[0], 'rPr') ?? kid(kid(firstParagraph, 'pPr'), 'defRPr'), theme, defaultPt)
    return { text, ...props }
  }
  // 셀을 가리키는 제목: 캐시된 글자를 쓴다.
  const text = readCache(kid(tx, 'strRef')).points.get(0)?.trim()
  return text ? { text, sizePx: ptToPx(defaultPt), bold: false, color: null } : null
}

// ---------- 캐시(저장된 값) ----------

interface Cache {
  count: number
  points: Map<number, string>
  formatCode: string | null
}

/** strRef/numRef/strLit/numLit/multiLvlStrRef 어느 것이든 점 번호 -> 글자 맵으로 읽는다(여러 단계면 가장 안쪽 단계). */
function readCache(ref: Element | null): Cache {
  const empty: Cache = { count: 0, points: new Map(), formatCode: null }
  if (!ref) return empty
  let cache = kid(ref, 'numCache') ?? kid(ref, 'strCache') ?? kid(ref, 'multiLvlStrCache') ?? ref
  const level = kid(cache, 'lvl')
  if (level) cache = level
  const points = new Map<number, string>()
  let maxIdx = -1
  for (const pt of kids(cache, 'pt')) {
    const idx = Number(pt.getAttribute('idx'))
    if (!Number.isInteger(idx) || idx < 0) continue
    points.set(idx, kid(pt, 'v')?.textContent ?? '')
    maxIdx = Math.max(maxIdx, idx)
  }
  const declared = Number(kid(cache, 'ptCount')?.getAttribute('val') ?? 0)
  return { count: Math.max(Number.isFinite(declared) ? declared : 0, maxIdx + 1), points, formatCode: kid(cache, 'formatCode')?.textContent ?? null }
}

function numericValues(valEl: Element | null): { values: (number | null)[]; formatCode: string | null } {
  const cache = readCache(kid(valEl, 'numRef') ?? kid(valEl, 'numLit'))
  const values: (number | null)[] = []
  for (let i = 0; i < cache.count; i++) {
    const raw = cache.points.get(i)
    const n = raw === undefined || raw === '' ? NaN : Number(raw)
    values.push(Number.isFinite(n) ? n : null)
  }
  return { values, formatCode: cache.formatCode }
}

function categoryLabels(catEl: Element | null): string[] {
  if (!catEl) return []
  const numeric = kid(catEl, 'numRef') ?? kid(catEl, 'numLit')
  const cache = readCache(numeric ?? kid(catEl, 'strRef') ?? kid(catEl, 'strLit') ?? kid(catEl, 'multiLvlStrRef'))
  const labels: string[] = []
  for (let i = 0; i < cache.count; i++) {
    const raw = cache.points.get(i) ?? ''
    // 숫자 카테고리(날짜 등)는 캐시의 표시 형식대로 글자로 만든다.
    labels.push(numeric && raw !== '' && Number.isFinite(Number(raw)) ? formatNumber(Number(raw), cache.formatCode) : raw)
  }
  return labels
}

function seriesName(ser: Element): string | null {
  const tx = kid(ser, 'tx')
  const name = readCache(kid(tx, 'strRef')).points.get(0) ?? kid(tx, 'v')?.textContent ?? null
  return name && name.trim() ? name : null
}

// ---------- 모양 ----------

function solidColor(spPr: Element | null, theme: ThemeColors | null): string | null {
  return kid(spPr, 'noFill') ? null : resolveDrawingColor(kid(spPr, 'solidFill'), theme)
}

function lineOf(spPr: Element | null): Element | null {
  return kid(spPr, 'ln')
}

function lineWidthPx(ln: Element | null): number | null {
  const w = Number(ln?.getAttribute('w'))
  return Number.isFinite(w) && w > 0 ? w / EMU_PER_PX : null
}

function parseLabels(el: Element | null, theme: ThemeColors | null): ChartLabels | null {
  if (!el || flag(el, 'delete', false)) return null
  const showValue = flag(el, 'showVal', false)
  const showPercent = flag(el, 'showPercent', false)
  const showCategory = flag(el, 'showCatName', false)
  const showSeries = flag(el, 'showSerName', false)
  if (!showValue && !showPercent && !showCategory && !showSeries) return null
  const numFmt = kid(el, 'numFmt')
  const props = txPrProps(kid(el, 'txPr'), theme, DEFAULT_TEXT_PT)
  return {
    showValue,
    showPercent,
    showCategory,
    showSeries,
    separator: kid(el, 'separator')?.textContent ?? DEFAULT_LABEL_SEPARATOR,
    formatCode: numFmt && numFmt.getAttribute('sourceLinked') !== '1' ? numFmt.getAttribute('formatCode') : null,
    position: val(el, 'dLblPos'),
    sizePx: props.sizePx,
    color: props.color,
  }
}

function parseAxis(el: Element | null, theme: ThemeColors | null): ChartAxis {
  const scaling = kid(el, 'scaling')
  const numFmt = kid(el, 'numFmt')
  const props = txPrProps(kid(el, 'txPr'), theme, DEFAULT_TEXT_PT)
  const gridLine = lineOf(kid(kid(el, 'majorGridlines'), 'spPr'))
  return {
    // 축 요소가 아예 없는 차트(원형)는 축을 그리지 않는다.
    hidden: el === null || flag(el, 'delete', false),
    title: parseTitle(kid(el, 'title'), theme, DEFAULT_TEXT_PT),
    // sourceLinked=1이면 "값의 원래 형식을 따른다"는 뜻이라 null로 둔다(그리는 쪽이 시리즈 형식을 쓴다).
    formatCode: numFmt && numFmt.getAttribute('sourceLinked') !== '1' ? numFmt.getAttribute('formatCode') : null,
    sizePx: props.sizePx,
    color: props.color,
    lineColor: resolveDrawingColor(kid(lineOf(kid(el, 'spPr')), 'solidFill'), theme),
    gridlines: Boolean(kid(el, 'majorGridlines')),
    gridColor: resolveDrawingColor(kid(gridLine, 'solidFill'), theme),
    min: numVal(scaling, 'min'),
    max: numVal(scaling, 'max'),
    majorUnit: numVal(el, 'majorUnit'),
    reversed: val(scaling, 'orientation') === 'maxMin',
  }
}

function parseLegend(el: Element | null, theme: ThemeColors | null): ChartLegend | null {
  if (!el) return null
  const pos = val(el, 'legendPos')
  const position: LegendPosition = pos === 'l' || pos === 't' || pos === 'b' ? pos : 'r'
  const props = txPrProps(kid(el, 'txPr'), theme, DEFAULT_TEXT_PT)
  return { position, sizePx: props.sizePx, color: props.color }
}

// ---------- 시리즈 ----------

interface ParsedSeries {
  series: ChartSeries
  categories: string[]
}

function parseSeries(ser: Element, family: ChartFamily, groupLabels: ChartLabels | null, groupMarker: boolean, theme: ThemeColors | null): ParsedSeries {
  const { values, formatCode } = numericValues(kid(ser, 'val'))
  const spPr = kid(ser, 'spPr')
  const ln = lineOf(spPr)
  const color = family === 'line' ? (kid(ln, 'noFill') ? null : (resolveDrawingColor(kid(ln, 'solidFill'), theme) ?? solidColor(spPr, theme))) : solidColor(spPr, theme)

  const pointColors: Record<number, string> = {}
  for (const dPt of kids(ser, 'dPt')) {
    const idx = numVal(dPt, 'idx')
    const pointColor = solidColor(kid(dPt, 'spPr'), theme)
    if (idx !== null && pointColor) pointColors[idx] = pointColor
  }

  const markerEl = kid(ser, 'marker')
  const dLbls = kid(ser, 'dLbls')
  return {
    series: {
      name: seriesName(ser),
      values,
      color,
      pointColors,
      lineWidthPx: lineWidthPx(ln),
      marker: markerEl ? val(markerEl, 'symbol') !== 'none' : groupMarker,
      // <c:smooth/>는 val 없이도 true. 선 차트가 아니면 의미 없다.
      smooth: flag(ser, 'smooth', false),
      // 시리즈가 자기 레이블 설정을 가지면(켜져 있든 꺼져 있든) 그게 우선, 없으면 그룹 설정을 따른다.
      labels: dLbls ? parseLabels(dLbls, theme) : groupLabels,
      formatCode,
    },
    categories: categoryLabels(kid(ser, 'cat')),
  }
}

// ---------- 공개 ----------

function skip(reason: ChartSkipReason): ChartParseResult {
  return { ok: false, reason }
}

/**
 * 차트 XML 하나를 ChartNode로. 지원하지 않는 차트나 값이 없는 차트는 사유와 함께 ok:false.
 * XML이 깨졌어도 던지지 않고 ok:false('chart')로 돌려준다 — 차트 하나가 시트 전체를 막으면 안 된다(호출부가 센다).
 */
export function parseChartXml(xml: string, theme: ThemeColors | null): ChartParseResult {
  let space: Element
  try {
    space = parseXml(xml).documentElement
  } catch {
    return skip('chart')
  }

  const chart = kid(space, 'chart')
  const plotArea = kid(chart, 'plotArea')
  if (!chart || !plotArea) return skip('chart')

  const groups = Array.from(plotArea.children).filter((c) => c.localName.endsWith('Chart'))
  if (groups.length === 0) return skip('chart')
  // 한 차트에 여러 종류(콤보)나 보조축이 섞이면 한 종류로 근사하지 않는다.
  if (groups.length > 1) return skip('chart:combo')

  const group = groups[0]
  const family = SUPPORTED_GROUPS[group.localName]
  if (!family) return skip(group.localName.includes('3D') ? 'chart:3d' : (UNSUPPORTED_REASONS[group.localName] ?? 'chart:other'))

  const groupingRaw = val(group, 'grouping') ?? (family === 'bar' ? 'clustered' : 'standard')
  // 선 차트의 누적(stacked/percentStacked)은 그리지 않는다.
  if (family === 'line' && groupingRaw !== 'standard') return skip('chart:other')
  const grouping: BarGrouping = groupingRaw === 'stacked' || groupingRaw === 'percentStacked' ? groupingRaw : 'clustered'

  const groupLabels = parseLabels(kid(group, 'dLbls'), theme)
  const groupMarker = flag(group, 'marker', true)
  const parsed = kids(group, 'ser')
    .sort((a, b) => (numVal(a, 'order') ?? 0) - (numVal(b, 'order') ?? 0))
    .map((ser) => parseSeries(ser, family, groupLabels, groupMarker, theme))
  if (parsed.length === 0) return skip('chart:nodata')

  // 원형은 첫 시리즈만 그린다(Excel도 그렇다).
  const used = family === 'pie' || family === 'doughnut' ? parsed.slice(0, 1) : parsed
  const categoryCount = Math.max(...used.map((p) => Math.max(p.series.values.length, p.categories.length)))
  if (!used.some((p) => p.series.values.some((v) => v !== null))) return skip('chart:nodata')

  const labelSource = used.find((p) => p.categories.length > 0)?.categories ?? []
  const categories = Array.from({ length: categoryCount }, (_, i) => (labelSource[i] !== undefined && labelSource[i] !== '' ? labelSource[i] : String(i + 1)))
  const series = used.map((p) => ({ ...p.series, values: Array.from({ length: categoryCount }, (_, i) => p.series.values[i] ?? null) }))

  const catAxisEl = kid(plotArea, 'catAx') ?? kid(plotArea, 'dateAx')
  const valAxisEl = kid(plotArea, 'valAx')
  const spPr = kid(space, 'spPr')
  const border = lineOf(spPr)
  const borderColor = kid(border, 'noFill') ? null : resolveDrawingColor(kid(border, 'solidFill'), theme)
  // 제목이 파일에 명시돼 있을 때만 그린다(제목을 안 적은 단일 시리즈 차트에 Excel이 붙이는 "자동 제목"은 만들지 않는다).
  const title = parseTitle(kid(chart, 'title'), theme, DEFAULT_TITLE_PT)
  const blanks = val(chart, 'dispBlanksAs')

  return {
    ok: true,
    node: {
      kind: 'chart',
      rotation: 0,
      flipH: false,
      flipV: false,
      family,
      barDirection: val(group, 'barDir') === 'bar' ? 'bar' : 'col',
      grouping,
      title,
      categories,
      series,
      palette: paletteOf(theme),
      legend: parseLegend(kid(chart, 'legend'), theme),
      categoryAxis: parseAxis(catAxisEl, theme),
      valueAxis: parseAxis(valAxisEl, theme),
      gapWidth: numVal(group, 'gapWidth') ?? DEFAULT_GAP_WIDTH,
      overlap: numVal(group, 'overlap') ?? (grouping === 'clustered' ? 0 : 100),
      pointsOnTicks: val(valAxisEl, 'crossBetween') === 'midCat',
      holeSize: numVal(group, 'holeSize') ?? DEFAULT_HOLE_SIZE,
      firstSliceAngle: numVal(group, 'firstSliceAng') ?? 0,
      blanksAs: blanks === 'zero' ? 'zero' : 'gap',
      background: spPr ? (kid(spPr, 'noFill') ? null : (resolveDrawingColor(kid(spPr, 'solidFill'), theme) ?? DEFAULT_BACKGROUND)) : DEFAULT_BACKGROUND,
      border: borderColor ? { color: borderColor, widthPx: lineWidthPx(border) ?? 1 } : null,
    },
  }
}
