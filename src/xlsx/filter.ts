/**
 * 자동 필터(헤더 행의 ▼ 버튼) 정보를 읽는다.
 *
 * 역할 분담: 필터 **범위**(`A3:N18`)는 ExcelJS 모델에서 오고, ExcelJS가 버리는 **조건**
 * (`<filterColumn>` — 어느 열에 조건이 걸려 있는지, 버튼이 숨겨졌는지)은 시트 원본 XML에서
 * 이 모듈이 읽는다. 조건으로 숨겨진 행 자체는 파일에 `hidden="1"`로 이미 저장돼 있어서
 * 별도 처리 없이 숨김 행으로 반영된다 — 여기서는 "어느 열에 필터가 걸렸는가"라는 표시용
 * 정보만 더한다.
 *
 * 큰 시트 XML 전체를 DOM으로 파싱하면 느리고 메모리를 많이 쓰므로, `<autoFilter>` 요소
 * 하나만 정규식으로 잘라 속성만 읽는다(시트 XML에서 이 요소는 항상 작다).
 */
import { cellAddress, parseCellAddress } from './cellRef'
import type { ZipArchive } from './zipReader'

export interface SheetFilter {
  /** 필터 버튼이 달린 헤더 행(1-based). 데이터 행은 headerRow+1 ~ lastRow. */
  headerRow: number
  /** 필터 범위의 마지막 행(1-based). */
  lastRow: number
  /** 필터 범위의 첫/마지막 열(1-based, 양끝 포함). */
  firstCol: number
  lastCol: number
  /** 버튼을 숨긴 열(1-based). */
  hiddenButtonCols: number[]
  /** 조건이 걸려 실제로 필터링 중인 열(1-based). */
  activeCols: number[]
}

export interface AutoFilterDetails {
  /** `ref`를 기준으로 한 0-based 열 번호(colId) */
  activeColIds: number[]
  hiddenButtonColIds: number[]
}

const AUTOFILTER_RE = /<(?:\w+:)?autoFilter\b[^>]*?(?:\/>|>[\s\S]*?<\/(?:\w+:)?autoFilter>)/
const FILTER_COLUMN_RE = /<(?:\w+:)?filterColumn\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?filterColumn>)/g
/** 필터 "조건"을 뜻하는 자식 요소들(이 중 하나라도 있으면 그 열은 실제로 필터링 중). */
const CRITERIA_RE = /<(?:\w+:)?(?:filters|customFilters|top10|colorFilter|dynamicFilter|iconFilter)\b/

function attr(attrs: string, name: string): string | null {
  const m = attrs.match(new RegExp(String.raw`\b${name}="([^"]*)"`))
  return m ? m[1] : null
}

/** `<autoFilter>` 요소 문자열에서 조건 있는 열/버튼 숨긴 열을 뽑는다. */
export function parseAutoFilterDetails(autoFilterXml: string): AutoFilterDetails {
  const activeColIds: number[] = []
  const hiddenButtonColIds: number[] = []
  for (const m of autoFilterXml.matchAll(FILTER_COLUMN_RE)) {
    const colIdText = attr(m[1], 'colId')
    if (colIdText === null) continue
    const colId = Number(colIdText)
    if (!Number.isInteger(colId)) continue
    const hidden = attr(m[1], 'hiddenButton')
    const show = attr(m[1], 'showButton')
    if (hidden === '1' || hidden === 'true' || show === '0' || show === 'false') hiddenButtonColIds.push(colId)
    if (m[2] && CRITERIA_RE.test(m[2])) activeColIds.push(colId)
  }
  return { activeColIds, hiddenButtonColIds }
}

/** 시트 원본 XML에서 자동 필터의 조건 정보를 읽는다. 자동 필터가 없으면 null. */
export async function readAutoFilterDetails(zip: ZipArchive, sheetPath: string): Promise<AutoFilterDetails | null> {
  const xml = await zip.readText(sheetPath)
  const match = xml?.match(AUTOFILTER_RE)
  return match ? parseAutoFilterDetails(match[0]) : null
}

/** ExcelJS의 worksheet.autoFilter(문자열 범위 또는 {from,to})를 "A3:N18" 문자열로. */
export function autoFilterRangeText(
  autoFilter: string | { from: string | { row: number; column: number }; to: string | { row: number; column: number } } | null | undefined,
): string | null {
  if (!autoFilter) return null
  if (typeof autoFilter === 'string') return autoFilter
  const point = (p: string | { row: number; column: number }) => (typeof p === 'string' ? p : cellAddress(p.row, p.column))
  return `${point(autoFilter.from)}:${point(autoFilter.to)}`
}

/** 범위 문자열과 조건 정보를 합쳐 화면용 SheetFilter로. 범위가 해석 안 되면 null. */
export function buildSheetFilter(rangeText: string, details: AutoFilterDetails | null): SheetFilter | null {
  const [a, b] = rangeText.split(':')
  try {
    const from = parseCellAddress(a)
    const to = parseCellAddress(b ?? a)
    const firstCol = Math.min(from.col, to.col)
    return {
      headerRow: Math.min(from.row, to.row),
      lastRow: Math.max(from.row, to.row),
      firstCol,
      lastCol: Math.max(from.col, to.col),
      hiddenButtonCols: (details?.hiddenButtonColIds ?? []).map((id) => firstCol + id),
      activeCols: (details?.activeColIds ?? []).map((id) => firstCol + id),
    }
  } catch {
    return null
  }
}
