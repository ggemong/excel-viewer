/**
 * 한 줄 글이 칸보다 길 때, 옆 칸이 비어 있으면 그 위로 글을 이어서 보여준다(Excel의 "넘침").
 *
 * 규칙(Excel과 같다):
 *  - 글(문자열)만 넘친다. 숫자는 넘치지 않는다. 줄바꿈 칸과 병합 칸도 넘치지 않는다.
 *  - 글이 가는 방향은 가로 정렬을 따른다: 기본/왼쪽 정렬은 오른쪽으로, 오른쪽 정렬은 왼쪽으로, 가운데 정렬은
 *    양쪽으로 같은 만큼.
 *  - 값이 들어 있는 칸(값이 빈 문자열이 아닌 칸), 병합에 걸린 칸, 필터 버튼이 달린 칸 앞에서 멈추고 그 칸을
 *    가리지 않는다. 배경색만 있는 빈 칸은 막지 않는다.
 *  - 숨긴 열은 건너뛴다(화면에 없으므로).
 *
 * 이 파일은 "얼마나 넘칠 수 있는가"만 계산한다. 실제로 넘쳐 보이게 하는 건 Grid.tsx의 몫이다.
 */
import type { CellStyle } from '../xlsx/cellStyle'
import type { SheetModel } from '../xlsx/types'

export interface Spill {
  /** 글이 차지할 수 있는 폭의 상한(px) = 칸 안쪽 폭 + 넘칠 수 있는 폭. 이 폭에서 글을 자른다. */
  maxWidth: number
}

export interface SpillInput {
  rows: SheetModel['rows']
  colWidths: number[]
  /** 화면에 그려지는 열 번호(숨긴 열 제외)를 왼쪽부터. */
  visibleCols: number[]
  /** "행,열" -> 병합 범위. 있으면 그 칸은 넘치지도, 넘침을 받지도 않는다. */
  merged: ReadonlyMap<string, unknown>
  /** 글이 넘어올 수 없는 칸인가(예: 필터 버튼이 달린 칸). */
  isBlocked: (row: number, col: number) => boolean
  /** 칸의 가로 여백+테두리 합(px) — 칸 폭에서 이만큼을 빼야 글이 들어갈 안쪽 폭이다. */
  cellChrome: number
  /** 글 한 줄의 폭(px). 잴 수 없으면 null(그러면 넘침을 계산하지 않는다 — 눈대중 폭으로 판단하지 않는다). */
  measure: (text: string, style: CellStyle | null) => number | null
}

/** 캔버스 폭과 실제 렌더링 폭의 소수점 차이로, 딱 맞는 글이 넘치는 것으로 판정되는 걸 막는 여유(px). */
const FIT_TOLERANCE_PX = 0.5
/** 같은 이유로 글 상자 폭에 더하는 여유 — 안 더하면 캔버스가 조금 좁게 잰 글의 마지막 글자가 잘린다. 옆 칸의 오른쪽 여백(10px)보다 훨씬 작다. */
const MAX_WIDTH_SLACK_PX = 1

function isEmptyValue(value: unknown): boolean {
  return value === null || value === undefined || value === ''
}

/**
 * @returns (행, 열) -> 넘침 정보. 넘칠 필요가 없거나 넘칠 수 없으면 null. 결과는 칸별로 캐시한다(같은 칸을 렌더링마다
 * 다시 재지 않도록) — 입력(시트, 글꼴 등)이 바뀌면 새로 만든다.
 */
export function createSpillResolver(input: SpillInput): (row: number, col: number) => Spill | null {
  const { rows, colWidths, visibleCols, merged, isBlocked, cellChrome, measure } = input
  const visibleIndex = new Map(visibleCols.map((col, i) => [col, i]))
  const cache = new Map<string, Spill | null>()

  /** 방향(+1 오른쪽/-1 왼쪽)으로 막히기 전까지 이어진 빈 칸들의 폭 합. needed를 채우면 더 보지 않는다. */
  const freeWidth = (row: number, from: number, step: 1 | -1, needed: number): number => {
    let total = 0
    for (let i = from + step; i >= 0 && i < visibleCols.length && total < needed; i += step) {
      const col = visibleCols[i]
      if (merged.has(`${row},${col}`) || isBlocked(row, col) || !isEmptyValue(rows[row - 1]?.[col - 1]?.value)) break
      total += colWidths[col - 1]
    }
    return total
  }

  const compute = (row: number, col: number): Spill | null => {
    const cell = rows[row - 1]?.[col - 1]
    const index = visibleIndex.get(col)
    if (!cell || index === undefined || typeof cell.value !== 'string' || cell.value === '') return null
    if (cell.style?.align?.wrap || cell.hyperlink || merged.has(`${row},${col}`) || isBlocked(row, col)) return null

    const width = measure(cell.value, cell.style)
    if (width === null) return null
    const inner = colWidths[col - 1] - cellChrome
    const overflow = width - inner
    if (overflow <= FIT_TOLERANCE_PX) return null

    // 글이 칸 밖으로 쓸 수 있는 폭(room). 가운데 정렬은 양쪽 모두 반만큼씩 비어 있어야 넘친다 — 한쪽이 막혀 있으면
    // 글이 한쪽으로 쏠려 가운데 정렬이 깨지므로 넘치지 않고 칸 안에서 잘린다.
    const align = cell.style?.align?.h ?? 'left'
    let room: number
    if (align === 'right') room = Math.min(overflow, freeWidth(row, index, -1, overflow))
    else if (align === 'center') room = 2 * Math.min(overflow / 2, freeWidth(row, index, -1, overflow / 2), freeWidth(row, index, 1, overflow / 2))
    else room = Math.min(overflow, freeWidth(row, index, 1, overflow))
    return room > 0 ? { maxWidth: inner + room + MAX_WIDTH_SLACK_PX } : null
  }

  return (row, col) => {
    const key = `${row},${col}`
    if (cache.has(key)) return cache.get(key) ?? null
    const spill = compute(row, col)
    cache.set(key, spill)
    return spill
  }
}
