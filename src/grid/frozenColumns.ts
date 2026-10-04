/**
 * 열 고정(틀 고정의 가로 쪽): 왼쪽 N개 열이 가로로 스크롤해도 화면 왼쪽에 붙어 있게 한다.
 *
 * 이 앱은 열을 가상화하지 않아 모든 열이 DOM에 있으므로, 고정 열의 칸마다 `position: sticky; left: X`를 주면 된다.
 * 여기서는 그 X(행 번호 칸 폭 + 앞선 고정 열 폭의 합)와 고정 영역의 전체 폭을 계산한다. 스타일을 입히는 일은 Grid.tsx 몫이다.
 */
export interface FrozenColumnLayout {
  /** 고정 열 번호 -> sticky left(px). 행 번호 칸의 폭(rowNumWidth)이 이미 더해져 있다. */
  left: ReadonlyMap<number, number>
  /** 고정 영역의 전체 폭(px) = 행 번호 칸 + 보이는 고정 열들. 이 폭만큼은 가로 스크롤로도 가려지지 않는다. */
  width: number
  /** 화면에 보이는 마지막 고정 열(경계선을 그을 열). 고정 열이 없으면 null. */
  edgeCol: number | null
}

/**
 * @param colWidths 열 번호-1 순서의 폭(px)
 * @param visibleCols 화면에 그려지는 열 번호(숨긴 열 제외), 왼쪽부터
 * @param frozenCols 파일이 고정한 열 개수(1번 열부터 센 개수 — 숨긴 열도 센다. Excel의 xSplit과 같다)
 * @param rowNumWidth 행 번호 칸의 폭(px). 행 번호 칸도 항상 왼쪽에 붙어 있으므로 고정 열은 그 오른쪽에서 시작한다.
 */
export function frozenColumnLayout(colWidths: number[], visibleCols: number[], frozenCols: number, rowNumWidth: number): FrozenColumnLayout {
  const left = new Map<number, number>()
  let offset = rowNumWidth
  let edgeCol: number | null = null
  for (const col of visibleCols) {
    if (col > frozenCols) break
    left.set(col, offset)
    offset += colWidths[col - 1]
    edgeCol = col
  }
  return { left, width: offset, edgeCol }
}
