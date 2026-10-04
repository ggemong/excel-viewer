/**
 * 줄바꿈되는 글이 실제로 몇 px 높이를 차지하는지 브라우저에게 직접 물어서 잰다.
 *
 * 왜 직접 줄 수를 계산하지 않고 DOM으로 재나: 줄바꿈 위치는 브라우저의 줄바꿈 규칙(한글 음절 사이, 공백,
 * 긴 단어의 강제 끊기 등)과 글꼴에 달려 있어서, 캔버스 폭으로 같은 규칙을 흉내 내면 칸마다 한 줄씩 틀릴 수
 * 있다. 실제 칸과 같은 CSS 클래스(.grid-cell)와 같은 인라인 서식(cellStyleProps)으로 그려서 재면 "화면에
 * 그려질 모양"과 어긋날 수 없다.
 *
 * 한 번에 모아서 재는 이유: 요소를 하나 만들 때마다 높이를 읽으면 그때마다 레이아웃을 다시 계산한다.
 * 전부 붙인 뒤에 읽으면 레이아웃이 한 번만 돈다.
 */
import { cellStyleProps, type CellStyle } from '../xlsx/cellStyle'

export interface WrapRequest {
  text: string
  style: CellStyle | null
  /** 칸 전체 폭(px, 안쪽 여백·테두리 포함). 안쪽 폭은 CSS 여백에서 저절로 나온다. */
  cellWidth: number
}

/** 요청 하나당 높이(px) 하나를 같은 순서로 돌려준다. rowFit이 이 모양에만 의존해 테스트에서 가짜로 바꿀 수 있다. */
export type MeasureWrappedHeights = (requests: WrapRequest[]) => number[]

const MEASURE_HOST_CLASS = 'grid-measure-host'
const MEASURE_CELL_CLASS = 'grid-cell grid-cell--measure'

let cachedChrome: number | null = null

/**
 * 칸의 가로 여백+테두리 합(px) — CSS(.grid-cell)에서 직접 읽는다. 칸 폭에서 이만큼을 빼야 글이 들어갈 안쪽 폭이 되는데,
 * 여백을 코드에 숫자로 박으면 디자인 토큰을 바꿀 때 글자 넘침 판단이 조용히 어긋난다.
 * 0이면 스타일이 아직 적용되지 않은 것이므로 저장하지 않고 다음에 다시 읽는다.
 */
export function cellHorizontalChrome(): number {
  if (cachedChrome !== null) return cachedChrome
  if (typeof document === 'undefined') return 0
  const host = document.createElement('div')
  host.className = MEASURE_HOST_CLASS
  const probe = document.createElement('div')
  probe.className = MEASURE_CELL_CLASS
  host.appendChild(probe)
  document.body.appendChild(host)
  try {
    const style = getComputedStyle(probe)
    const sum = [style.paddingLeft, style.paddingRight, style.borderLeftWidth, style.borderRightWidth].reduce((total, v) => total + (Number.parseFloat(v) || 0), 0)
    if (sum > 0) cachedChrome = sum
    return sum
  } finally {
    host.remove()
  }
}

/** @throws 브라우저가 아니면(문서가 없으면) 던진다 — 높이를 0으로 돌려주면 모든 행이 "맞는다"고 오해하게 된다. */
export const measureWrappedHeights: MeasureWrappedHeights = (requests) => {
  if (requests.length === 0) return []
  if (typeof document === 'undefined') throw new Error('줄바꿈 높이는 브라우저에서만 잴 수 있어요')

  const host = document.createElement('div')
  host.className = MEASURE_HOST_CLASS
  const cells = requests.map((request) => {
    const el = document.createElement('div')
    el.className = MEASURE_CELL_CLASS
    Object.assign(el.style, cellStyleProps(request.style, true), { width: `${request.cellWidth}px` })
    el.textContent = request.text
    host.appendChild(el)
    return el
  })
  document.body.appendChild(host)
  try {
    return cells.map((el) => el.getBoundingClientRect().height)
  } finally {
    host.remove()
  }
}
