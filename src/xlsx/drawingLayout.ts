/**
 * 그림·도형의 셀 기준 위치(앵커)를 시트 좌표(px)로 바꾸는 순수 계산.
 *
 * 열너비/행높이가 제각각이고 숨긴 열/행은 폭이 0이라, "N번째 열의 왼쪽 x"는 앞쪽 열 폭의
 * 누적합이다. 그리드(src/grid/Grid.tsx)가 실제로 쌓는 방식과 똑같이 숨긴 칸을 0으로 세어야
 * 그림이 셀과 어긋나지 않는다.
 *
 * 이 계산은 두 곳에서 쓴다 — read.ts(그림이 데이터 영역 밖까지 걸쳐 있으면 시트 크기를
 * 늘리기 위해)와 Grid(그림을 어디에 그릴지). 한쪽에만 두면 두 계산이 어긋나므로 한 모듈에 둔다.
 */
import type { DrawingItem, DrawingNode } from './drawingTypes'

export interface Axis {
  /** offsets[i] = i번째(0-based) 칸의 시작 위치. 길이는 count + 1. */
  offsets: number[]
  count: number
  /** 데이터 범위를 벗어난 칸의 크기(Excel 기본 열폭/행높이). */
  defaultSize: number
}

export interface PlacedBox {
  left: number
  top: number
  width: number
  height: number
}

export function buildAxis(sizes: number[], hidden: boolean[], defaultSize: number): Axis {
  const offsets = [0]
  for (let i = 0; i < sizes.length; i++) {
    offsets.push(offsets[i] + (hidden[i] ? 0 : sizes[i]))
  }
  return { offsets, count: sizes.length, defaultSize }
}

/** index번째 칸의 시작 위치. 데이터 범위 밖이면 기본 크기로 이어 붙여 계산한다. */
export function axisOffset(axis: Axis, index: number): number {
  if (index <= axis.count) return axis.offsets[Math.max(0, index)]
  return axis.offsets[axis.count] + (index - axis.count) * axis.defaultSize
}

/** 시작 위치 `px`를 덮으려면 칸이 최소 몇 개 있어야 하는가. */
export function axisCountFor(axis: Axis, px: number): number {
  const end = axis.offsets[axis.count]
  if (px <= end) {
    // 이진 탐색: offsets[n] >= px 인 가장 작은 n
    let lo = 0
    let hi = axis.count
    while (lo < hi) {
      const mid = (lo + hi) >> 1
      if (axis.offsets[mid] >= px) hi = mid
      else lo = mid + 1
    }
    return lo
  }
  return axis.count + Math.ceil((px - end) / axis.defaultSize)
}

export function placeItem(item: DrawingItem, cols: Axis, rows: Axis): PlacedBox {
  const a = item.anchor
  if (a.kind === 'absolute') return { left: a.x, top: a.y, width: a.width, height: a.height }

  const left = axisOffset(cols, a.from.col) + a.from.colOffset
  const top = axisOffset(rows, a.from.row) + a.from.rowOffset
  if (a.kind === 'oneCell') return { left, top, width: a.width, height: a.height }

  const right = axisOffset(cols, a.to.col) + a.to.colOffset
  const bottom = axisOffset(rows, a.to.row) + a.to.rowOffset
  return { left, top, width: Math.max(0, right - left), height: Math.max(0, bottom - top) }
}

/** 모든 개체가 들어가려면 열/행이 최소 몇 개 필요한가. 개체가 없으면 0. */
export function requiredExtent(items: DrawingItem[], cols: Axis, rows: Axis): { cols: number; rows: number } {
  let needCols = 0
  let needRows = 0
  for (const item of items) {
    const box = placeItem(item, cols, rows)
    needCols = Math.max(needCols, axisCountFor(cols, box.left + box.width))
    needRows = Math.max(needRows, axisCountFor(rows, box.top + box.height))
  }
  return { cols: needCols, rows: needRows }
}

function nodeBlobs(node: DrawingNode, out: Blob[]): void {
  if (node.kind === 'picture' && node.blob) out.push(node.blob)
  if (node.kind === 'group') node.children.forEach((c) => nodeBlobs(c.node, out))
}

/** 화면에 그려야 할 그림 Blob 전부(묶음 안 포함) — object URL을 만들 대상. */
export function collectPictureBlobs(items: DrawingItem[]): Blob[] {
  const out: Blob[] = []
  items.forEach((item) => nodeBlobs(item.node, out))
  return out
}
