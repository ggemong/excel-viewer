import { describe, expect, it } from 'vitest'
import { axisCountFor, axisOffset, buildAxis, placeItem, requiredExtent } from './drawingLayout'
import type { DrawingItem, ShapeNode } from './drawingTypes'

const node: ShapeNode = {
  kind: 'shape',
  preset: 'rect',
  adjust: {},
  fill: null,
  line: null,
  text: null,
  rotation: 0,
  flipH: false,
  flipV: false,
}
const pt = (col: number, row: number, colOffset = 0, rowOffset = 0) => ({ col, row, colOffset, rowOffset })

describe('축 계산', () => {
  const cols = buildAxis([50, 100, 30], [false, true, false], 64) // 숨긴 둘째 열은 폭 0

  it('숨긴 칸은 0으로 세어 누적한다', () => {
    expect(cols.offsets).toEqual([0, 50, 50, 80])
  })
  it('데이터 범위를 넘으면 기본 크기로 이어 붙인다', () => {
    expect(axisOffset(cols, 3)).toBe(80)
    expect(axisOffset(cols, 5)).toBe(80 + 2 * 64)
  })
  it('필요한 칸 수: 범위 안은 이진 탐색, 밖은 기본 크기로 올림', () => {
    expect(axisCountFor(cols, 0)).toBe(0)
    expect(axisCountFor(cols, 50)).toBe(1)
    expect(axisCountFor(cols, 51)).toBe(3)
    expect(axisCountFor(cols, 80 + 65)).toBe(3 + 2)
  })
})

describe('개체 배치', () => {
  const cols = buildAxis([50, 100, 30], [false, false, false], 64)
  const rows = buildAxis([20, 20, 20], [false, false, false], 20)

  it('두 셀 앵커: 시작/끝 셀의 오프셋 합으로 상자를 만든다', () => {
    const item: DrawingItem = { anchor: { kind: 'twoCell', from: pt(1, 0, 5, 2), to: pt(3, 2, 0, 10) }, node }
    expect(placeItem(item, cols, rows)).toEqual({ left: 55, top: 2, width: 180 - 55, height: 50 - 2 })
  })
  it('한 셀 앵커는 크기가 고정, 절대 앵커는 좌표 그대로', () => {
    expect(placeItem({ anchor: { kind: 'oneCell', from: pt(1, 1), width: 10, height: 12 }, node }, cols, rows)).toEqual({ left: 50, top: 20, width: 10, height: 12 })
    expect(placeItem({ anchor: { kind: 'absolute', x: 7, y: 8, width: 9, height: 10 }, node }, cols, rows)).toEqual({ left: 7, top: 8, width: 9, height: 10 })
  })
  it('시트 영역 확장: 데이터 0행/0열 시트에 그림만 있어도 필요한 칸 수를 구한다', () => {
    const empty = { cols: buildAxis([], [], 64), rows: buildAxis([], [], 20) }
    const item: DrawingItem = { anchor: { kind: 'twoCell', from: pt(0, 0), to: pt(7, 37) }, node }
    expect(requiredExtent([item], empty.cols, empty.rows)).toEqual({ cols: 7, rows: 37 })
  })
  it('개체가 없으면 0', () => {
    expect(requiredExtent([], cols, rows)).toEqual({ cols: 0, rows: 0 })
  })
})
