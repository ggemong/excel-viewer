import { describe, expect, it } from 'vitest'
import { frozenColumnLayout } from './frozenColumns'

const ROW_NUM = 44

describe('frozenColumnLayout', () => {
  it('고정 열의 left는 행 번호 칸 폭 + 앞선 고정 열 폭의 합', () => {
    const layout = frozenColumnLayout([100, 80, 60, 50], [1, 2, 3, 4], 2, ROW_NUM)
    expect(layout.left.get(1)).toBe(44)
    expect(layout.left.get(2)).toBe(144)
    expect(layout.left.has(3)).toBe(false)
    expect(layout.width).toBe(44 + 100 + 80)
    expect(layout.edgeCol).toBe(2)
  })

  it('고정한 열이 없으면 비어 있고 폭은 행 번호 칸뿐', () => {
    const layout = frozenColumnLayout([100], [1], 0, ROW_NUM)
    expect(layout.left.size).toBe(0)
    expect(layout.width).toBe(ROW_NUM)
    expect(layout.edgeCol).toBeNull()
  })

  it('숨긴 열은 폭을 차지하지 않지만 고정 열 개수에는 센다(Excel의 xSplit과 같다)', () => {
    // 2번 열이 숨겨져 있고 3개를 고정 -> 1, 3번 열이 고정된다
    const layout = frozenColumnLayout([100, 80, 60, 50], [1, 3, 4], 3, ROW_NUM)
    expect([...layout.left.keys()]).toEqual([1, 3])
    expect(layout.left.get(3)).toBe(144)
    expect(layout.width).toBe(44 + 100 + 60)
    expect(layout.edgeCol).toBe(3)
  })

  it('고정 열 개수가 열 수보다 많아도 있는 열까지만', () => {
    const layout = frozenColumnLayout([100, 80], [1, 2], 9, ROW_NUM)
    expect(layout.left.size).toBe(2)
    expect(layout.edgeCol).toBe(2)
  })
})
