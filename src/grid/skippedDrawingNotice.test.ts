import { describe, expect, it } from 'vitest'
import { skippedDrawingNotice } from './skippedDrawingNotice'

describe('skippedDrawingNotice', () => {
  it('없으면 null', () => {
    expect(skippedDrawingNotice({})).toBeNull()
    expect(skippedDrawingNotice({ chart: 0 })).toBeNull()
  })
  it('종류별 개수를 한글로', () => {
    expect(skippedDrawingNotice({ chart: 2, graphicFrame: 1 })).toBe('이 시트에는 아직 표시하지 못하는 개체가 있어요: 차트 2개, 표·슬라이서 등 1개')
  })
  it('모르는 종류도 숨기지 않고 이름 그대로', () => {
    expect(skippedDrawingNotice({ weird: 3 })).toContain('weird 3개')
  })
})
