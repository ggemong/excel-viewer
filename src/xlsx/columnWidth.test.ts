import { describe, expect, it } from 'vitest'
import { excelColumnWidthToPx, excelPointsToPx, measureDefaultFontWidth } from './columnWidth'

describe('excelPointsToPx', () => {
  it('포인트 -> px는 96/72 고정 비율(폰트와 무관)', () => {
    expect(excelPointsToPx(72)).toBe(96)
    expect(excelPointsToPx(15)).toBe(20)
  })
})

describe('excelColumnWidthToPx', () => {
  it('ECMA-376 공식대로 변환한다(직접 계산해서 검증한 값 — Excel 기본 열너비 8.43, MDW=7)', () => {
    expect(excelColumnWidthToPx(8.43, 7)).toBe(59)
  })

  it('MDW가 커지면(폭 넓은 폰트) 같은 charWidth도 더 넓은 px가 된다', () => {
    expect(excelColumnWidthToPx(10, 10)).toBeGreaterThan(excelColumnWidthToPx(10, 7))
  })
})

describe('measureDefaultFontWidth', () => {
  it('캔버스를 못 쓰는 환경(jsdom)에서는 크래시 없이 근사값으로 대신한다', () => {
    expect(measureDefaultFontWidth()).toBeGreaterThan(0)
  })
})
