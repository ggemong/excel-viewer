import { describe, expect, it } from 'vitest'
import { presetPath, rectPath } from './shapeGeometry'

describe('presetPath', () => {
  it('rect', () => {
    expect(presetPath('rect', 100, 50, {})).toBe('M0 0H100V50H0Z')
    expect(rectPath(10, 20)).toBe('M0 0H10V20H0Z')
  })

  it('rightArrow 기본값: 몸통 높이 절반, 머리 길이는 짧은 변의 절반', () => {
    // w=100 h=50 -> ss=50, 머리 25, 몸통 폭 25 (y 12.5~37.5)
    expect(presetPath('rightArrow', 100, 50, {})).toBe('M0 12.5L75 12.5L75 0L100 25L75 50L75 37.5L0 37.5Z')
  })

  it('downArrow 기본값', () => {
    // w=50 h=100 -> ss=50, 머리 25, 몸통 폭 25 (x 12.5~37.5)
    expect(presetPath('downArrow', 50, 100, {})).toBe('M12.5 0L37.5 0L37.5 75L50 75L25 100L0 75L12.5 75Z')
  })

  it('화살표 머리 길이 조절값은 상한이 있다(몸통이 음수가 되지 않게)', () => {
    const d = presetPath('rightArrow', 100, 50, { adj2: 9_999_999 })!
    expect(d.startsWith('M0 12.5L0 12.5')).toBe(true) // 머리가 전체 길이를 차지 -> x1=0
  })

  it('사각 말풍선: 꼬리가 아래로 나가면 아래 변에 꼬리 끝(xPos,yPos)이 찍힌다', () => {
    // 100x100, adj1=-20833 adj2=62500 -> tip=(29.17, 112.5)
    const d = presetPath('wedgeRectCallout', 100, 100, {})!
    expect(d).toContain('L29.17 112.5')
    expect(d.startsWith('M0 0L16.67 0')).toBe(true) // 위쪽 꼬리는 퇴화(일반 변)
  })

  it('사각 말풍선: 꼬리가 왼쪽으로 크게 나가면 왼쪽 변에 붙는다', () => {
    // w=80,h=40: dxPos=-115.72(가로로 훨씬 멀다) -> 꼬리 끝 x=40-115.72=-75.72, y=20+4=24
    const d = presetPath('wedgeRoundRectCallout', 80, 40, { adj1: -144644, adj2: 10000, adj3: 16667 })!
    expect(d).toContain('L-75.72 24')
    // 꼬리 밑변은 왼쪽 변 위(x=0)의 두 점 사이
    expect(d).toContain('L0 33.33L-75.72 24L0 23.33')
  })

  it('타원 말풍선은 꼬리 끝 좌표를 포함한 닫힌 경로', () => {
    const d = presetPath('wedgeEllipseCallout', 100, 100, {})!
    expect(d).toContain('L29.17 112.5Z')
    expect(d).toContain('A50 50 0 1 1')
  })

  it('지원하지 않는 프리셋은 null (호출부가 사각형으로 근사)', () => {
    expect(presetPath('star5', 10, 10, {})).toBeNull()
  })
})
