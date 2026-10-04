import { describe, expect, it } from 'vitest'
import { polar, polylinePath, slicePath, smoothPath } from './chartGeometry'

describe('경로', () => {
  it('꺾은선', () => {
    expect(polylinePath([{ x: 0, y: 0 }, { x: 10, y: 5 }])).toBe('M0 0L10 5')
    expect(polylinePath([])).toBe('')
  })

  it('부드러운 곡선은 처음과 끝 점을 지나고 점이 둘 이하면 꺾은선과 같다', () => {
    const pts = [{ x: 0, y: 0 }, { x: 10, y: 10 }, { x: 20, y: 0 }]
    const d = smoothPath(pts)
    expect(d.startsWith('M0 0C')).toBe(true)
    expect(d.endsWith('20 0')).toBe(true)
    expect(d.match(/C/g)).toHaveLength(2)
    expect(smoothPath(pts.slice(0, 2))).toBe(polylinePath(pts.slice(0, 2)))
  })
})

describe('polar', () => {
  it('각도는 12시 방향에서 시계 방향', () => {
    const top = polar(0, 0, 10, 0)
    expect(top.x).toBeCloseTo(0)
    expect(top.y).toBeCloseTo(-10)
    const right = polar(0, 0, 10, Math.PI / 2)
    expect(right.x).toBeCloseTo(10)
    expect(right.y).toBeCloseTo(0)
  })
})

describe('slicePath', () => {
  it('원형 조각은 중심에서 시작해 호로 이어진다 — 반 바퀴 이하면 large 플래그 0', () => {
    const { d, evenOdd } = slicePath(100, 100, 50, 0, 0, Math.PI / 2)
    expect(d).toBe('M100 100L100 50A50 50 0 0 1 150 100Z')
    expect(evenOdd).toBe(false)
  })

  it('반 바퀴를 넘으면 large 플래그 1', () => {
    expect(slicePath(0, 0, 10, 0, 0, Math.PI * 1.5).d).toContain('A10 10 0 1 1')
  })

  it('도넛 조각은 바깥 호와 안쪽 호(반대 방향)로 닫힌다', () => {
    const { d } = slicePath(0, 0, 10, 5, 0, Math.PI / 2)
    expect(d).toContain('A10 10 0 0 1')
    expect(d).toContain('A5 5 0 0 0')
  })

  it('한 바퀴를 다 차지하면 퇴화한 호 대신 원을 그린다(도넛은 구멍 뚫린 원)', () => {
    const pie = slicePath(0, 0, 10, 0, 0, Math.PI * 2)
    expect(pie.d.match(/a/g)).toHaveLength(2)
    expect(pie.evenOdd).toBe(false)
    const donut = slicePath(0, 0, 10, 5, 0, Math.PI * 2)
    expect(donut.d.match(/a/g)).toHaveLength(4)
    expect(donut.evenOdd).toBe(true)
  })
})
