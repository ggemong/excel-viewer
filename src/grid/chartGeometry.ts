/**
 * 차트를 그릴 때 쓰는 경로 기하 — 부드러운 선(베지어), 원형 조각(호). SVG path의 `d` 문자열만 만든다.
 */

export interface Point {
  x: number
  y: number
}

const fmt = (n: number) => Number(n.toFixed(2))
const pair = (p: Point) => `${fmt(p.x)} ${fmt(p.y)}`

/** 점들을 꺾은선으로 잇는다. 점이 없으면 ''. */
export function polylinePath(points: Point[]): string {
  return points.map((p, i) => `${i === 0 ? 'M' : 'L'}${pair(p)}`).join('')
}

/**
 * 점들을 부드러운 곡선으로 잇는다(Catmull-Rom 스플라인을 3차 베지어로 변환). Excel의 "곡선 선"과 모양이 같다.
 * 점이 둘 이하면 곡선이 될 수 없어 꺾은선과 같다.
 */
export function smoothPath(points: Point[]): string {
  if (points.length <= 2) return polylinePath(points)
  let d = `M${pair(points[0])}`
  for (let i = 0; i < points.length - 1; i++) {
    const p0 = points[i - 1] ?? points[i]
    const p1 = points[i]
    const p2 = points[i + 1]
    const p3 = points[i + 2] ?? p2
    const c1 = { x: p1.x + (p2.x - p0.x) / 6, y: p1.y + (p2.y - p0.y) / 6 }
    const c2 = { x: p2.x - (p3.x - p1.x) / 6, y: p2.y - (p3.y - p1.y) / 6 }
    d += `C${pair(c1)} ${pair(c2)} ${pair(p2)}`
  }
  return d
}

/** 12시 방향에서 시계 방향으로 잰 각도(라디안)의 원 위 점. */
export function polar(cx: number, cy: number, radius: number, angle: number): Point {
  return { x: cx + radius * Math.sin(angle), y: cy - radius * Math.cos(angle) }
}

/** 온전한 한 바퀴로 보는 기준(부동소수점 오차 흡수). */
const FULL_TURN = Math.PI * 2 - 1e-6

/**
 * 원형(안쪽 반지름 0)이나 도넛의 조각 한 개.
 * 각도는 12시 방향 기준 시계 방향 라디안. 한 바퀴를 다 차지하면 호가 퇴화하므로 원(도넛은 구멍 뚫린 원)으로 그린다.
 *
 * @returns path `d`와, 한 바퀴 조각이 구멍을 뚫기 위해 fillRule evenodd가 필요한지
 */
export function slicePath(cx: number, cy: number, outer: number, inner: number, start: number, end: number): { d: string; evenOdd: boolean } {
  if (end - start >= FULL_TURN) {
    const circle = (r: number) => `M${fmt(cx - r)} ${fmt(cy)}a${fmt(r)} ${fmt(r)} 0 1 0 ${fmt(2 * r)} 0a${fmt(r)} ${fmt(r)} 0 1 0 ${fmt(-2 * r)} 0Z`
    return { d: inner > 0 ? circle(outer) + circle(inner) : circle(outer), evenOdd: inner > 0 }
  }
  const large = end - start > Math.PI ? 1 : 0
  const o0 = polar(cx, cy, outer, start)
  const o1 = polar(cx, cy, outer, end)
  if (inner <= 0) return { d: `M${fmt(cx)} ${fmt(cy)}L${pair(o0)}A${fmt(outer)} ${fmt(outer)} 0 ${large} 1 ${pair(o1)}Z`, evenOdd: false }
  const i0 = polar(cx, cy, inner, start)
  const i1 = polar(cx, cy, inner, end)
  return { d: `M${pair(o0)}A${fmt(outer)} ${fmt(outer)} 0 ${large} 1 ${pair(o1)}L${pair(i1)}A${fmt(inner)} ${fmt(inner)} 0 ${large} 0 ${pair(i0)}Z`, evenOdd: false }
}
