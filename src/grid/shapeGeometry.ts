/**
 * DrawingML 프리셋 도형(rect, rightArrow, wedgeRectCallout ...)을 SVG path로 바꾼다.
 * 공식은 ECMA-376의 presetShapeDefinitions를 따르되, 실제 업무 파일에서 쓰이는 것만
 * 구현한다(사각형·둥근 사각형·타원·삼각형·마름모·직선·4방향 화살표·말풍선 3종). 나머지
 * 프리셋은 null을 돌려주고, 호출부가 사각형으로 대신 그린다 — 도형 위치와 텍스트는 살리되
 * 모양만 근사한다.
 *
 * 좌표는 도형 상자 왼쪽 위가 (0,0), 크기가 (w,h)인 px 공간. 말풍선 꼬리처럼 상자 밖으로
 * 나가는 부분이 있으므로 그리는 쪽 SVG는 overflow를 열어 둬야 한다.
 */

const UNIT = 100_000
const DEFAULT_ARROW_ADJUST = 50_000
const DEFAULT_ROUND_ADJUST = 16_667
const DEFAULT_TRIANGLE_ADJUST = 50_000
const DEFAULT_CALLOUT_ADJ1 = -20_833
const DEFAULT_CALLOUT_ADJ2 = 62_500
/** 말풍선 꼬리 밑변을 12등분한 위치(ECMA 정의 그대로): 꼬리 방향에 따라 7~10 또는 2~5. */
const CALLOUT_BASE_NEAR = [2, 5] as const
const CALLOUT_BASE_FAR = [7, 10] as const
const CALLOUT_GRID = 12
/** 타원 말풍선 꼬리 밑변의 반각(도). */
const ELLIPSE_CALLOUT_HALF_ANGLE_DEG = 12

type Adjust = Record<string, number>

const n = (v: number) => String(Number(v.toFixed(2)))
const pin = (lo: number, v: number, hi: number) => Math.max(lo, Math.min(hi, v))
const poly = (points: [number, number][]) => `M${points.map(([x, y]) => `${n(x)} ${n(y)}`).join('L')}Z`

export function rectPath(w: number, h: number): string {
  return `M0 0H${n(w)}V${n(h)}H0Z`
}

function roundRectPath(w: number, h: number, radius: number): string {
  const r = pin(0, radius, Math.min(w, h) / 2)
  return `M${n(r)} 0H${n(w - r)}A${n(r)} ${n(r)} 0 0 1 ${n(w)} ${n(r)}V${n(h - r)}A${n(r)} ${n(r)} 0 0 1 ${n(w - r)} ${n(h)}H${n(r)}A${n(r)} ${n(r)} 0 0 1 0 ${n(h - r)}V${n(r)}A${n(r)} ${n(r)} 0 0 1 ${n(r)} 0Z`
}

function ellipsePath(w: number, h: number): string {
  const rx = w / 2
  const ry = h / 2
  return `M0 ${n(ry)}A${n(rx)} ${n(ry)} 0 1 1 ${n(w)} ${n(ry)}A${n(rx)} ${n(ry)} 0 1 1 0 ${n(ry)}Z`
}

function horizontalArrow(w: number, h: number, adjust: Adjust, direction: 'right' | 'left'): string {
  const ss = Math.min(w, h)
  const a1 = pin(0, adjust.adj1 ?? DEFAULT_ARROW_ADJUST, UNIT)
  const a2 = pin(0, adjust.adj2 ?? DEFAULT_ARROW_ADJUST, ss > 0 ? (UNIT * w) / ss : 0)
  const headLen = (ss * a2) / UNIT
  const shaft = (h * a1) / (2 * UNIT)
  const y1 = h / 2 - shaft
  const y2 = h / 2 + shaft
  if (direction === 'right') {
    const x1 = w - headLen
    return poly([[0, y1], [x1, y1], [x1, 0], [w, h / 2], [x1, h], [x1, y2], [0, y2]])
  }
  return poly([[w, y1], [headLen, y1], [headLen, 0], [0, h / 2], [headLen, h], [headLen, y2], [w, y2]])
}

function verticalArrow(w: number, h: number, adjust: Adjust, direction: 'down' | 'up'): string {
  const ss = Math.min(w, h)
  const a1 = pin(0, adjust.adj1 ?? DEFAULT_ARROW_ADJUST, UNIT)
  const a2 = pin(0, adjust.adj2 ?? DEFAULT_ARROW_ADJUST, ss > 0 ? (UNIT * h) / ss : 0)
  const headLen = (ss * a2) / UNIT
  const shaft = (w * a1) / (2 * UNIT)
  const x1 = w / 2 - shaft
  const x2 = w / 2 + shaft
  if (direction === 'down') {
    const y1 = h - headLen
    return poly([[x1, 0], [x2, 0], [x2, y1], [w, y1], [w / 2, h], [0, y1], [x1, y1]])
  }
  return poly([[x1, h], [x2, h], [x2, headLen], [w, headLen], [w / 2, 0], [0, headLen], [x1, headLen]])
}

/** 사각 말풍선 꼬리 계산(ECMA wedgeRectCallout 정의): 꼬리가 어느 변에 붙고 밑변이 어디인지. */
function wedgeGeometry(w: number, h: number, adjust: Adjust) {
  const dxPos = (w * (adjust.adj1 ?? DEFAULT_CALLOUT_ADJ1)) / UNIT
  const dyPos = (h * (adjust.adj2 ?? DEFAULT_CALLOUT_ADJ2)) / UNIT
  const xPos = w / 2 + dxPos
  const yPos = h / 2 + dyPos
  // 꼬리 끝이 위/아래 변에 더 가까운가(dz>0) 좌/우 변에 더 가까운가
  const verticalTail = Math.abs(dyPos) - Math.abs((dxPos * h) / w) > 0

  const [xg1, xg2] = dxPos > 0 ? CALLOUT_BASE_FAR : CALLOUT_BASE_NEAR
  const [yg1, yg2] = dyPos > 0 ? CALLOUT_BASE_FAR : CALLOUT_BASE_NEAR
  const x1 = (w * xg1) / CALLOUT_GRID
  const x2 = (w * xg2) / CALLOUT_GRID
  const y1 = (h * yg1) / CALLOUT_GRID
  const y2 = (h * yg2) / CALLOUT_GRID

  return {
    x1,
    x2,
    y1,
    y2,
    // 각 변에서 꼬리 끝이 찍히는 점(꼬리가 그 변에 없으면 변 위의 평범한 점으로 퇴화)
    left: verticalTail ? [0, y1] : [dxPos > 0 ? 0 : xPos, dxPos > 0 ? y1 : yPos],
    top: verticalTail ? [dyPos > 0 ? x1 : xPos, dyPos > 0 ? 0 : yPos] : [x1, 0],
    right: verticalTail ? [w, y1] : [dxPos > 0 ? xPos : w, dxPos > 0 ? yPos : y1],
    bottom: verticalTail ? [dyPos > 0 ? xPos : x1, dyPos > 0 ? yPos : h] : [x1, h],
  }
}

function wedgeRectPath(w: number, h: number, adjust: Adjust): string {
  const g = wedgeGeometry(w, h, adjust)
  return poly([
    [0, 0], [g.x1, 0], g.top as [number, number], [g.x2, 0], [w, 0],
    [w, g.y1], g.right as [number, number], [w, g.y2], [w, h],
    [g.x2, h], g.bottom as [number, number], [g.x1, h], [0, h],
    [0, g.y2], g.left as [number, number], [0, g.y1],
  ])
}

function wedgeRoundRectPath(w: number, h: number, adjust: Adjust): string {
  const g = wedgeGeometry(w, h, adjust)
  const r = pin(0, (Math.min(w, h) * (adjust.adj3 ?? DEFAULT_ROUND_ADJUST)) / UNIT, Math.min(w, h) / 2)
  const arc = (x: number, y: number) => `A${n(r)} ${n(r)} 0 0 1 ${n(x)} ${n(y)}`
  const L = (p: number[]) => `L${n(p[0])} ${n(p[1])}`
  return [
    `M0 ${n(r)}`, arc(r, 0),
    `L${n(g.x1)} 0`, L(g.top), `L${n(g.x2)} 0`, `L${n(w - r)} 0`, arc(w, r),
    `L${n(w)} ${n(g.y1)}`, L(g.right), `L${n(w)} ${n(g.y2)}`, `L${n(w)} ${n(h - r)}`, arc(w - r, h),
    `L${n(g.x2)} ${n(h)}`, L(g.bottom), `L${n(g.x1)} ${n(h)}`, `L${n(r)} ${n(h)}`, arc(0, h - r),
    `L0 ${n(g.y2)}`, L(g.left), `L0 ${n(g.y1)}`, 'Z',
  ].join('')
}

function wedgeEllipsePath(w: number, h: number, adjust: Adjust): string {
  const rx = w / 2
  const ry = h / 2
  const tipX = rx + (w * (adjust.adj1 ?? DEFAULT_CALLOUT_ADJ1)) / UNIT
  const tipY = ry + (h * (adjust.adj2 ?? DEFAULT_CALLOUT_ADJ2)) / UNIT
  // 꼬리 방향의 타원 매개각 t (점 = 중심 + (rx cos t, ry sin t))
  const t = Math.atan2((tipY - ry) / ry, (tipX - rx) / rx)
  const half = (ELLIPSE_CALLOUT_HALF_ANGLE_DEG * Math.PI) / 180
  const at = (a: number) => `${n(rx + rx * Math.cos(a))} ${n(ry + ry * Math.sin(a))}`
  // 꼬리 밑변 한 점에서 큰 호로 반대편 점까지 타원을 돌고, 꼬리 끝으로 갔다가 닫는다.
  return `M${at(t + half)}A${n(rx)} ${n(ry)} 0 1 1 ${at(t - half)}L${n(tipX)} ${n(tipY)}Z`
}

/**
 * 프리셋 이름 -> SVG path. 지원하지 않는 프리셋이면 null(호출부가 사각형으로 근사).
 * @param adjust adj1, adj2 ... 조절값(100000 = 1.0)
 */
export function presetPath(preset: string, w: number, h: number, adjust: Adjust): string | null {
  switch (preset) {
    case 'rect':
      return rectPath(w, h)
    case 'roundRect':
      return roundRectPath(w, h, (Math.min(w, h) * (adjust.adj ?? DEFAULT_ROUND_ADJUST)) / UNIT)
    case 'ellipse':
      return ellipsePath(w, h)
    case 'triangle':
      return poly([[(w * (adjust.adj ?? DEFAULT_TRIANGLE_ADJUST)) / UNIT, 0], [w, h], [0, h]])
    case 'diamond':
      return poly([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]])
    case 'line':
    case 'straightConnector1':
      return `M0 0L${n(w)} ${n(h)}`
    case 'rightArrow':
      return horizontalArrow(w, h, adjust, 'right')
    case 'leftArrow':
      return horizontalArrow(w, h, adjust, 'left')
    case 'downArrow':
      return verticalArrow(w, h, adjust, 'down')
    case 'upArrow':
      return verticalArrow(w, h, adjust, 'up')
    case 'wedgeRectCallout':
      return wedgeRectPath(w, h, adjust)
    case 'wedgeRoundRectCallout':
      return wedgeRoundRectPath(w, h, adjust)
    case 'wedgeEllipseCallout':
      return wedgeEllipsePath(w, h, adjust)
    default:
      return null
  }
}
