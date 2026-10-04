/**
 * 시트 위에 떠 있는 그림·도형의 화면용 모델. xlsx의 `xl/drawings/drawingN.xml`(DrawingML)을
 * src/xlsx/drawing.ts가 읽어서 이 모양으로 만든다. 읽기 전용 뷰어라 다시 쓰는 경로는 없다.
 *
 * 좌표 단위는 전부 px(EMU/9525)이고, 셀 기준 위치(앵커)와 객체 자체의 변형(회전·뒤집기)을
 * 분리해서 들고 있다 — 실제 픽셀 위치는 열너비/행높이를 알아야 정해지므로 렌더링 시점에
 * (src/grid/drawingLayout.ts) 계산한다.
 */
import type { ChartNode } from './chartTypes'

/** 셀 안 위치. col/row는 0-based, 오프셋은 그 셀 왼쪽 위에서의 px. */
export interface CellPoint {
  col: number
  row: number
  colOffset: number
  rowOffset: number
}

export type DrawingAnchor =
  /** 두 셀 사이에 맞춰 늘어난다(열/행 크기를 따라감). */
  | { kind: 'twoCell'; from: CellPoint; to: CellPoint }
  /** 시작 셀 + 고정 크기. */
  | { kind: 'oneCell'; from: CellPoint; width: number; height: number }
  /** 시트 왼쪽 위 기준 절대 좌표. */
  | { kind: 'absolute'; x: number; y: number; width: number; height: number }

export interface DrawingTransform {
  /** 시계 방향 도(degree). */
  rotation: number
  flipH: boolean
  flipV: boolean
}

/** 부모 상자(그림 묶음) 크기 대비 비율 위치. */
export interface RelativeRect {
  x: number
  y: number
  w: number
  h: number
}

export interface PictureNode extends DrawingTransform {
  kind: 'picture'
  /** 브라우저가 그릴 수 있는 형식이면 Blob, 아니면 null(unsupportedFormat에 확장자). */
  blob: Blob | null
  unsupportedFormat: string | null
  /** 자르기: 원본 이미지에서 잘라낸 비율(0~1). */
  crop: { l: number; t: number; r: number; b: number }
  alt: string
}

export interface ShapeLine {
  color: string
  /** px */
  width: number
  dash: 'solid' | 'dashed' | 'dotted'
}

export interface TextRun {
  text: string
  sizePx: number
  bold: boolean
  italic: boolean
  /** null이면 앱의 기본 글자색을 따른다(Excel의 "자동" 색이 다크 테마에서 안 보이는 문제 회피). */
  color: string | null
}

export interface TextParagraph {
  align: 'left' | 'center' | 'right'
  runs: TextRun[]
}

export interface ShapeText {
  paragraphs: TextParagraph[]
  vAlign: 'top' | 'middle' | 'bottom'
  wrap: boolean
  /** px */
  insets: { l: number; t: number; r: number; b: number }
}

export interface ShapeNode extends DrawingTransform {
  kind: 'shape'
  /** DrawingML 프리셋 도형 이름(rect, rightArrow, wedgeRectCallout ...). */
  preset: string
  /** 프리셋 조절값(adj1, adj2 ...) — 100000 = 1.0 단위 그대로. */
  adjust: Record<string, number>
  /** CSS 색(투명도 포함 가능). null이면 채우기 없음. */
  fill: string | null
  line: ShapeLine | null
  text: ShapeText | null
}

export interface GroupNode extends DrawingTransform {
  kind: 'group'
  children: { rel: RelativeRect; node: DrawingNode }[]
}

export type DrawingNode = PictureNode | ShapeNode | GroupNode | ChartNode

export interface DrawingItem {
  anchor: DrawingAnchor
  node: DrawingNode
}
