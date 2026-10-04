import type { CSSProperties, ReactNode } from 'react'
import type { PlacedBox } from '../xlsx/drawingLayout'
import type { DrawingItem, DrawingNode, DrawingTransform, PictureNode, ShapeNode, ShapeText } from '../xlsx/drawingTypes'
import { ChartView } from './ChartView'
import { presetPath, rectPath } from './shapeGeometry'

/**
 * 시트 위에 떠 있는 그림·도형을 그린다. 셀 선택/복사를 방해하지 않도록 전부
 * pointer-events: none이고, 위치는 부모(기준 위치가 잡힌 컨테이너)에 대한 절대 좌표다.
 *
 * 차트는 ChartView(SVG)가 그린다. 한 컴포넌트에 그림/도형/묶음이 모여 있는 이유: 셋 다 "상자 하나를 회전·뒤집어서 그 안을
 * 채우는" 같은 구조라 변형(transform) 처리를 한 곳에서 공유하고, 묶음이 자식을 재귀로
 * 그릴 때 같은 NodeView를 그대로 쓰기 위해서다.
 */

export interface PlacedDrawing {
  item: DrawingItem
  /** sheet.drawings 안의 순서 — 검색 결과가 도형을 가리키는 키. */
  index: number
  box: PlacedBox
}

interface DrawingLayerProps {
  placed: PlacedDrawing[]
  /** 시트 좌표 -> 컨테이너 좌표 이동량(왼쪽 행번호 열, 위쪽 틀고정 높이 등). */
  offsetX: number
  offsetY: number
  urlFor: (blob: Blob) => string | undefined
  /** 검색 결과로 강조할 도형(drawings 인덱스 -> 상태). 안 주면 강조 없음. */
  searchState?: (index: number) => 'match' | 'current' | undefined
}

const JUSTIFY_BY_VALIGN: Record<'top' | 'middle' | 'bottom', CSSProperties['justifyContent']> = {
  top: 'flex-start',
  middle: 'center',
  bottom: 'flex-end',
}

/** OOXML은 "뒤집은 뒤 회전" 순서 — CSS transform은 오른쪽부터 적용되므로 rotate를 앞에 쓴다. */
function transformCss(t: DrawingTransform): CSSProperties {
  const parts: string[] = []
  if (t.rotation) parts.push(`rotate(${t.rotation}deg)`)
  if (t.flipH || t.flipV) parts.push(`scale(${t.flipH ? -1 : 1}, ${t.flipV ? -1 : 1})`)
  return parts.length > 0 ? { transform: parts.join(' '), transformOrigin: 'center' } : {}
}

function PictureView({ node, urlFor }: { node: PictureNode; urlFor: DrawingLayerProps['urlFor'] }) {
  if (!node.blob) {
    return <div className="drawing-unsupported">그림을 표시할 수 없어요{node.unsupportedFormat ? ` (.${node.unsupportedFormat})` : ''}</div>
  }
  const url = urlFor(node.blob)
  if (!url) return null

  const { l, t, r, b } = node.crop
  const visibleW = Math.max(0.01, 1 - l - r)
  const visibleH = Math.max(0.01, 1 - t - b)
  const cropped = l > 0 || t > 0 || r > 0 || b > 0
  // 자르기: 보이는 영역이 상자를 꽉 채우도록 원본을 키우고 잘린 만큼 밀어낸다.
  const imgStyle: CSSProperties = cropped
    ? { position: 'absolute', maxWidth: 'none', width: `${100 / visibleW}%`, height: `${100 / visibleH}%`, left: `${(-100 * l) / visibleW}%`, top: `${(-100 * t) / visibleH}%` }
    : { width: '100%', height: '100%', maxWidth: 'none' }

  return (
    <div className="drawing-fill" style={{ overflow: 'hidden' }}>
      <img src={url} alt={node.alt} draggable={false} decoding="async" style={imgStyle} />
    </div>
  )
}

function ShapeTextView({ text }: { text: ShapeText }) {
  return (
    <div
      className="drawing-fill"
      style={{
        display: 'flex',
        flexDirection: 'column',
        justifyContent: JUSTIFY_BY_VALIGN[text.vAlign],
        padding: `${text.insets.t}px ${text.insets.r}px ${text.insets.b}px ${text.insets.l}px`,
        boxSizing: 'border-box',
        overflow: 'hidden',
        whiteSpace: text.wrap ? 'pre-wrap' : 'pre',
      }}
    >
      {text.paragraphs.map((p, i) => (
        <div key={i} style={{ textAlign: p.align }}>
          {p.runs.map((run, j) =>
            run.text === '\n' ? (
              <br key={j} />
            ) : (
              <span key={j} style={{ fontSize: run.sizePx, fontWeight: run.bold ? 700 : undefined, fontStyle: run.italic ? 'italic' : undefined, color: run.color ?? undefined }}>
                {run.text}
              </span>
            ),
          )}
        </div>
      ))}
    </div>
  )
}

function ShapeView({ node, width, height }: { node: ShapeNode; width: number; height: number }) {
  // 지원하지 않는 프리셋은 사각형으로 근사한다(위치·채우기·글자는 살리고 모양만 단순화).
  const d = presetPath(node.preset, width, height, node.adjust) ?? rectPath(width, height)
  const line = node.line
  const dash = line ? (line.dash === 'dashed' ? `${line.width * 4} ${line.width * 3}` : line.dash === 'dotted' ? `${line.width} ${line.width * 2}` : undefined) : undefined

  return (
    <>
      <svg width={width} height={height} className="drawing-fill" style={{ overflow: 'visible' }} aria-hidden="true">
        <path d={d} fill={node.fill ?? 'none'} stroke={line?.color ?? 'none'} strokeWidth={line?.width} strokeDasharray={dash} strokeLinejoin="miter" />
      </svg>
      {node.text && <ShapeTextView text={node.text} />}
    </>
  )
}

function NodeView({ node, width, height, urlFor }: { node: DrawingNode; width: number; height: number; urlFor: DrawingLayerProps['urlFor'] }): ReactNode {
  switch (node.kind) {
    case 'picture':
      return <PictureView node={node} urlFor={urlFor} />
    case 'shape':
      return <ShapeView node={node} width={width} height={height} />
    case 'chart':
      return <ChartView node={node} width={width} height={height} />
    case 'group':
      return (
        <>
          {node.children.map((child, i) => (
            <div
              key={i}
              className="drawing-item"
              style={{
                left: `${child.rel.x * 100}%`,
                top: `${child.rel.y * 100}%`,
                width: `${child.rel.w * 100}%`,
                height: `${child.rel.h * 100}%`,
                ...transformCss(child.node),
              }}
            >
              <NodeView node={child.node} width={child.rel.w * width} height={child.rel.h * height} urlFor={urlFor} />
            </div>
          ))}
        </>
      )
  }
}

export function DrawingLayer({ placed, offsetX, offsetY, urlFor, searchState }: DrawingLayerProps) {
  return (
    <>
      {placed.map(({ item, box, index }) => (
        <div
          key={index}
          className="drawing-item"
          data-search={searchState?.(index)}
          style={{ left: offsetX + box.left, top: box.top - offsetY, width: box.width, height: box.height, ...transformCss(item.node) }}
        >
          <NodeView node={item.node} width={box.width} height={box.height} urlFor={urlFor} />
        </div>
      ))}
    </>
  )
}
