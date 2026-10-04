import { memo, useMemo } from 'react'
import { contrastColor } from '../xlsx/cellStyle'
import type { ChartNode } from '../xlsx/chartTypes'
import { textWidthAt } from '../xlsx/textMetrics'
import { buildChartScene, type SceneItem } from './chartScene'
import { useFontsVersion } from './useFontsVersion'

/**
 * 차트를 SVG로 그린다. 무엇을 어디에 그릴지는 전부 chartScene.ts(순수 함수)가 정하고, 여기는 그 도형 목록을 SVG 요소로
 * 옮기기만 한다. 시트 위 다른 개체처럼 pointer-events: none이라 셀 선택을 방해하지 않는다(.drawing-item).
 */

/** 캔버스를 못 쓰는 환경에서만 쓰는 글자 폭 비율(글자당 폭 / 크기). 브라우저에서는 실제 글꼴로 잰다. */
const FALLBACK_GLYPH_RATIO = 0.55

const measure = (text: string, size: number, bold: boolean) => textWidthAt(text, size, bold) ?? text.length * size * FALLBACK_GLYPH_RATIO

/** 차트 영역 배경에 맞는 기본 글자색. 배경이 없으면(투명) 앱의 글자색을 따른다. */
function textColorFor(chart: ChartNode): string {
  return chart.background && /^#[0-9a-fA-F]{6}$/.test(chart.background) ? contrastColor(chart.background) : 'var(--text)'
}

function SceneNode({ item }: { item: SceneItem }) {
  switch (item.t) {
    case 'rect':
      return <rect x={item.x} y={item.y} width={Math.max(0, item.w)} height={Math.max(0, item.h)} fill={item.fill} stroke={item.stroke} strokeWidth={item.strokeW} />
    case 'line':
      return <line x1={item.x1} y1={item.y1} x2={item.x2} y2={item.y2} stroke={item.stroke} strokeWidth={item.strokeW} />
    case 'circle':
      return <circle cx={item.cx} cy={item.cy} r={item.r} fill={item.fill} />
    case 'path':
      return <path d={item.d} fill={item.fill ?? 'none'} stroke={item.stroke ?? 'none'} strokeWidth={item.strokeW} strokeLinejoin="round" strokeLinecap="round" fillRule={item.evenOdd ? 'evenodd' : undefined} />
    case 'text':
      return (
        <text
          x={item.x}
          y={item.y}
          fontSize={item.size}
          fontWeight={item.bold ? 700 : undefined}
          textAnchor={item.anchor}
          dominantBaseline={item.baseline === 'middle' ? 'central' : item.baseline === 'hanging' ? 'hanging' : undefined}
          fill={item.fill}
          transform={item.rotate ? `rotate(${item.rotate} ${item.x} ${item.y})` : undefined}
        >
          {item.text}
        </text>
      )
  }
}

interface ChartViewProps {
  node: ChartNode
  width: number
  height: number
}

/** 차트 모델이 그대로면 선택/스크롤로 시트가 다시 그려져도 도형 계산을 다시 하지 않는다. */
export const ChartView = memo(function ChartView({ node, width, height }: ChartViewProps) {
  // 글꼴이 늦게 도착하면 글자 폭이 달라져 범례·축 여백이 바뀌므로 그때 다시 계산한다.
  const fontsVersion = useFontsVersion()
  // fontsVersion은 값으로 쓰지 않지만 "글꼴이 바뀌었으니 다시 재라"는 신호라 의존성에 둔다.
  // oxlint-disable-next-line react-hooks/exhaustive-deps
  const scene = useMemo(() => buildChartScene(node, width, height, measure), [node, width, height, fontsVersion])
  const label = node.title?.text ?? '차트'
  return (
    <svg width={width} height={height} className="drawing-fill" style={{ color: textColorFor(node), overflow: 'hidden' }} role="img" aria-label={label}>
      {scene.items.map((item, i) => (
        <SceneNode key={i} item={item} />
      ))}
    </svg>
  )
})
