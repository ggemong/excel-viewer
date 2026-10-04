import { describe, expect, it } from 'vitest'
import { applyConditionalStyle, cellStyleProps, extractCellStyle } from './cellStyle'
import type { ThemeColors } from './themeColor'

const theme: ThemeColors = {
  slots: ['#FFFFFF', '#000000', '#EEECE1', '#1F497D', '#4F81BD', '', '', '', '', '', '', ''],
}

describe('extractCellStyle', () => {
  it('서식이 전혀 없으면 null', () => {
    expect(extractCellStyle({})).toBeNull()
  })

  it('단색 배경 + 명시된 글자색 + 굵게/기울임을 그대로 뽑는다', () => {
    const style = extractCellStyle({
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } },
      font: { bold: true, italic: true, color: { argb: 'FFFF0000' } },
    })
    expect(style).toEqual({ bg: '#FFFF00', color: '#FF0000', bold: true, italic: true, border: null, align: null, colorExplicit: true })
  })

  it('배경만 있고 글자색이 없으면 밝기에 맞춰 대비색을 자동으로 고른다', () => {
    const onLight = extractCellStyle({ fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } } })
    const onDark = extractCellStyle({ fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF000000' } } })
    expect(onLight?.color).toBe('#1a1a1a')
    expect(onDark?.color).toBe('#f5f5f5')
  })

  it('단색이 아닌 채우기(패턴/그라디언트)는 배경을 무시한다', () => {
    const style = extractCellStyle({ fill: { type: 'pattern', pattern: 'darkTrellis', fgColor: { argb: 'FFFF0000' } } })
    expect(style?.bg ?? null).toBeNull()
  })

  it('굵게만 있어도(배경/글자색 없이) null이 아니다', () => {
    expect(extractCellStyle({ font: { bold: true } })).toEqual({
      bg: null,
      color: null,
      bold: true,
      italic: false,
      border: null,
      align: null,
    })
  })

  it('argb가 없으면 테마 인덱스로 배경색을 해석한다(Excel 기본 팔레트가 실제로 이 경우)', () => {
    const style = extractCellStyle({ fill: { type: 'pattern', pattern: 'solid', fgColor: { theme: 4 } } }, theme)
    expect(style?.bg).toBe('#4F81BD')
  })

  it('테마 팔레트를 안 넘기면(theme 인자 생략) 테마 색은 그냥 무시한다', () => {
    const style = extractCellStyle({ fill: { type: 'pattern', pattern: 'solid', fgColor: { theme: 4 } } })
    expect(style?.bg ?? null).toBeNull()
  })

  it('argb와 theme가 둘 다 있으면 argb가 우선(더 구체적인 지정)', () => {
    const style = extractCellStyle(
      { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFABCDEF', theme: 4 } } },
      theme,
    )
    expect(style?.bg).toBe('#ABCDEF')
  })

  it('테두리 4방향을 CSS 두께/스타일/색으로 변환한다', () => {
    const style = extractCellStyle({
      border: {
        top: { style: 'thin', color: { argb: 'FFFF0000' } },
        bottom: { style: 'thick' },
        left: { style: 'dashed', color: { argb: 'FF00FF00' } },
      },
    })
    expect(style?.border).toEqual({
      top: { width: '1px', style: 'solid', color: '#FF0000' },
      right: null,
      bottom: { width: '3px', style: 'solid', color: 'var(--line-strong)' },
      left: { width: '1px', style: 'dashed', color: '#00FF00' },
    })
  })

  it('정렬/줄바꿈을 뽑는다(가운데 세로 정렬은 OOXML의 center를 middle로 받는다)', () => {
    const style = extractCellStyle({ alignment: { horizontal: 'center', vertical: 'middle', wrapText: true } })
    expect(style?.align).toEqual({ h: 'center', v: 'middle', wrap: true })
  })

  it('알 수 없는 정렬 값(distributed 등)만 있으면 의미 있는 서식이 없는 것과 같아 null', () => {
    expect(extractCellStyle({ alignment: { horizontal: 'distributed' } })).toBeNull()
  })

  it('알 수 없는 정렬 값이어도 굵게처럼 다른 서식이 같이 있으면 그 서식은 살아있다', () => {
    const style = extractCellStyle({ font: { bold: true }, alignment: { horizontal: 'distributed' } })
    expect(style?.align).toBeNull()
    expect(style?.bold).toBe(true)
  })

  it('테두리만 지정해도 ExcelJS가 같이 붙이는 기본 폰트(theme:1, 보정 없음)는 글자색으로 안 믿는다', () => {
    // 실제 ExcelJS round-trip으로 확인된 상황: 사용자가 글자색을 건드린 적 없어도
    // 테두리 하나만 지정하면 font.color={theme:1}이 같이 저장된다 — 이걸 그대로
    // 믿으면 이 앱의 다크 테마에서 검정 글자가 안 보인다.
    const style = extractCellStyle(
      {
        border: { top: { style: 'thin', color: { argb: 'FFBBBBBB' } } },
        font: { color: { theme: 1 } },
      },
      theme,
    )
    expect(style?.color).toBeNull()
  })

  it('theme:1이어도 tint 보정이 있으면 사용자가 실제로 고른 색이니 그대로 쓴다', () => {
    const style = extractCellStyle({ font: { color: { theme: 1, tint: 0.5 } } }, theme)
    expect(style?.color).not.toBeNull()
  })
})

describe('cellStyleProps', () => {
  const style = { bg: '#FFFF00', color: '#FF0000', bold: true, italic: false, border: null, align: null }

  it('style이 null이면 빈 객체', () => {
    expect(cellStyleProps(null, false)).toEqual({})
  })

  it('suppressBg가 false면 배경을 포함한다', () => {
    const props = cellStyleProps(style, false)
    expect(props.background).toBe('#FFFF00')
    expect(props.color).toBe('#FF0000')
    expect(props.fontWeight).toBe(700)
  })

  it('suppressBg가 true면 배경만 비우고 글자색/굵기는 유지한다(선택·비교 강조색이 우선)', () => {
    const props = cellStyleProps(style, true)
    expect(props.background).toBeUndefined()
    expect(props.color).toBe('#FF0000')
    expect(props.fontWeight).toBe(700)
  })

  it('정렬 지정이 없으면 justifyContent/alignItems 키 자체를 안 넣는다(숫자 추측 fallback이 안 덮어써지게)', () => {
    const props = cellStyleProps(style, false)
    expect('justifyContent' in props).toBe(false)
    expect('alignItems' in props).toBe(false)
  })

  it('정렬이 지정돼 있으면 justifyContent/alignItems/whiteSpace로 매핑한다', () => {
    const aligned = { ...style, align: { h: 'center' as const, v: 'top' as const, wrap: true } }
    const props = cellStyleProps(aligned, false)
    expect(props.justifyContent).toBe('center')
    expect(props.alignItems).toBe('flex-start')
    expect(props.whiteSpace).toBe('normal')
  })

  it('테두리를 border-{side} 문자열로 합친다', () => {
    const bordered = {
      ...style,
      border: { top: { width: '1px', style: 'solid', color: '#FF0000' }, right: null, bottom: null, left: null },
    }
    const props = cellStyleProps(bordered, false)
    expect(props.borderTop).toBe('1px solid #FF0000')
    expect(props.borderRight).toBeUndefined()
  })
})

describe('applyConditionalStyle', () => {
  it('조건부서식이 정한 속성만 덮어쓰고 나머지 셀 서식은 남긴다', () => {
    const base = { bg: '#FFFFFF', color: '#112233', bold: true, italic: false, border: null, align: null, colorExplicit: true as const }
    const out = applyConditionalStyle(base, { bg: '#FFC7CE', strike: true })
    expect(out).toMatchObject({ bg: '#FFC7CE', bold: true, strike: true })
    expect(out.color).toBe('#112233')
  })

  it('배경만 바뀌고 글자색이 자동 계산된 값이었으면 새 배경에 맞춰 다시 고른다', () => {
    const base = { bg: '#000000', color: '#f5f5f5', bold: false, italic: false, border: null, align: null }
    expect(applyConditionalStyle(base, { bg: '#FFFFFF' }).color).toBe('#1a1a1a')
  })

  it('원래 서식이 없던 칸에 연한 배경만 입혀도 읽히는 글자색이 붙는다', () => {
    const out = applyConditionalStyle(null, { bg: '#FFF2CC' })
    expect(out.bg).toBe('#FFF2CC')
    expect(out.color).toBe('#1a1a1a')
  })

  it('조건부서식이 글자색을 정하면 그 색을 명시 색으로 쓴다', () => {
    expect(applyConditionalStyle(null, { color: '#FF0000' })).toMatchObject({ color: '#FF0000', colorExplicit: true })
  })

  it('취소선/밑줄은 textDecoration으로 변환된다', () => {
    const out = applyConditionalStyle(null, { strike: true, underline: true })
    expect(cellStyleProps(out, false).textDecoration).toBe('line-through underline')
  })
})
