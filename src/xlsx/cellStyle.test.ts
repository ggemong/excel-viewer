import { describe, expect, it } from 'vitest'
import { cellStyleProps, extractCellStyle } from './cellStyle'

describe('extractCellStyle', () => {
  it('서식이 전혀 없으면 null', () => {
    expect(extractCellStyle({})).toBeNull()
  })

  it('단색 배경 + 명시된 글자색 + 굵게/기울임을 그대로 뽑는다', () => {
    const style = extractCellStyle({
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } },
      font: { bold: true, italic: true, color: { argb: 'FFFF0000' } },
    })
    expect(style).toEqual({ bg: '#FFFF00', color: '#FF0000', bold: true, italic: true })
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
    expect(extractCellStyle({ font: { bold: true } })).toEqual({ bg: null, color: null, bold: true, italic: false })
  })
})

describe('cellStyleProps', () => {
  const style = { bg: '#FFFF00', color: '#FF0000', bold: true, italic: false }

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
})
