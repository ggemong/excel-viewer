import { describe, expect, it } from 'vitest'
import type { CellStyle } from '../xlsx/cellStyle'
import type { CellModel } from '../xlsx/types'
import { createSpillResolver, type SpillInput } from './textSpill'

const CHROME = 21
const COL_W = 100 // 안쪽 79px
const GLYPH = 10 // 가짜 측정기: 글자당 10px

const cell = (value: string | number | null, style: CellStyle | null = null): CellModel => ({ address: 'A1', value, formula: null, numFmt: null, style, hyperlink: null })
const align = (h: 'left' | 'center' | 'right' | null, wrap = false): CellStyle => ({ bg: null, color: null, bold: false, italic: false, border: null, align: { h, v: null, wrap } })

function resolverFor(row: (CellModel | undefined)[], over: Partial<SpillInput> = {}) {
  const colCount = row.length
  return createSpillResolver({
    rows: [row],
    colWidths: new Array(colCount).fill(COL_W),
    visibleCols: Array.from({ length: colCount }, (_, i) => i + 1),
    merged: new Map(),
    isBlocked: () => false,
    cellChrome: CHROME,
    measure: (text) => text.length * GLYPH,
    ...over,
  })
}

describe('createSpillResolver', () => {
  it('칸 안에 들어가는 글은 넘치지 않는다', () => {
    expect(resolverFor([cell('가'.repeat(7)), undefined])(1, 1)).toBeNull() // 70px < 79px
  })

  it('긴 글은 오른쪽 빈 칸 위로 넘친다 — 필요한 만큼만(글 폭까지)', () => {
    const spill = resolverFor([cell('가'.repeat(12)), undefined, undefined])(1, 1) // 120px, 안쪽 79 -> 41px 더
    expect(spill?.maxWidth).toBeCloseTo(79 + 41 + 1, 5)
  })

  it('옆 칸에 값이 있으면 거기서 멈춘다(그 칸을 가리지 않는다) — 빈 칸 폭까지만', () => {
    // 빈 칸 하나(100px)만 있고 그다음은 값 -> 최대 79 + 100
    const spill = resolverFor([cell('가'.repeat(30)), undefined, cell('옆')])(1, 1)
    expect(spill?.maxWidth).toBeCloseTo(79 + 100 + 1, 5)
  })

  it('바로 옆 칸에 값이 있으면 넘치지 않는다', () => {
    expect(resolverFor([cell('가'.repeat(12)), cell('옆')])(1, 1)).toBeNull()
  })

  it('숫자 0도 값이다 — 막는다. 빈 문자열은 값이 아니다 — 막지 않는다', () => {
    expect(resolverFor([cell('가'.repeat(12)), cell(0)])(1, 1)).toBeNull()
    expect(resolverFor([cell('가'.repeat(12)), cell(''), undefined])(1, 1)).not.toBeNull()
  })

  it('숫자 칸, 줄바꿈 칸, 하이퍼링크 칸은 넘치지 않는다', () => {
    expect(resolverFor([cell(123456789012), undefined])(1, 1)).toBeNull()
    expect(resolverFor([cell('가'.repeat(12), align(null, true)), undefined])(1, 1)).toBeNull()
    expect(resolverFor([{ ...cell('가'.repeat(12)), hyperlink: 'https://a.example' }, undefined])(1, 1)).toBeNull()
  })

  it('병합에 걸린 칸은 넘치지도 넘침을 받지도 않는다', () => {
    const merged = new Map([['1,2', {}]])
    expect(resolverFor([cell('가'.repeat(12)), undefined], { merged })(1, 1)).toBeNull()
    const masterMerged = new Map([['1,1', {}]])
    expect(resolverFor([cell('가'.repeat(12)), undefined], { merged: masterMerged })(1, 1)).toBeNull()
  })

  it('필터 버튼이 달린 칸은 넘침을 막는다', () => {
    expect(resolverFor([cell('가'.repeat(12)), undefined], { isBlocked: (_r, c) => c === 2 })(1, 1)).toBeNull()
  })

  it('숨긴 열은 건너뛰어 그 너머의 빈 칸으로 넘친다', () => {
    const row = [cell('가'.repeat(12)), cell('숨김'), undefined]
    const spill = resolverFor(row, { visibleCols: [1, 3] })(1, 1)
    expect(spill).not.toBeNull()
  })

  it('오른쪽 정렬은 왼쪽 빈 칸으로 넘친다', () => {
    const row = [undefined, cell('가'.repeat(12), align('right'))]
    const spill = resolverFor(row)(1, 2)
    expect(spill?.maxWidth).toBeCloseTo(79 + 41 + 1, 5)
    expect(resolverFor([cell('왼'), cell('가'.repeat(12), align('right'))])(1, 2)).toBeNull()
  })

  it('가운데 정렬은 양쪽이 모두 비어 있어야 넘친다', () => {
    const text = '가'.repeat(12) // 41px 더 -> 양쪽 20.5px씩
    const both = resolverFor([undefined, cell(text, align('center')), undefined])(1, 2)
    expect(both?.maxWidth).toBeCloseTo(79 + 41 + 1, 5)
    expect(resolverFor([cell('왼'), cell(text, align('center')), undefined])(1, 2)).toBeNull()
    expect(resolverFor([undefined, cell(text, align('center')), cell('오')])(1, 2)).toBeNull()
  })

  it('글 폭을 잴 수 없으면(null) 넘침을 만들지 않는다 — 눈대중으로 판단하지 않는다', () => {
    expect(resolverFor([cell('가'.repeat(12)), undefined], { measure: () => null })(1, 1)).toBeNull()
  })

  it('결과는 칸별로 캐시된다', () => {
    let calls = 0
    const resolve = resolverFor([cell('가'.repeat(12)), undefined], { measure: (t) => (calls++, t.length * GLYPH) })
    resolve(1, 1)
    resolve(1, 1)
    expect(calls).toBe(1)
  })
})
