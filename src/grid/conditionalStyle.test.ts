import { describe, expect, it } from 'vitest'
import type { CfRule, ConditionalFormat } from '../xlsx/conditionalFormat'
import type { CellModel } from '../xlsx/types'
import { createConditionalStyler } from './conditionalStyle'

const TODAY = 46299
const clock = { today: TODAY, now: TODAY }

const cell = (value: string | number | null): CellModel => ({ address: 'A1', value, formula: null, numFmt: null, style: null, hyperlink: null })

/** 행: 1=헤더, 2..: 데이터. 열: J(10)=기한, K(11)=상태코드, L(12)=처리표시 */
function rowsOf(data: { J?: number; K?: number; L?: string }[]): (CellModel | undefined)[][] {
  const row = (d: { J?: number; K?: number; L?: string }) => {
    const r: (CellModel | undefined)[] = new Array(12).fill(undefined)
    if (d.J !== undefined) r[9] = cell(d.J)
    if (d.K !== undefined) r[10] = cell(d.K)
    if (d.L !== undefined) r[11] = cell(d.L)
    return r
  }
  return [new Array(12).fill(undefined), ...data.map(row)]
}

const rule = (over: Partial<CfRule> & Pick<CfRule, 'formulae' | 'style'>): CfRule => ({ priority: 1, stopIfTrue: false, kind: 'expression', operator: null, ...over })
const fmt = (sqref: [number, number, number, number][], rules: CfRule[]): ConditionalFormat => ({
  ranges: sqref.map(([r0, c0, r1, c1]) => ({ r0, c0, r1, c1 })),
  rules,
})

describe('createConditionalStyler', () => {
  it('규칙이 없으면 null(평가 비용을 아예 안 씀)', () => {
    expect(createConditionalStyler({ rows: [], conditionalFormats: [] })).toBeNull()
  })

  it('행 전체 색칠: 기준 칸(A2) 기준 상대 참조가 행마다 밀려 그 행의 상태 열을 본다', () => {
    const rows = rowsOf([{ K: 1 }, { K: 2 }, { K: 1 }])
    const styler = createConditionalStyler({ rows, conditionalFormats: [fmt([[2, 1, 4, 14]], [rule({ formulae: ['$K2=1'], style: { bg: '#FFC7CE' } })])] }, clock)!
    expect(styler(2, 1)).toEqual({ bg: '#FFC7CE' })
    expect(styler(2, 14)).toEqual({ bg: '#FFC7CE' })
    expect(styler(3, 1)).toBeNull()
    expect(styler(4, 5)).toEqual({ bg: '#FFC7CE' })
  })

  it('적용 범위 밖은 칠하지 않는다', () => {
    const styler = createConditionalStyler({ rows: rowsOf([{ K: 1 }]), conditionalFormats: [fmt([[2, 1, 2, 3]], [rule({ formulae: ['$K2=1'], style: { bg: '#FF0000' } })])] }, clock)!
    expect(styler(2, 4)).toBeNull()
    expect(styler(2, 3)).not.toBeNull()
  })

  it('여러 범위(sqref)와 첫 범위 기준 앵커', () => {
    const rows = rowsOf([{ K: 1 }, { K: 1 }, { K: 1 }])
    const styler = createConditionalStyler(
      { rows, conditionalFormats: [fmt([[2, 1, 2, 2], [4, 1, 4, 2]], [rule({ formulae: ['$K2=1'], style: { bg: '#00FF00' } })])] },
      clock,
    )!
    expect(styler(2, 1)).not.toBeNull()
    expect(styler(3, 1)).toBeNull()
    expect(styler(4, 2)).not.toBeNull()
  })

  it('여러 규칙이 동시에 참이면 속성별로 합쳐지고, 같은 속성은 우선순위 높은 규칙이 이긴다', () => {
    const rows = rowsOf([{ J: TODAY - 5, L: '불가' }])
    const styler = createConditionalStyler(
      {
        rows,
        conditionalFormats: [
          fmt([[2, 1, 2, 14]], [
            rule({ priority: 3, formulae: ['$L2="불가"'], style: { bg: '#DDDDDD' } }),
            rule({ priority: 1, formulae: ['$J2<TODAY()'], style: { color: '#FF0000' } }),
            rule({ priority: 2, formulae: ['$L2="불가"'], style: { bg: '#FFF2CC', strike: true } }),
          ]),
        ],
      },
      clock,
    )!
    expect(styler(2, 1)).toEqual({ color: '#FF0000', bg: '#FFF2CC', strike: true })
  })

  it('stopIfTrue: 참이면 아래 우선순위 규칙은 보지 않는다', () => {
    const styler = createConditionalStyler(
      {
        rows: rowsOf([{ K: 1 }]),
        conditionalFormats: [
          fmt([[2, 1, 2, 3]], [
            rule({ priority: 1, stopIfTrue: true, formulae: ['$K2=1'], style: { bg: '#111111' } }),
            rule({ priority: 2, formulae: ['$K2=1'], style: { color: '#EEEEEE' } }),
          ]),
        ],
      },
      clock,
    )!
    expect(styler(2, 1)).toEqual({ bg: '#111111' })
  })

  it('블록이 달라도 우선순위는 시트 전체에서 비교한다', () => {
    const styler = createConditionalStyler(
      {
        rows: rowsOf([{ K: 1 }]),
        conditionalFormats: [
          fmt([[2, 1, 2, 3]], [rule({ priority: 9, formulae: ['$K2=1'], style: { bg: '#999999' } })]),
          fmt([[2, 1, 2, 3]], [rule({ priority: 1, formulae: ['$K2=1'], style: { bg: '#111111' } })]),
        ],
      },
      clock,
    )!
    expect(styler(2, 1)?.bg).toBe('#111111')
  })

  it('cellIs: 보다 큼 / 보다 작음 / 사이 — 빈 칸은 0으로 비교된다(Excel과 동일)', () => {
    const rows: (CellModel | undefined)[][] = [[cell(150)], [cell(50)], [undefined]]
    const make = (operator: string, formulae: string[]) =>
      createConditionalStyler({ rows, conditionalFormats: [fmt([[1, 1, 3, 1]], [rule({ kind: 'cellIs', operator, formulae, style: { bg: '#F00' } })])] }, clock)!

    const gt = make('greaterThan', ['100'])
    expect(gt(1, 1)).not.toBeNull()
    expect(gt(2, 1)).toBeNull()
    const lt = make('lessThan', ['100'])
    expect(lt(2, 1)).not.toBeNull()
    expect(lt(3, 1)).not.toBeNull()
    const between = make('between', ['40', '60'])
    expect(between(2, 1)).not.toBeNull()
    expect(between(1, 1)).toBeNull()
  })

  it('이 파일의 기한 초과 규칙: 처리 표시 없음 + 기한 지남 + 기한 있음', () => {
    const rows = rowsOf([{ J: TODAY - 2 }, { J: TODAY - 2, L: 'O' }, { J: TODAY + 2 }, {}])
    const styler = createConditionalStyler(
      { rows, conditionalFormats: [fmt([[2, 1, 5, 14]], [rule({ formulae: ['AND($L2="",$J2<TODAY(),$J2<>"")'], style: { color: '#FF0000' } })])] },
      clock,
    )!
    expect(styler(2, 1)?.color).toBe('#FF0000')
    expect(styler(3, 1)).toBeNull()
    expect(styler(4, 1)).toBeNull()
    expect(styler(5, 1)).toBeNull()
  })

  it('결과는 칸별로 캐시된다(같은 칸은 같은 객체)', () => {
    const styler = createConditionalStyler({ rows: rowsOf([{ K: 1 }]), conditionalFormats: [fmt([[2, 1, 2, 3]], [rule({ formulae: ['$K2=1'], style: { bg: '#123456' } })])] }, clock)!
    expect(styler(2, 1)).toBe(styler(2, 1))
  })
})
