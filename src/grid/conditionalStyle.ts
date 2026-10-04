/**
 * 칸 하나에 조건부서식 규칙을 적용한 결과 스타일을 계산한다(수식 평가는 src/formula).
 *
 * Excel 규칙 적용 방식:
 *  - 규칙은 시트 전체에서 우선순위(priority) 숫자가 작은 것부터 본다.
 *  - 여러 규칙이 동시에 참이면 **속성별로** 합쳐지고, 같은 속성(예: 채우기)이 겹치면 우선순위가 높은 규칙이 이긴다.
 *    (그래서 "빨간 글자" 규칙과 "연한 배경+취소선" 규칙이 함께 걸려도 둘 다 보인다.)
 *  - stopIfTrue인 규칙이 참이면 그 아래 규칙은 보지 않는다.
 *  - 수식의 상대 참조는 규칙 적용 범위의 첫 범위 왼쪽 위 칸을 기준으로 밀린다.
 *
 * 렌더링 때마다 모든 칸을 다시 계산하면 느리므로 칸별 결과를 캐시한다. 조건부서식 결과는 선택/스크롤과
 * 무관하고 시트 데이터(rows)와 규칙에만 의존하므로, 호출부는 그 둘이 바뀔 때만 새로 만들면 된다.
 */
import { evaluateFormula, isTruthy, type EvalContext, type Scalar } from '../formula/evaluate'
import { parseFormula, type Node } from '../formula/parse'
import type { CfRange, CfRule, CfStyle, ConditionalFormat } from '../xlsx/conditionalFormat'
import type { SheetModel } from '../xlsx/types'

/** (행,열)을 한 숫자 키로 합칠 때 쓰는 열 수 상한(엑셀 최대 16384열보다 크게). */
const COL_KEY_STRIDE = 20_000

const CELL_IS_OPERATORS: Record<string, string> = {
  greaterThan: '>',
  lessThan: '<',
  equal: '=',
  notEqual: '<>',
  greaterThanOrEqual: '>=',
  lessThanOrEqual: '<=',
}

interface CompiledRule {
  rule: CfRule
  /** 칸 값이 아니라 "판정 수식"으로 변환된 구문 트리(cellIs도 비교식으로 바꿔 둔다). */
  ast: Node
  ranges: CfRange[]
  anchorRow: number
  anchorCol: number
}

function inRanges(ranges: CfRange[], row: number, col: number): boolean {
  return ranges.some((r) => row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1)
}

/** cellIs 규칙을 일반 판정 수식(AST)으로: `현재칸 <연산> 기준값`, between은 두 비교의 AND/OR. */
function cellIsAst(rule: CfRule, anchorRow: number, anchorCol: number): Node | null {
  const self: Node = { t: 'ref', row: anchorRow, col: anchorCol, rowAbs: false, colAbs: false }
  const operand = (i: number): Node | null => (rule.formulae[i] ? parseFormula(rule.formulae[i]) : null)
  const op = rule.operator ?? ''
  if (op === 'between' || op === 'notBetween') {
    const lo = operand(0)
    const hi = operand(1)
    if (!lo || !hi) return null
    return op === 'between'
      ? { t: 'call', name: 'AND', args: [{ t: 'bin', op: '>=', l: self, r: lo }, { t: 'bin', op: '<=', l: self, r: hi }] }
      : { t: 'call', name: 'OR', args: [{ t: 'bin', op: '<', l: self, r: lo }, { t: 'bin', op: '>', l: self, r: hi }] }
  }
  const symbol = CELL_IS_OPERATORS[op]
  const rhs = operand(0)
  return symbol && rhs ? { t: 'bin', op: symbol, l: self, r: rhs } : null
}

function compile(formats: ConditionalFormat[]): CompiledRule[] {
  const compiled: CompiledRule[] = []
  for (const format of formats) {
    const first = format.ranges[0]
    for (const rule of format.rules) {
      try {
        const ast = rule.kind === 'cellIs' ? cellIsAst(rule, first.r0, first.c0) : parseFormula(rule.formulae[0])
        // 읽기 단계에서 이미 걸러졌어야 하지만, 해석이 안 되는 규칙이 섞여 와도 다른 규칙까지 깨지 않게 건너뛴다.
        if (ast) compiled.push({ rule, ast, ranges: format.ranges, anchorRow: first.r0, anchorCol: first.c0 })
      } catch {
        // 건너뜀
      }
    }
  }
  return compiled.sort((a, b) => a.rule.priority - b.rule.priority)
}

/** 같은 속성은 먼저 정해진(우선순위 높은) 값이 이긴다. */
function mergeInto(target: CfStyle, source: CfStyle): void {
  for (const key of Object.keys(source) as (keyof CfStyle)[]) {
    if (target[key] === undefined) (target as Record<string, unknown>)[key] = source[key]
  }
}

export type ConditionalStyler = (row: number, col: number) => CfStyle | null

/**
 * 시트의 조건부서식 평가기. 규칙이 하나도 없으면 null(호출부가 평가 비용을 아예 안 쓰게).
 * @param clock 테스트에서 오늘 날짜를 고정하는 용도
 */
export function createConditionalStyler(
  sheet: Pick<SheetModel, 'rows' | 'conditionalFormats'>,
  clock?: EvalContext['clock'],
): ConditionalStyler | null {
  const rules = compile(sheet.conditionalFormats)
  if (rules.length === 0) return null

  const getValue = (row: number, col: number): Scalar => sheet.rows[row - 1]?.[col - 1]?.value ?? null
  const cache = new Map<number, CfStyle | null>()

  return (row, col) => {
    const key = row * COL_KEY_STRIDE + col
    const hit = cache.get(key)
    if (hit !== undefined) return hit

    let result: CfStyle | null = null
    for (const c of rules) {
      if (!inRanges(c.ranges, row, col)) continue
      const value = evaluateFormula(c.ast, { row, col, anchorRow: c.anchorRow, anchorCol: c.anchorCol, getValue, clock })
      if (!isTruthy(value)) continue
      result ??= {}
      mergeInto(result, c.rule.style)
      if (c.rule.stopIfTrue) break
    }
    cache.set(key, result)
    return result
  }
}
