import { describe, expect, it } from 'vitest'
import { evaluateFormula, isTruthy, SUPPORTED_FUNCTIONS, type EvalContext, type Scalar } from './evaluate'
import { collectFunctionNames, FormulaSyntaxError, parseFormula, UnsupportedFormulaError } from './parse'

/** 2026-10-04 (직렬번호) — 테스트에서 시계를 고정한다. */
const TODAY = 46299

function ctxFor(cells: Record<string, Scalar>, over: Partial<EvalContext> = {}): EvalContext {
  const key = (r: number, c: number) => `${String.fromCharCode(64 + c)}${r}`
  return { row: 1, col: 1, anchorRow: 1, anchorCol: 1, getValue: (r, c) => cells[key(r, c)] ?? null, clock: { today: TODAY, now: TODAY + 0.5 }, ...over }
}
const run = (formula: string, cells: Record<string, Scalar> = {}, over: Partial<EvalContext> = {}) => evaluateFormula(parseFormula(formula), ctxFor(cells, over))

describe('parseFormula', () => {
  it('연산자 우선순위는 Excel과 같다', () => {
    expect(run('1+2*3')).toBe(7)
    expect(run('-2^2')).toBe(4) // Excel은 단항 마이너스가 거듭제곱보다 먼저
    expect(run('2^3^2')).toBe(64) // 왼쪽 결합
    expect(run('50%')).toBe(0.5)
    expect(run('"a"&"b"="ab"')).toBe(true) // & 가 비교보다 먼저
  })

  it('문자열 안의 따옴표 두 개는 따옴표 하나', () => {
    expect(run('"a""b"')).toBe('a"b')
  })

  it('지원하지 않는 형태는 조용히 틀리게 계산하지 않고 알려준다', () => {
    expect(() => parseFormula("Sheet2!A1=1")).toThrow(UnsupportedFormulaError)
    expect(() => parseFormula("'My Sheet'!A1")).toThrow(UnsupportedFormulaError)
    expect(() => parseFormula('MyName>1')).toThrow(UnsupportedFormulaError)
    expect(() => parseFormula('SUM(A:A)')).toThrow(UnsupportedFormulaError)
    expect(() => parseFormula('{1,2}')).toThrow(UnsupportedFormulaError)
    expect(() => parseFormula('1+')).toThrow(FormulaSyntaxError)
    expect(() => parseFormula('(1+2')).toThrow(FormulaSyntaxError)
  })

  it('함수 이름을 모은다', () => {
    expect([...collectFunctionNames(parseFormula('AND(A1="",TODAY()>B1,LEN(TRIM(C1))>0)'))].sort()).toEqual(['AND', 'LEN', 'TODAY', 'TRIM'])
  })
})

describe('상대 참조', () => {
  it('기준 칸에서 평가 칸까지 거리만큼 상대 참조만 밀린다 — 행 전체 색칠의 핵심', () => {
    // 기준 칸 A1에 쓴 `$B1=1` 을 5행에서 평가하면 $B5
    const cells = { B1: 0, B5: 1 }
    expect(run('$B1=1', cells, { row: 5, col: 1 })).toBe(true)
    expect(run('$B1=1', cells, { row: 1, col: 1 })).toBe(false)
  })

  it('절대 참조는 밀리지 않고, 열 상대는 열 거리만큼 밀린다', () => {
    const cells = { A1: 7, C3: 9 }
    expect(run('$A$1', cells, { row: 3, col: 3 })).toBe(7)
    expect(run('A1', cells, { row: 3, col: 3 })).toBe(9) // 기준(1,1)에서 (2,2)만큼 → C3
  })
})

describe('Excel 비교 규칙', () => {
  it('빈 칸은 ""·0·FALSE처럼 비교된다', () => {
    expect(run('A1=""', {})).toBe(true)
    expect(run('A1=0', {})).toBe(true)
    expect(run('A1<>""', { A1: '가' })).toBe(true)
  })
  it('종류가 다르면 숫자 < 문자 < 불리언, 같다고 하지 않는다', () => {
    expect(run('1<"a"')).toBe(true)
    expect(run('"a"<TRUE')).toBe(true)
    expect(run('TRUE=1')).toBe(false)
  })
  it('문자 비교는 대소문자를 구분하지 않는다', () => {
    expect(run('"ABC"="abc"')).toBe(true)
  })
  it('오류는 전파되고 규칙은 적용되지 않는다', () => {
    expect(run('1/0')).toEqual({ error: '#DIV/0!' })
    expect(isTruthy(run('1/0=1'))).toBe(false)
    expect(isTruthy(run('"abc"'))).toBe(false)
    expect(isTruthy(run('2'))).toBe(true)
  })
})

describe('함수', () => {
  it('AND/OR/NOT/IF/IFERROR', () => {
    expect(run('AND(1=1,2=2)')).toBe(true)
    expect(run('AND(1=1,2=3)')).toBe(false)
    expect(run('OR(1=2,2=2)')).toBe(true)
    expect(run('NOT(1=2)')).toBe(true)
    expect(run('IF(1=1,"예","아니오")')).toBe('예')
    expect(run('IFERROR(1/0,"오류")')).toBe('오류')
  })
  it('TODAY와 날짜 비교 — 기한 지난 행 판정', () => {
    expect(run('A1<TODAY()', { A1: TODAY - 1 })).toBe(true)
    expect(run('A1<TODAY()', { A1: TODAY + 1 })).toBe(false)
    expect(run('YEAR(TODAY())')).toBe(2026)
    expect(run('MONTH(TODAY())')).toBe(10)
    expect(run('DAY(TODAY())')).toBe(4)
    expect(run('WEEKDAY(TODAY())')).toBe(1) // 일요일
    expect(run('WEEKDAY(TODAY(),2)')).toBe(7)
  })
  it('문자 함수', () => {
    expect(run('LEN(TRIM("  a   b "))')).toBe(3)
    expect(run('LEFT("abcdef",2)&RIGHT("abcdef",2)')).toBe('abef')
    expect(run('MID("abcdef",2,3)')).toBe('bcd')
    expect(run('UPPER("aB")&LOWER("aB")')).toBe('ABab')
    expect(run('SEARCH("B","abc")')).toBe(2)
    expect(run('ISERROR(SEARCH("z","abc"))')).toBe(true)
    expect(run('FIND("B","abc")')).toEqual({ error: '#VALUE!' })
    expect(run('NOT(ISERROR(SEARCH("수정",A1)))', { A1: '문구 수정 요청' })).toBe(true)
  })
  it('정보 함수', () => {
    expect(run('ISBLANK(A1)', {})).toBe(true)
    expect(run('ISBLANK(A1)', { A1: '' })).toBe(false)
    expect(run('ISNUMBER(A1)', { A1: 3 })).toBe(true)
    expect(run('ISTEXT(A1)', { A1: '가' })).toBe(true)
    expect(run('ISTEXT(A1)', { A1: 3 })).toBe(false)
  })
  it('집계와 COUNTIF', () => {
    const cells = { A1: 1, A2: 2, A3: 3, A4: '가', A5: '가' }
    expect(run('SUM(A1:A3)', cells)).toBe(6)
    expect(run('MAX(A1:A3)', cells)).toBe(3)
    expect(run('MIN(A1:A3)', cells)).toBe(1)
    expect(run('AVERAGE(A1:A3)', cells)).toBe(2)
    expect(run('COUNT(A1:A5)', cells)).toBe(3)
    expect(run('COUNTA(A1:A5)', cells)).toBe(5)
    expect(run('COUNTIF(A1:A5,"가")', cells)).toBe(2)
    expect(run('COUNTIF(A1:A5,">1")', cells)).toBe(2)
    expect(run('COUNTIF(A1:A5,"<>가")', cells)).toBe(3)
    expect(run('COUNTIF(A1:A5,A4)>1', cells)).toBe(true) // 중복 값 강조 패턴
  })
  it('수학/행열', () => {
    expect(run('MOD(-1,3)')).toBe(2)
    expect(run('MOD(5,0)')).toEqual({ error: '#DIV/0!' })
    expect(run('ROUND(2.675,2)')).toBeCloseTo(2.68, 5)
    expect(run('ABS(-3)+INT(2.9)')).toBe(5)
    expect(run('MOD(ROW(),2)=0', {}, { row: 4 })).toBe(true) // 줄무늬 패턴
    expect(run('COLUMN()', {}, { col: 3 })).toBe(3)
  })
  it('DATE', () => {
    expect(run('DATE(2026,10,4)')).toBe(TODAY)
  })
  it('SUPPORTED_FUNCTIONS는 구현된 이름과 일치', () => {
    for (const name of ['AND', 'OR', 'TODAY', 'COUNTIF', 'SEARCH']) expect(SUPPORTED_FUNCTIONS.has(name)).toBe(true)
    expect(SUPPORTED_FUNCTIONS.has('VLOOKUP')).toBe(false)
  })
})

describe('이 파일의 실제 조건부서식 수식', () => {
  // 현황 시트: J=기한(날짜), K=상태 코드, L=처리 표시. 기준 칸은 A1이고 5행에서 평가.
  const at5 = (formula: string, cells: Record<string, Scalar>) => run(formula, cells, { row: 5, col: 3 })

  it('기한 지난 미처리 행(빨간 글자)', () => {
    const f = 'AND($L1="",$J1<TODAY(),$J1<>"")'
    expect(at5(f, { J5: TODAY - 3 })).toBe(true) // 처리 표시 없음 + 기한 지남
    expect(at5(f, { J5: TODAY - 3, L5: 'O' })).toBe(false) // 처리 완료
    expect(at5(f, { J5: TODAY + 3 })).toBe(false) // 아직 기한 안 지남
    expect(at5(f, {})).toBe(false) // 기한이 비어 있으면 (빈 칸=0이라 < TODAY 지만) 제외
  })
  it('불가/보류 행(취소선 + 연한 채우기)', () => {
    const f = 'OR($L1="불가",$L1="보류")'
    expect(at5(f, { L5: '불가' })).toBe(true)
    expect(at5(f, { L5: '보류' })).toBe(true)
    expect(at5(f, { L5: 'O' })).toBe(false)
  })
  it('상태 코드 일치 행($K1=1, $L1="O")', () => {
    expect(at5('$K1=1', { K5: 1 })).toBe(true)
    expect(at5('$K1=1', { K5: 2 })).toBe(false)
    expect(at5('$L1="O"', { L5: 'o' })).toBe(true) // 대소문자 무시
  })
})
