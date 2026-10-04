/**
 * 조건부서식 수식(`AND($L5="",$J5<TODAY(),$J5<>"")` 등)을 읽는 파서.
 *
 * 범위: 조건부서식에서 실제로 쓰이는 수식 부분집합 — 숫자/문자/불리언 리터럴, 셀 참조(절대/상대
 * `$A$1`), 범위(`A1:B5`), 사칙/거듭제곱/퍼센트/문자 연결/비교 연산자, 함수 호출. 다른 시트 참조,
 * 이름 정의, 배열 상수, 열 전체(`A:A`) 참조는 지원하지 않고 UnsupportedFormulaError를 던진다 —
 * 호출부는 그 규칙을 조용히 틀리게 적용하지 않고 "지원하지 않는 조건부서식"으로 세어 알린다.
 *
 * 연산자 우선순위는 Excel을 따른다(낮음 -> 높음): 비교 < & < + - < * / < ^ < 단항 +- < % .
 * 예) `-2^2`는 4, `2^3^2`는 (2^3)^2 = 64 (Excel은 거듭제곱이 왼쪽 결합).
 */

export class FormulaSyntaxError extends Error {}
export class UnsupportedFormulaError extends Error {}

export interface RefNode {
  t: 'ref'
  /** 1-based */
  col: number
  row: number
  colAbs: boolean
  rowAbs: boolean
}

export type Node =
  | { t: 'num'; v: number }
  | { t: 'str'; v: string }
  | { t: 'bool'; v: boolean }
  | RefNode
  | { t: 'range'; a: RefNode; b: RefNode }
  | { t: 'unary'; op: '-' | '+'; x: Node }
  | { t: 'percent'; x: Node }
  | { t: 'bin'; op: string; l: Node; r: Node }
  | { t: 'call'; name: string; args: Node[] }

type Token =
  | { k: 'num'; v: number }
  | { k: 'str'; v: string }
  | { k: 'ref'; v: RefNode }
  | { k: 'ident'; v: string }
  | { k: 'op'; v: string }

const REF_PATTERN = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/
const IDENT_START = /[A-Za-z_가-힣]/
const IDENT_PART = /[A-Za-z0-9_.가-힣]/
const NUMBER_PATTERN = /^(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/
const TWO_CHAR_OPS = ['<=', '>=', '<>']
const ONE_CHAR_OPS = '=<>&+-*/^%(),:'

function columnNumber(letters: string): number {
  let col = 0
  for (const ch of letters.toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64)
  return col
}

function toRef(text: string): RefNode | null {
  const m = text.match(REF_PATTERN)
  if (!m) return null
  return { t: 'ref', col: columnNumber(m[2]), row: Number(m[4]), colAbs: m[1] === '$', rowAbs: m[3] === '$' }
}

function tokenize(src: string): Token[] {
  const tokens: Token[] = []
  let i = 0
  while (i < src.length) {
    const c = src[i]
    if (/\s/.test(c)) {
      i++
    } else if (c === '"') {
      let s = ''
      i++
      for (;;) {
        if (i >= src.length) throw new FormulaSyntaxError('문자열이 닫히지 않았어요')
        if (src[i] === '"') {
          if (src[i + 1] === '"') {
            s += '"'
            i += 2
            continue
          }
          i++
          break
        }
        s += src[i++]
      }
      tokens.push({ k: 'str', v: s })
    } else if (c === "'") {
      throw new UnsupportedFormulaError('다른 시트 참조')
    } else if (/[0-9.]/.test(c) && NUMBER_PATTERN.test(src.slice(i))) {
      const m = src.slice(i).match(NUMBER_PATTERN)![0]
      tokens.push({ k: 'num', v: Number(m) })
      i += m.length
    } else if (TWO_CHAR_OPS.includes(src.slice(i, i + 2))) {
      tokens.push({ k: 'op', v: src.slice(i, i + 2) })
      i += 2
    } else if (c === '$' || IDENT_START.test(c)) {
      let j = i + 1
      while (j < src.length && (IDENT_PART.test(src[j]) || src[j] === '$')) j++
      const word = src.slice(i, j)
      if (src[j] === '!') throw new UnsupportedFormulaError('다른 시트 참조')
      const nextNonSpace = src.slice(j).match(/^\s*(.)/)?.[1]
      const ref = nextNonSpace === '(' ? null : toRef(word)
      if (ref) tokens.push({ k: 'ref', v: ref })
      else if (word.startsWith('$')) throw new FormulaSyntaxError(`이해할 수 없는 참조: ${word}`)
      else tokens.push({ k: 'ident', v: word })
      i = j
    } else if (ONE_CHAR_OPS.includes(c)) {
      tokens.push({ k: 'op', v: c })
      i++
    } else if (c === '{') {
      throw new UnsupportedFormulaError('배열 상수')
    } else {
      throw new FormulaSyntaxError(`이해할 수 없는 글자: ${c}`)
    }
  }
  return tokens
}

class Parser {
  private pos = 0
  private readonly tokens: Token[]

  constructor(tokens: Token[]) {
    this.tokens = tokens
  }

  parse(): Node {
    const node = this.comparison()
    if (this.pos < this.tokens.length) throw new FormulaSyntaxError('수식 끝에 남는 부분이 있어요')
    return node
  }

  private peekOp(...ops: string[]): string | null {
    const t = this.tokens[this.pos]
    return t && t.k === 'op' && ops.includes(t.v) ? t.v : null
  }

  private binary(next: () => Node, ...ops: string[]): Node {
    let left = next()
    for (let op = this.peekOp(...ops); op; op = this.peekOp(...ops)) {
      this.pos++
      left = { t: 'bin', op, l: left, r: next() }
    }
    return left
  }

  private comparison = (): Node => this.binary(this.concat, '=', '<>', '<', '>', '<=', '>=')
  private concat = (): Node => this.binary(this.additive, '&')
  private additive = (): Node => this.binary(this.multiplicative, '+', '-')
  private multiplicative = (): Node => this.binary(this.power, '*', '/')
  private power = (): Node => this.binary(this.unary, '^')

  private unary = (): Node => {
    const op = this.peekOp('-', '+')
    if (op) {
      this.pos++
      return { t: 'unary', op: op as '-' | '+', x: this.unary() }
    }
    return this.postfix()
  }

  private postfix(): Node {
    let node = this.primary()
    while (this.peekOp('%')) {
      this.pos++
      node = { t: 'percent', x: node }
    }
    return node
  }

  private primary(): Node {
    const t = this.tokens[this.pos++]
    if (!t) throw new FormulaSyntaxError('수식이 중간에 끝났어요')
    if (t.k === 'num') return { t: 'num', v: t.v }
    if (t.k === 'str') return { t: 'str', v: t.v }
    if (t.k === 'ref') {
      if (this.peekOp(':')) {
        this.pos++
        const b = this.tokens[this.pos++]
        if (!b || b.k !== 'ref') throw new UnsupportedFormulaError('열/행 전체 참조')
        return { t: 'range', a: t.v, b: b.v }
      }
      return t.v
    }
    if (t.k === 'op' && t.v === '(') {
      const inner = this.comparison()
      if (!this.peekOp(')')) throw new FormulaSyntaxError('괄호가 닫히지 않았어요')
      this.pos++
      return inner
    }
    if (t.k === 'ident') {
      const upper = t.v.toUpperCase()
      if (this.peekOp('(')) {
        this.pos++
        const args: Node[] = []
        if (this.peekOp(')')) {
          this.pos++
          return { t: 'call', name: upper, args }
        }
        for (;;) {
          // 빈 인자(`IF(A1,,2)`)는 지원하지 않는다.
          args.push(this.comparison())
          if (this.peekOp(',')) {
            this.pos++
            continue
          }
          if (this.peekOp(')')) {
            this.pos++
            break
          }
          throw new FormulaSyntaxError('함수 인자 구분이 이상해요')
        }
        return { t: 'call', name: upper, args }
      }
      if (upper === 'TRUE') return { t: 'bool', v: true }
      if (upper === 'FALSE') return { t: 'bool', v: false }
      throw new UnsupportedFormulaError(`이름 정의: ${t.v}`)
    }
    throw new FormulaSyntaxError(`예상하지 못한 기호: ${t.v}`)
  }
}

export function parseFormula(src: string): Node {
  return new Parser(tokenize(src.trim().replace(/^=/, ''))).parse()
}

/** 수식 안에서 호출하는 함수 이름(대문자) 전부. */
export function collectFunctionNames(node: Node, out = new Set<string>()): Set<string> {
  switch (node.t) {
    case 'call':
      out.add(node.name)
      node.args.forEach((a) => collectFunctionNames(a, out))
      break
    case 'bin':
      collectFunctionNames(node.l, out)
      collectFunctionNames(node.r, out)
      break
    case 'unary':
    case 'percent':
      collectFunctionNames(node.x, out)
      break
  }
  return out
}
