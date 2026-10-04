/**
 * 조건부서식 수식 평가기. parse.ts가 만든 구문 트리를 한 칸 기준으로 계산해 값을 돌려준다.
 *
 * Excel 의미를 따르는 부분(틀리면 색이 엉뚱한 행에 칠해지므로 중요하다):
 *  - 상대 참조는 "수식이 쓰인 기준 칸(범위 첫 칸)"에서 평가 대상 칸까지의 거리만큼 밀린다.
 *  - 빈 칸은 숫자와 비교하면 0, 문자와 비교하면 "", 불리언과 비교하면 FALSE로 본다.
 *  - 서로 다른 종류를 비교하면 숫자 < 문자 < 불리언 순서이고, 문자 비교는 대소문자를 구분하지 않는다.
 *  - 오류(#DIV/0! 등)는 값처럼 전파되고, 최종 결과가 TRUE(또는 0이 아닌 숫자)일 때만 규칙이 적용된다.
 *
 * 지원 함수는 SUPPORTED_FUNCTIONS 한 곳에서 정한다 — 읽는 쪽(conditionalFormat.ts)이 이 목록으로
 * "표시 못 하는 규칙"을 미리 걸러내므로, 함수를 추가하려면 여기에 구현과 이름을 같이 넣으면 된다.
 *
 * 워크북 문맥(다른 시트, 파일 이름)이 있어야 의미가 있는 함수(INDIRECT, CELL, HYPERLINK)는 CONTEXT_FUNCTIONS로 따로 둔다.
 * 링크 이동 대상 계산(linkTarget.ts)은 쓰지만 조건부서식은 그 문맥을 주지 않으므로 SUPPORTED_FUNCTIONS에서는 빠진다 — 안 그러면
 * 조건부서식 규칙이 조용히 오류가 되어 "표시 못 하는 서식" 안내가 사라진다.
 */
import { nowSerial, serialToUtcDate, todaySerial } from '../xlsx/excelDate'
import { generalNumber } from '../xlsx/numberFormat'
import type { Node, RefNode } from './parse'

export type Scalar = number | string | boolean | null
export interface FormulaError {
  error: string
}
export type Value = Scalar | FormulaError

interface RangeValue {
  range: true
  r0: number
  c0: number
  r1: number
  c1: number
  /** 다른 시트의 범위일 때 그 시트 이름(INDIRECT). 없으면 지금 시트. */
  sheet?: string
}
type Result = Value | RangeValue

export interface EvalContext {
  /** 평가 대상 칸(1-based) */
  row: number
  col: number
  /** 수식이 쓰인 기준 칸 — 상대 참조가 여기서부터 밀린다. */
  anchorRow: number
  anchorCol: number
  getValue: (row: number, col: number) => Scalar
  /** 테스트에서 시간을 고정할 수 있게 주입한다. 없으면 실제 시계. */
  clock?: { today: number; now: number }
  /** 지금 시트 이름과 파일 이름 — CELL("filename")이 쓴다. */
  sheetName?: string
  fileName?: string
  /** 다른 시트의 칸 값(INDIRECT). 시트가 없으면 undefined. */
  sheetValue?: (sheet: string, row: number, col: number) => Scalar
  /** 시트의 사용된 행/열 수 — 열 전체 참조(`F:F`)를 실제 데이터 범위로 줄이는 데 쓴다. 시트가 없으면 null. */
  sheetExtent?: (sheet: string) => { rows: number; cols: number } | null
}

const ERR = {
  value: { error: '#VALUE!' },
  div0: { error: '#DIV/0!' },
  ref: { error: '#REF!' },
  na: { error: '#N/A' },
  num: { error: '#NUM!' },
} as const

const ERROR_CODES = new Set(['#VALUE!', '#DIV/0!', '#REF!', '#N/A', '#NUM!', '#NAME?', '#NULL!'])

const isError = (v: unknown): v is FormulaError => typeof v === 'object' && v !== null && 'error' in v
const isRange = (v: unknown): v is RangeValue => typeof v === 'object' && v !== null && 'range' in v

/** 셀에 문자열로 저장된 오류값("#N/A")도 오류로 본다. */
const asError = (v: Value): FormulaError | null => (isError(v) ? v : typeof v === 'string' && ERROR_CODES.has(v) ? { error: v } : null)

function toNumber(v: Value): number | FormulaError {
  if (isError(v)) return v
  if (typeof v === 'number') return v
  if (typeof v === 'boolean') return v ? 1 : 0
  if (v === null) return 0
  const t = v.trim()
  if (t === '') return ERR.value
  const n = Number(t)
  return Number.isNaN(n) ? ERR.value : n
}

function toText(v: Value): string | FormulaError {
  if (isError(v)) return v
  if (v === null) return ''
  if (typeof v === 'number') return generalNumber(v)
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
  return v
}

function toBool(v: Value): boolean | FormulaError {
  if (isError(v)) return v
  if (typeof v === 'boolean') return v
  if (typeof v === 'number') return v !== 0
  if (v === null) return false
  const u = v.trim().toUpperCase()
  return u === 'TRUE' ? true : u === 'FALSE' ? false : ERR.value
}

/** Excel 비교: -1/0/1. 빈 칸은 상대 종류에 맞춰 0/""/FALSE로 바꿔 비교한다. */
function compare(a: Scalar, b: Scalar): number {
  const norm = (x: Scalar, other: Scalar): number | string | boolean =>
    x !== null ? x : typeof other === 'string' ? '' : typeof other === 'boolean' ? false : 0
  const x = norm(a, b)
  const y = norm(b, a)
  const rank = (v: number | string | boolean) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2)
  if (rank(x) !== rank(y)) return rank(x) < rank(y) ? -1 : 1
  if (typeof x === 'string' && typeof y === 'string') {
    const lx = x.toLowerCase()
    const ly = y.toLowerCase()
    return lx === ly ? 0 : lx < ly ? -1 : 1
  }
  const nx = Number(x)
  const ny = Number(y)
  return nx === ny ? 0 : nx < ny ? -1 : 1
}

function resolveRef(ref: RefNode, ctx: EvalContext): { row: number; col: number } | null {
  const row = ref.rowAbs ? ref.row : ref.row + (ctx.row - ctx.anchorRow)
  const col = ref.colAbs ? ref.col : ref.col + (ctx.col - ctx.anchorCol)
  return row < 1 || col < 1 ? null : { row, col }
}

/** 범위 인자를 값 목록으로 펼친다(범위가 아니면 값 하나). 같은 범위 인자들은 크기가 작아 그대로 펼친다. */
function flatten(arg: Result, ctx: EvalContext): Value[] {
  if (!isRange(arg)) return [arg]
  const read = arg.sheet !== undefined ? (r: number, c: number) => ctx.sheetValue?.(arg.sheet!, r, c) ?? null : ctx.getValue
  const out: Value[] = []
  for (let r = arg.r0; r <= arg.r1; r++) for (let c = arg.c0; c <= arg.c1; c++) out.push(read(r, c))
  return out
}

/** COUNTIF 조건: `">5"`, `"<>x"`, `"abc*"`(와일드카드), 숫자, 문자. */
function makeCriteria(criteria: Scalar): (v: Scalar) => boolean {
  if (typeof criteria !== 'string') return (v) => v !== null && compare(v, criteria) === 0 && typeof v === typeof criteria
  const m = criteria.match(/^(<=|>=|<>|=|<|>)?(.*)$/s)!
  const op = m[1] ?? '='
  const operandText = m[2]
  const asNumber = operandText.trim() !== '' && !Number.isNaN(Number(operandText)) ? Number(operandText) : null
  return (v) => {
    if (asNumber !== null) {
      if (typeof v !== 'number') return op === '<>'
      const c = compare(v, asNumber)
      return op === '=' ? c === 0 : op === '<>' ? c !== 0 : op === '<' ? c < 0 : op === '>' ? c > 0 : op === '<=' ? c <= 0 : c >= 0
    }
    const text = v === null ? '' : typeof v === 'string' ? v : String(v)
    if (op === '=' || op === '<>') {
      const pattern = new RegExp(`^${operandText.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.')}$`, 'i')
      return pattern.test(text) === (op === '=')
    }
    const c = compare(text, operandText)
    return op === '<' ? c < 0 : op === '>' ? c > 0 : op === '<=' ? c <= 0 : c >= 0
  }
}

type Fn = (args: Node[], ctx: EvalContext, ev: (n: Node) => Result) => Result

/** 인자를 값으로 평가(범위는 #VALUE!). */
const scalarArg = (n: Node, ev: (n: Node) => Result): Value => {
  const r = ev(n)
  return isRange(r) ? ERR.value : r
}

const FUNCTIONS: Record<string, Fn> = {
  // 논리
  AND: (args, ctx, ev) => logical(args, ctx, ev, true),
  OR: (args, ctx, ev) => logical(args, ctx, ev, false),
  NOT: (args, _c, ev) => {
    const b = toBool(scalarArg(args[0], ev))
    return isError(b) ? b : !b
  },
  IF: (args, _c, ev) => {
    const cond = toBool(scalarArg(args[0], ev))
    if (isError(cond)) return cond
    if (cond) return args[1] ? scalarArg(args[1], ev) : true
    return args[2] ? scalarArg(args[2], ev) : false
  },
  IFERROR: (args, _c, ev) => {
    const v = scalarArg(args[0], ev)
    return asError(v) ? scalarArg(args[1], ev) : v
  },
  // 정보
  ISBLANK: (args, _c, ev) => scalarArg(args[0], ev) === null,
  ISNUMBER: (args, _c, ev) => typeof scalarArg(args[0], ev) === 'number',
  ISTEXT: (args, _c, ev) => {
    const v = scalarArg(args[0], ev)
    return typeof v === 'string' && !asError(v)
  },
  ISERROR: (args, _c, ev) => asError(scalarArg(args[0], ev)) !== null,
  // 문자
  LEN: (args, _c, ev) => {
    const t = toText(scalarArg(args[0], ev))
    return isError(t) ? t : t.length
  },
  TRIM: (args, _c, ev) => {
    const t = toText(scalarArg(args[0], ev))
    return isError(t) ? t : t.trim().replace(/ +/g, ' ')
  },
  UPPER: (args, _c, ev) => mapText(args, ev, (t) => t.toUpperCase()),
  LOWER: (args, _c, ev) => mapText(args, ev, (t) => t.toLowerCase()),
  LEFT: (args, _c, ev) => slice(args, ev, (t, n) => t.slice(0, n), 1),
  RIGHT: (args, _c, ev) => slice(args, ev, (t, n) => (n === 0 ? '' : t.slice(-n)), 1),
  MID: (args, _c, ev) => {
    const t = toText(scalarArg(args[0], ev))
    const start = toNumber(scalarArg(args[1], ev))
    const len = toNumber(scalarArg(args[2], ev))
    if (isError(t)) return t
    if (isError(start)) return start
    if (isError(len)) return len
    return start < 1 || len < 0 ? ERR.value : t.slice(start - 1, start - 1 + len)
  },
  SEARCH: (args, _c, ev) => find(args, ev, false),
  FIND: (args, _c, ev) => find(args, ev, true),
  SUBSTITUTE: (args, _c, ev) => {
    const text = toText(scalarArg(args[0], ev))
    const from = toText(scalarArg(args[1], ev))
    const to = toText(scalarArg(args[2], ev))
    const nth = args[3] ? toNumber(scalarArg(args[3], ev)) : 0
    if (isError(text)) return text
    if (isError(from)) return from
    if (isError(to)) return to
    if (isError(nth)) return nth
    if (from === '') return text
    if (nth === 0) return text.split(from).join(to)
    // n번째 것만 바꾼다
    let at = -1
    for (let i = 0; i < nth; i++) {
      at = text.indexOf(from, at + 1)
      if (at < 0) return text
    }
    return text.slice(0, at) + to + text.slice(at + from.length)
  },
  EXACT: (args, _c, ev) => {
    const a = toText(scalarArg(args[0], ev))
    const b = toText(scalarArg(args[1], ev))
    return isError(a) ? a : isError(b) ? b : a === b
  },
  // 수학
  ABS: (args, _c, ev) => mapNumber(args, ev, Math.abs),
  INT: (args, _c, ev) => mapNumber(args, ev, Math.floor),
  ROUND: (args, _c, ev) => {
    const n = toNumber(scalarArg(args[0], ev))
    const d = toNumber(scalarArg(args[1], ev))
    if (isError(n)) return n
    if (isError(d)) return d
    const f = 10 ** d
    return (Math.sign(n) * Math.round(Math.abs(n) * f)) / f
  },
  MOD: (args, _c, ev) => {
    const n = toNumber(scalarArg(args[0], ev))
    const d = toNumber(scalarArg(args[1], ev))
    if (isError(n)) return n
    if (isError(d)) return d
    return d === 0 ? ERR.div0 : n - d * Math.floor(n / d)
  },
  // 집계
  SUM: (args, ctx, ev) => aggregate(args, ctx, ev, (nums) => nums.reduce((a, b) => a + b, 0)),
  MAX: (args, ctx, ev) => aggregate(args, ctx, ev, (nums) => (nums.length ? Math.max(...nums) : 0)),
  MIN: (args, ctx, ev) => aggregate(args, ctx, ev, (nums) => (nums.length ? Math.min(...nums) : 0)),
  AVERAGE: (args, ctx, ev) => {
    const nums = numbersOf(args, ctx, ev)
    return isError(nums) ? nums : nums.length === 0 ? ERR.div0 : nums.reduce((a, b) => a + b, 0) / nums.length
  },
  COUNT: (args, ctx, ev) => args.reduce((n, a) => n + flatten(ev(a), ctx).filter((v) => typeof v === 'number').length, 0),
  COUNTA: (args, ctx, ev) => args.reduce((n, a) => n + flatten(ev(a), ctx).filter((v) => v !== null).length, 0),
  COUNTIF: (args, ctx, ev) => {
    const range = flatten(ev(args[0]), ctx)
    const criteria = scalarArg(args[1], ev)
    if (isError(criteria)) return criteria
    const test = makeCriteria(criteria)
    return range.filter((v) => !isError(v) && test(v)).length
  },
  // 조회: 정확히 일치(match_type 0)만. 정렬된 목록에서 가까운 값을 찾는 근사 일치는 지원하지 않는다(#N/A).
  MATCH: (args, ctx, ev) => {
    const needle = scalarArg(args[0], ev)
    const haystack = ev(args[1])
    const type = args[2] ? toNumber(scalarArg(args[2], ev)) : 1
    if (isError(needle)) return needle
    if (isError(type)) return type
    if (type !== 0) return ERR.na
    // 찾을 범위가 오류면(예: 없는 시트를 가리킨 INDIRECT의 #REF!) 그 오류를 그대로 알린다 — #N/A로 바꾸면 원인이 가려진다.
    if (isError(haystack)) return haystack
    const index = flatten(haystack, ctx).findIndex((v) => !isError(v) && v !== null && compare(v, needle) === 0)
    return index < 0 ? ERR.na : index + 1
  },
  // 행/열
  ROW: (args, ctx) => {
    if (args.length === 0) return ctx.row
    const node = args[0]
    const ref = node.t === 'ref' ? node : node.t === 'range' ? node.a : null
    const at = ref ? resolveRef(ref, ctx) : null
    return at ? at.row : ERR.ref
  },
  COLUMN: (args, ctx) => {
    if (args.length === 0) return ctx.col
    const node = args[0]
    const ref = node.t === 'ref' ? node : node.t === 'range' ? node.a : null
    const at = ref ? resolveRef(ref, ctx) : null
    return at ? at.col : ERR.ref
  },
  // 날짜
  TODAY: (_a, ctx) => ctx.clock?.today ?? todaySerial(),
  NOW: (_a, ctx) => ctx.clock?.now ?? nowSerial(),
  YEAR: (args, _c, ev) => datePart(args, ev, (d) => d.getUTCFullYear()),
  MONTH: (args, _c, ev) => datePart(args, ev, (d) => d.getUTCMonth() + 1),
  DAY: (args, _c, ev) => datePart(args, ev, (d) => d.getUTCDate()),
  WEEKDAY: (args, _c, ev) => {
    const type = args[1] ? toNumber(scalarArg(args[1], ev)) : 1
    if (isError(type)) return type
    return datePart(args, ev, (d) => {
      const sundayFirst = d.getUTCDay() // 0=일
      if (type === 2) return sundayFirst === 0 ? 7 : sundayFirst // 월=1..일=7
      if (type === 3) return (sundayFirst + 6) % 7 // 월=0..일=6
      return sundayFirst + 1 // 일=1..토=7
    })
  },
  DATE: (args, _c, ev) => {
    const parts = args.map((a) => toNumber(scalarArg(a, ev)))
    const bad = parts.find(isError)
    if (bad) return bad as FormulaError
    const [y, m, d] = parts as number[]
    return (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86_400_000
  },
}

function logical(args: Node[], ctx: EvalContext, ev: (n: Node) => Result, isAnd: boolean): Value {
  let seen = false
  let acc = isAnd
  for (const a of args) {
    const r = ev(a)
    for (const v of flatten(r, ctx)) {
      if (isError(v)) return v
      // 범위/참조 안의 문자·빈 칸은 무시한다(직접 쓴 문자열 인자는 위에서 #VALUE!가 되도록 toBool에 맡긴다).
      if (isRange(r) && (typeof v === 'string' || v === null)) continue
      if (v === null) continue
      const b = toBool(v)
      if (isError(b)) return b
      seen = true
      acc = isAnd ? acc && b : acc || b
    }
  }
  return seen ? acc : ERR.value
}

function mapText(args: Node[], ev: (n: Node) => Result, f: (t: string) => string): Value {
  const t = toText(scalarArg(args[0], ev))
  return isError(t) ? t : f(t)
}

function mapNumber(args: Node[], ev: (n: Node) => Result, f: (n: number) => number): Value {
  const n = toNumber(scalarArg(args[0], ev))
  return isError(n) ? n : f(n)
}

function slice(args: Node[], ev: (n: Node) => Result, f: (t: string, n: number) => string, defaultCount: number): Value {
  const t = toText(scalarArg(args[0], ev))
  const n = args[1] ? toNumber(scalarArg(args[1], ev)) : defaultCount
  if (isError(t)) return t
  if (isError(n)) return n
  return n < 0 ? ERR.value : f(t, n)
}

function find(args: Node[], ev: (n: Node) => Result, caseSensitive: boolean): Value {
  const needle = toText(scalarArg(args[0], ev))
  const hay = toText(scalarArg(args[1], ev))
  const start = args[2] ? toNumber(scalarArg(args[2], ev)) : 1
  if (isError(needle)) return needle
  if (isError(hay)) return hay
  if (isError(start)) return start
  const h = caseSensitive ? hay : hay.toLowerCase()
  const n = caseSensitive ? needle : needle.toLowerCase()
  const i = h.indexOf(n, Math.max(0, start - 1))
  return i < 0 ? ERR.value : i + 1
}

function datePart(args: Node[], ev: (n: Node) => Result, f: (d: Date) => number): Value {
  const n = toNumber(scalarArg(args[0], ev))
  return isError(n) ? n : f(serialToUtcDate(n))
}

function numbersOf(args: Node[], ctx: EvalContext, ev: (n: Node) => Result): number[] | FormulaError {
  const out: number[] = []
  for (const a of args) {
    const r = ev(a)
    for (const v of flatten(r, ctx)) {
      if (isError(v)) return v
      // 범위 안의 숫자만 센다. 직접 쓴 인자는 숫자로 바꿀 수 있으면 센다.
      if (isRange(r)) {
        if (typeof v === 'number') out.push(v)
      } else {
        const n = toNumber(v)
        if (isError(n)) return n
        out.push(n)
      }
    }
  }
  return out
}

function aggregate(args: Node[], ctx: EvalContext, ev: (n: Node) => Result, f: (nums: number[]) => number): Value {
  const nums = numbersOf(args, ctx, ev)
  return isError(nums) ? nums : f(nums)
}

/**
 * 워크북 문맥이 있어야 의미가 있는 함수. 링크 이동 대상 계산은 문맥(다른 시트, 파일 이름)을 주지만 조건부서식은 주지 않으므로,
 * 조건부서식에서는 지원하지 않는 함수로 취급한다(SUPPORTED_FUNCTIONS 참고).
 */
const CONTEXT_FUNCTIONS: Record<string, Fn> = {
  /** 글자로 쓴 참조(`'시트'!F:F`, `A1:B5`)를 범위로. 열 전체는 시트의 실제 사용 범위까지로 줄인다. */
  INDIRECT: (args, ctx, ev) => {
    const text = toText(scalarArg(args[0], ev))
    return isError(text) ? text : refFromText(text, ctx)
  },
  /** CELL("filename")만: `[파일이름]시트이름`(Excel은 앞에 폴더 경로가 붙는다 — 경로는 모르므로 비운다). */
  CELL: (args, ctx, ev) => {
    const kind = toText(scalarArg(args[0], ev))
    if (isError(kind)) return kind
    return kind.toLowerCase() === 'filename' && ctx.fileName !== undefined && ctx.sheetName !== undefined ? `[${ctx.fileName}]${ctx.sheetName}` : ERR.value
  },
  /** 칸에 보이는 글자는 두 번째 인자(없으면 대상 글자). 이동 대상은 linkTarget.ts가 첫 인자로 따로 계산한다. */
  HYPERLINK: (args, _c, ev) => (args[1] ? scalarArg(args[1], ev) : scalarArg(args[0], ev)),
}

const ALL_FUNCTIONS: Record<string, Fn> = { ...FUNCTIONS, ...CONTEXT_FUNCTIONS }

/** 구현된 함수 이름(대문자) 중 조건부서식이 쓸 수 있는 것 — 읽는 쪽(conditionalFormat.ts)이 지원 여부를 판단하는 기준. */
export const SUPPORTED_FUNCTIONS: ReadonlySet<string> = new Set(Object.keys(FUNCTIONS))

const SHEET_REF_RE = /^(?:('(?:[^']|'')+'|[^'!]+)!)?(.+)$/
const COLUMN_SPAN_RE = /^\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})$/
const CELL_SPAN_RE = /^\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/

function columnIndex(letters: string): number {
  let col = 0
  for (const ch of letters.toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64)
  return col
}

/** INDIRECT의 글자 참조를 범위로. 모르는 시트나 이해 못 하는 모양은 #REF!. */
function refFromText(text: string, ctx: EvalContext): Result {
  const m = SHEET_REF_RE.exec(text.trim())
  if (!m) return ERR.ref
  const sheet = m[1] === undefined ? undefined : m[1].startsWith("'") ? m[1].slice(1, -1).replace(/''/g, "'") : m[1]
  const target = sheet ?? ctx.sheetName
  // 다른 시트를 가리키는데 그 시트를 읽을 방법이 없거나 그런 시트가 없으면 틀린 값을 만들지 않고 #REF!.
  if (sheet !== undefined && (!ctx.sheetValue || (ctx.sheetExtent && !ctx.sheetExtent(sheet)))) return ERR.ref
  const ref = m[2]

  const columns = COLUMN_SPAN_RE.exec(ref)
  if (columns) {
    const extent = target !== undefined ? ctx.sheetExtent?.(target) : null
    if (!extent) return ERR.ref
    const [a, b] = [columnIndex(columns[1]), columnIndex(columns[2])]
    return { range: true, r0: 1, c0: Math.min(a, b), r1: Math.max(1, extent.rows), c1: Math.max(a, b), sheet }
  }
  const cells = CELL_SPAN_RE.exec(ref)
  if (!cells) return ERR.ref
  const c0 = columnIndex(cells[1])
  const r0 = Number(cells[2])
  const c1 = cells[3] ? columnIndex(cells[3]) : c0
  const r1 = cells[4] ? Number(cells[4]) : r0
  // 칸 하나를 가리키면 범위가 아니라 그 칸의 값이다(Excel도 한 칸 참조는 값처럼 쓰인다).
  if (r0 === r1 && c0 === c1) return sheet !== undefined ? (ctx.sheetValue?.(sheet, r0, c0) ?? null) : ctx.getValue(r0, c0)
  return { range: true, r0: Math.min(r0, r1), c0: Math.min(c0, c1), r1: Math.max(r0, r1), c1: Math.max(c0, c1), sheet }
}

export function evaluateFormula(ast: Node, ctx: EvalContext): Value {
  const ev = (n: Node): Result => {
    switch (n.t) {
      case 'num':
      case 'str':
      case 'bool':
        return n.v
      case 'ref': {
        const at = resolveRef(n, ctx)
        return at ? ctx.getValue(at.row, at.col) : ERR.ref
      }
      case 'range': {
        const a = resolveRef(n.a, ctx)
        const b = resolveRef(n.b, ctx)
        if (!a || !b) return ERR.ref
        return { range: true, r0: Math.min(a.row, b.row), c0: Math.min(a.col, b.col), r1: Math.max(a.row, b.row), c1: Math.max(a.col, b.col) }
      }
      case 'unary': {
        const x = toNumber(scalarArg(n.x, ev))
        return isError(x) ? x : n.op === '-' ? -x : x
      }
      case 'percent': {
        const x = toNumber(scalarArg(n.x, ev))
        return isError(x) ? x : x / 100
      }
      case 'call': {
        const fn = ALL_FUNCTIONS[n.name]
        return fn ? fn(n.args, ctx, ev) : { error: '#NAME?' }
      }
      case 'bin':
        return binary(n.op, scalarArg(n.l, ev), scalarArg(n.r, ev))
    }
  }
  const result = ev(ast)
  return isRange(result) ? ERR.value : result
}

function binary(op: string, l: Value, r: Value): Value {
  if (isError(l)) return l
  if (isError(r)) return r
  if (op === '&') {
    const a = toText(l)
    const b = toText(r)
    return isError(a) ? a : isError(b) ? b : a + b
  }
  if (['=', '<>', '<', '>', '<=', '>='].includes(op)) {
    const c = compare(l, r)
    return op === '=' ? c === 0 : op === '<>' ? c !== 0 : op === '<' ? c < 0 : op === '>' ? c > 0 : op === '<=' ? c <= 0 : c >= 0
  }
  const a = toNumber(l)
  const b = toNumber(r)
  if (isError(a)) return a
  if (isError(b)) return b
  switch (op) {
    case '+':
      return a + b
    case '-':
      return a - b
    case '*':
      return a * b
    case '/':
      return b === 0 ? ERR.div0 : a / b
    case '^':
      return a ** b
  }
  return ERR.value
}

/** 조건부서식 규칙의 결과가 "참"인가: TRUE 또는 0이 아닌 숫자만. 문자·오류·빈 값은 거짓. */
export function isTruthy(v: Value): boolean {
  return v === true || (typeof v === 'number' && v !== 0)
}
