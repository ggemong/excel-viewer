/**
 * Excel 표시 형식 코드(`#,##0"원"`, `0000`, `yyyy-mm-dd`, `h:mm AM/PM` 등)를 해석해서 숫자
 * 하나를 화면 문자열로 만든다. 날짜/시간 값도 Excel 직렬번호(숫자)로 들어와 같은 경로를
 * 탄다(src/xlsx/cellValue.ts의 dateToSerial) — 날짜를 문자열로 들고 다니면 시간대에 따라
 * 시각이 밀리기 때문이다.
 *
 * 이전 구현은 정규식 몇 개로 "자주 쓰는 것만" 뽑았는데, 실제 한글 파일에서 접미사(`"원"`)를
 * 접두사로 찍거나 `0000`의 앞자리 0이 사라지는 등 틀린 결과가 나왔다. 부분 패치로는 끝이
 * 없어서 서식 코드를 구간 분리 -> 토큰화 -> 렌더링의 정식 구조로 바꿨다.
 *
 * 범위(의도적 제한): 색상 구간([Red] 등)과 조건 구간([>100])은 무시한다 — 이 앱은 어두운
 * 테마라 Excel의 순수 파랑/검정 같은 색을 그대로 쓰면 안 보이는 경우가 많다. 분수 서식
 * (`# ?/?`)은 해석하지 않고 일반 숫자로 보여준다. 조건 구간은 양수;음수;0;텍스트 순서만
 * 지원한다.
 */

import { EXCEL_EPOCH_UTC_MS, MS_PER_DAY } from './excelDate'

type Lang = 'ko' | 'en'

type Token =
  | { t: 'lit'; s: string }
  | { t: 'digit'; c: '0' | '#' | '?' }
  | { t: 'point' }
  | { t: 'comma' }
  | { t: 'percent' }
  | { t: 'text' }
  | { t: 'general' }
  | { t: 'exp'; plus: boolean; digits: number }
  | { t: 'date'; kind: 'y' | 'm' | 'd' | 'h' | 's' | 'a'; len: number }
  | { t: 'elapsed'; kind: 'h' | 'm' | 's'; len: number }
  | { t: 'ampm'; short: boolean }
  | { t: 'subsec'; digits: number }

interface ParsedSection {
  tokens: Token[]
  lang: Lang
}

const KO_WEEKDAYS = ['일', '월', '화', '수', '목', '금', '토']
const EN_WEEKDAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
const EN_MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']

/** Windows LCID의 하위 10비트가 주 언어 ID — 0x09 영어, 0x12 한국어. */
const PRIMARY_LANG_MASK = 0x3ff
const LANG_EN = 0x09
const LANG_KO = 0x12

const SECONDS_PER_DAY = 86_400

/** General 표시에서 숫자를 지수 표기로 바꾸는 경계(Excel 기본 열폭에서 11자리가 한계). */
const GENERAL_EXP_UPPER = 1e11
const GENERAL_EXP_LOWER = 1e-9
const GENERAL_MIN_PRECISION = 10

/**
 * ExcelJS가 내장 형식 ID를 돌려줄 때 쓰는 코드 중, Excel이 "시스템 로케일 기본 날짜"로
 * 그리는 것들 — 문자 그대로 해석하면 미국식(mm-dd-yy)이 되어 한국 사용자에게 틀리게 보인다.
 */
const LOCALE_DEPENDENT_CODES: Record<string, string> = {
  'mm-dd-yy': 'yyyy-mm-dd',
  'm/d/yy': 'yyyy-mm-dd',
  'm/d/yy h:mm': 'yyyy-mm-dd hh:mm',
}

/** `;`로 구간을 나눈다 — 따옴표/대괄호/역슬래시 안의 `;`는 구분자가 아니다. */
function splitSections(code: string): string[] {
  const sections: string[] = []
  let cur = ''
  let inQuote = false
  let inBracket = false
  for (let i = 0; i < code.length; i++) {
    const ch = code[i]
    if (inQuote) {
      cur += ch
      if (ch === '"') inQuote = false
    } else if (inBracket) {
      cur += ch
      if (ch === ']') inBracket = false
    } else if (ch === '\\') {
      cur += ch + (code[i + 1] ?? '')
      i++
    } else if (ch === '"') {
      inQuote = true
      cur += ch
    } else if (ch === '[') {
      inBracket = true
      cur += ch
    } else if (ch === ';') {
      sections.push(cur)
      cur = ''
    } else {
      cur += ch
    }
  }
  sections.push(cur)
  return sections
}

function langFromLcid(hex: string, current: Lang): Lang {
  const lcid = Number.parseInt(hex, 16)
  if (Number.isNaN(lcid)) return current
  const primary = lcid & PRIMARY_LANG_MASK
  if (primary === LANG_KO) return 'ko'
  if (primary === LANG_EN) return 'en'
  return current
}

function tokenize(section: string): ParsedSection {
  const tokens: Token[] = []
  // 로케일 표기([$-409] 등)가 없으면 한국어 — 이 제품의 사용자 기본 환경.
  let lang: Lang = 'ko'
  let i = 0
  while (i < section.length) {
    const c = section[i]
    const rest = section.slice(i)

    if (c === '"') {
      const j = section.indexOf('"', i + 1)
      const end = j < 0 ? section.length : j
      tokens.push({ t: 'lit', s: section.slice(i + 1, end) })
      i = end + 1
    } else if (c === '\\') {
      tokens.push({ t: 'lit', s: section[i + 1] ?? '' })
      i += 2
    } else if (c === '_') {
      tokens.push({ t: 'lit', s: ' ' })
      i += 2
    } else if (c === '*') {
      i += 2
    } else if (c === '[') {
      const j = section.indexOf(']', i)
      const end = j < 0 ? section.length : j
      const content = section.slice(i + 1, end)
      if (content.startsWith('$')) {
        const body = content.slice(1)
        const dash = body.lastIndexOf('-')
        const symbol = dash < 0 ? body : body.slice(0, dash)
        if (dash >= 0) lang = langFromLcid(body.slice(dash + 1), lang)
        if (symbol) tokens.push({ t: 'lit', s: symbol })
      } else {
        const elapsed = content.match(/^(h{1,2}|m{1,2}|s{1,2})$/i)
        if (elapsed) {
          tokens.push({ t: 'elapsed', kind: elapsed[1][0].toLowerCase() as 'h' | 'm' | 's', len: elapsed[1].length })
        }
        // 색상/조건/DBNum 등은 의도적으로 무시(모듈 설명 참고)
      }
      i = end + 1
    } else if (c === '.') {
      tokens.push({ t: 'point' })
      i++
    } else if (c === ',') {
      tokens.push({ t: 'comma' })
      i++
    } else if (c === '0' || c === '#' || c === '?') {
      tokens.push({ t: 'digit', c })
      i++
    } else if (c === '%') {
      tokens.push({ t: 'percent' })
      i++
    } else if (c === '@') {
      tokens.push({ t: 'text' })
      i++
    } else if (/^general/i.test(rest)) {
      tokens.push({ t: 'general' })
      i += 7
    } else if (/^(am\/pm|a\/p)/i.test(rest)) {
      const short = /^a\/p/i.test(rest)
      tokens.push({ t: 'ampm', short })
      i += short ? 3 : 5
    } else if (/^e[+-]/i.test(rest)) {
      let n = 0
      while (section[i + 2 + n] === '0') n++
      tokens.push({ t: 'exp', plus: section[i + 1] === '+', digits: n })
      i += 2 + n
    } else if (/[ymdhs]/i.test(c)) {
      let n = 1
      while (section[i + n]?.toLowerCase() === c.toLowerCase()) n++
      tokens.push({ t: 'date', kind: c.toLowerCase() as 'y' | 'm' | 'd' | 'h' | 's', len: n })
      i += n
    } else if (c === 'a' || c === 'A') {
      let n = 1
      while (section[i + n]?.toLowerCase() === 'a') n++
      tokens.push({ t: 'date', kind: 'a', len: n })
      i += n
    } else {
      tokens.push({ t: 'lit', s: c })
      i++
    }
  }

  // 초 뒤의 ".000"은 소수 초 자리 — point+digit(0)... 을 subsec 하나로 합친다.
  const merged: Token[] = []
  for (let k = 0; k < tokens.length; k++) {
    const tk = tokens[k]
    const prev = merged[merged.length - 1]
    const prevIsSeconds = prev && ((prev.t === 'date' && prev.kind === 's') || (prev.t === 'elapsed' && prev.kind === 's'))
    if (tk.t === 'point' && prevIsSeconds && tokens[k + 1]?.t === 'digit') {
      let n = 0
      while (tokens[k + 1 + n]?.t === 'digit') n++
      merged.push({ t: 'subsec', digits: n })
      k += n
    } else {
      merged.push(tk)
    }
  }
  return { tokens: merged, lang }
}

function isDateSection(tokens: Token[]): boolean {
  return tokens.some((tk) => tk.t === 'date' || tk.t === 'elapsed' || tk.t === 'ampm')
}

/** 안 쓰는 부분을 건너뛰고 앞/뒤에서 가장 가까운 비리터럴 토큰을 찾는다. */
function neighbor(tokens: Token[], index: number, dir: 1 | -1): Token | undefined {
  for (let k = index + dir; k >= 0 && k < tokens.length; k += dir) {
    if (tokens[k].t !== 'lit') return tokens[k]
  }
  return undefined
}

/** 날짜 서식의 `m`은 달/분 둘 다 가능 — 바로 앞이 시, 바로 뒤가 초이면 분이다. */
function isMinuteToken(tokens: Token[], index: number): boolean {
  const prev = neighbor(tokens, index, -1)
  const next = neighbor(tokens, index, 1)
  const isHour = (tk?: Token) => !!tk && ((tk.t === 'date' && tk.kind === 'h') || (tk.t === 'elapsed' && tk.kind === 'h'))
  const isSec = (tk?: Token) => !!tk && ((tk.t === 'date' && tk.kind === 's') || (tk.t === 'elapsed' && tk.kind === 's'))
  return isHour(prev) || isSec(next)
}

function pad(n: number, len: number): string {
  return String(n).padStart(len, '0')
}

function renderDate(serial: number, { tokens, lang }: ParsedSection): string {
  if (serial < 0) return '########'
  const subsec = tokens.find((tk) => tk.t === 'subsec') as Extract<Token, { t: 'subsec' }> | undefined

  // 소수 초 자리가 없으면 가장 가까운 초로 반올림해서 59.9999 같은 부동소수 잡음을 없앤다.
  const totalMs = Math.round(serial * MS_PER_DAY)
  const totalSecs = subsec ? Math.floor(totalMs / 1000) : Math.round(serial * SECONDS_PER_DAY)
  const ms = subsec ? totalMs % 1000 : 0
  const days = Math.floor(totalSecs / SECONDS_PER_DAY)
  const secOfDay = totalSecs - days * SECONDS_PER_DAY

  const date = new Date(EXCEL_EPOCH_UTC_MS + days * MS_PER_DAY)
  const year = date.getUTCFullYear()
  const month = date.getUTCMonth()
  const day = date.getUTCDate()
  const weekday = date.getUTCDay()
  const hour = Math.floor(secOfDay / 3600)
  const minute = Math.floor((secOfDay % 3600) / 60)
  const second = secOfDay % 60
  const hasAmPm = tokens.some((tk) => tk.t === 'ampm')

  return tokens
    .map((tk, idx) => {
      switch (tk.t) {
        case 'lit':
          return tk.s
        case 'date': {
          if (tk.kind === 'y') return tk.len <= 2 ? pad(year % 100, 2) : pad(year, 4)
          if (tk.kind === 'd') {
            if (tk.len <= 2) return tk.len === 1 ? String(day) : pad(day, 2)
            return lang === 'ko' ? (tk.len === 3 ? KO_WEEKDAYS[weekday] : `${KO_WEEKDAYS[weekday]}요일`) : tk.len === 3 ? EN_WEEKDAYS[weekday].slice(0, 3) : EN_WEEKDAYS[weekday]
          }
          if (tk.kind === 'a') return tk.len <= 3 ? KO_WEEKDAYS[weekday] : `${KO_WEEKDAYS[weekday]}요일`
          if (tk.kind === 'h') {
            const h = hasAmPm ? hour % 12 || 12 : hour
            return tk.len === 1 ? String(h) : pad(h, 2)
          }
          if (tk.kind === 's') return tk.len === 1 ? String(second) : pad(second, 2)
          // 'm': 분 또는 달
          if (isMinuteToken(tokens, idx)) return tk.len === 1 ? String(minute) : pad(minute, 2)
          if (tk.len <= 2) return tk.len === 1 ? String(month + 1) : pad(month + 1, 2)
          if (lang === 'ko') return `${month + 1}월`
          const name = EN_MONTHS[month]
          return tk.len === 3 ? name.slice(0, 3) : tk.len === 5 ? name[0] : name
        }
        case 'elapsed': {
          const total = tk.kind === 'h' ? Math.floor(totalSecs / 3600) : tk.kind === 'm' ? Math.floor(totalSecs / 60) : totalSecs
          return pad(total, tk.len)
        }
        case 'ampm': {
          const pm = hour >= 12
          if (lang === 'ko') return pm ? '오후' : '오전'
          return tk.short ? (pm ? 'P' : 'A') : pm ? 'PM' : 'AM'
        }
        case 'subsec':
          return `.${pad(ms, 3).slice(0, tk.digits)}`
        default:
          return ''
      }
    })
    .join('')
}

/** 반올림은 Excel처럼 "보이는 십진수" 기준 — 2.675를 이진 부동소수 그대로 toFixed하면 2.67이 되어 틀린다. */
function roundDecimal(x: number, digits: number): number {
  if (!Number.isFinite(x)) return x
  const str = String(x)
  if (str.includes('e')) return Number(x.toFixed(Math.min(digits, 100)))
  return Number(`${Math.round(Number(`${str}e${digits}`))}e-${digits}`)
}

function groupThousands(digits: string): string {
  return digits.replace(/\B(?=(\d{3})+(?!\d))/g, ',')
}

/** 서식 코드 없는("General") 숫자 표시 — 0.1+0.2 같은 부동소수 잡음을 걷고 큰 수는 지수 표기. */
export function generalNumber(value: number): string {
  if (!Number.isFinite(value)) return String(value)
  if (value === 0) return '0'
  const abs = Math.abs(value)

  if (abs >= GENERAL_EXP_UPPER || abs < GENERAL_EXP_LOWER) {
    const [mant, exp] = value.toExponential(5).split('e')
    const trimmed = mant.includes('.') ? mant.replace(/\.?0+$/, '') : mant
    const sign = exp.startsWith('-') ? '-' : '+'
    return `${trimmed}E${sign}${exp.replace(/^[+-]/, '').padStart(2, '0')}`
  }

  const intDigits = abs >= 1 ? Math.floor(Math.log10(abs)) + 1 : 1
  const rounded = Number(value.toPrecision(Math.max(GENERAL_MIN_PRECISION, intDigits)))
  if (abs < 1e-6) {
    return rounded.toFixed(Math.min(100, GENERAL_MIN_PRECISION - Math.floor(Math.log10(abs)))).replace(/\.?0+$/, '')
  }
  return String(rounded)
}

function renderNumber(abs: number, { tokens }: ParsedSection, negativePrefix: string): string {
  const general = tokens.find((tk) => tk.t === 'general')
  if (general) {
    return negativePrefix + tokens.map((tk) => (tk.t === 'general' ? generalNumber(abs) : tk.t === 'lit' ? tk.s : tk.t === 'percent' ? '%' : '')).join('')
  }

  const digitCount = tokens.filter((tk) => tk.t === 'digit').length
  if (digitCount === 0 && !tokens.some((tk) => tk.t === 'exp')) {
    // 숫자 자리표시가 없는 구간(예: `;;;` 이나 `"없음"`)은 리터럴만 보인다.
    return tokens.map((tk) => (tk.t === 'lit' ? tk.s : tk.t === 'percent' ? '%' : '')).join('')
  }

  // 분수 서식(`# ?/?`)은 해석하지 않고 일반 숫자로 대신한다(모듈 설명 참고).
  for (let k = 1; k < tokens.length - 1; k++) {
    const tk = tokens[k]
    if (tk.t === 'lit' && tk.s === '/' && tokens[k - 1].t === 'digit' && tokens[k + 1].t === 'digit') {
      return negativePrefix + generalNumber(abs)
    }
  }

  const lastDigit = tokens.reduce((acc, tk, k) => (tk.t === 'digit' ? k : acc), -1)
  const firstDigit = tokens.findIndex((tk) => tk.t === 'digit')

  // 마지막 자리표시 뒤에 바로 붙은 쉼표는 "천 단위로 나누기", 숫자 사이 쉼표는 "천 단위 구분".
  let scaleCommas = 0
  for (let k = lastDigit + 1; k < tokens.length && tokens[k].t === 'comma'; k++) scaleCommas++
  const grouping = tokens.some((tk, k) => tk.t === 'comma' && k > firstDigit && k < lastDigit)
  const percentCount = tokens.filter((tk) => tk.t === 'percent').length
  const value = (abs * 100 ** percentCount) / 1000 ** scaleCommas

  const expToken = tokens.find((tk) => tk.t === 'exp') as Extract<Token, { t: 'exp' }> | undefined
  const pointIdx = tokens.findIndex((tk) => tk.t === 'point')
  const mantissaEnd = expToken ? tokens.indexOf(expToken) : tokens.length
  const intTokens = tokens.slice(0, pointIdx < 0 ? mantissaEnd : pointIdx)
  const fracTokens = pointIdx < 0 ? [] : tokens.slice(pointIdx + 1, mantissaEnd)
  const fracDigits = fracTokens.filter((tk) => tk.t === 'digit').length
  const intPlaceholders = intTokens.filter((tk) => tk.t === 'digit').length

  let intStr: string
  let fracStr: string
  let expSuffix = ''

  if (expToken) {
    const intN = Math.max(1, intPlaceholders)
    let exp = value === 0 ? 0 : Math.floor(Math.log10(value)) - (intN - 1)
    let mant = value === 0 ? 0 : roundDecimal(value / 10 ** exp, fracDigits)
    if (mant >= 10 ** intN) {
      mant /= 10
      exp += 1
    }
    ;[intStr, fracStr = ''] = mant.toFixed(fracDigits).split('.')
    expSuffix = `E${exp < 0 ? '-' : expToken.plus ? '+' : ''}${pad(Math.abs(exp), expToken.digits)}`
  } else {
    ;[intStr, fracStr = ''] = roundDecimal(value, fracDigits).toFixed(fracDigits).split('.')
    // 정수부에 '0' 자리표시가 하나도 없으면(`#.00`) 0은 아예 안 찍는다.
    if (intStr === '0' && !intTokens.some((tk) => tk.t === 'digit' && tk.c === '0')) intStr = ''
  }

  let intOut: string
  if (grouping) {
    const firstIntDigit = intTokens.findIndex((tk) => tk.t === 'digit')
    const lastIntDigit = intTokens.reduce((acc, tk, k) => (tk.t === 'digit' ? k : acc), -1)
    const zeroPlaceholders = intTokens.filter((tk) => tk.t === 'digit' && tk.c === '0').length
    const lit = (tk: Token) => (tk.t === 'lit' ? tk.s : tk.t === 'percent' ? '%' : '')
    intOut =
      intTokens.slice(0, firstIntDigit).map(lit).join('') +
      groupThousands(intStr.padStart(zeroPlaceholders, '0')) +
      intTokens.slice(lastIntDigit + 1).map(lit).join('')
  } else {
    // 자리표시를 오른쪽부터 채운다 — `000-0000-0000` 처럼 리터럴이 섞여도 자릿수가 맞는다.
    const digits = intStr.split('').filter(Boolean)
    let di = digits.length - 1
    const out = intTokens.map(() => '')
    for (let k = intTokens.length - 1; k >= 0; k--) {
      const tk = intTokens[k]
      if (tk.t === 'digit') {
        if (di >= 0) out[k] = digits[di--]
        else out[k] = tk.c === '0' ? '0' : tk.c === '?' ? ' ' : ''
      } else if (tk.t === 'lit') out[k] = tk.s
      else if (tk.t === 'percent') out[k] = '%'
    }
    if (di >= 0) {
      const leftmost = intTokens.findIndex((tk) => tk.t === 'digit')
      const leftover = digits.slice(0, di + 1).join('')
      if (leftmost >= 0) out[leftmost] = leftover + out[leftmost]
      else out.push(leftover)
    }
    intOut = out.join('')
  }

  let fracOut = ''
  if (fracTokens.length > 0) {
    const fd = fracStr.split('')
    let lastNonZero = -1
    fd.forEach((ch, k) => {
      if (ch !== '0') lastNonZero = k
    })
    let j = 0
    fracOut = fracTokens
      .map((tk) => {
        if (tk.t === 'digit') {
          const idx = j++
          if (idx <= lastNonZero) return fd[idx]
          return tk.c === '0' ? '0' : tk.c === '?' ? ' ' : ''
        }
        if (tk.t === 'lit') return tk.s
        if (tk.t === 'percent') return '%'
        return ''
      })
      .join('')
  }

  const trailing = expToken ? tokens.slice(mantissaEnd + 1).map((tk) => (tk.t === 'lit' ? tk.s : '')).join('') : ''
  return negativePrefix + intOut + (pointIdx >= 0 ? '.' : '') + fracOut + expSuffix + trailing
}

function normalizeCode(code: string): string {
  return LOCALE_DEPENDENT_CODES[code] ?? code
}

/**
 * 숫자(또는 날짜 직렬번호) 하나를 Excel 표시 형식 코드대로 문자열로 만든다.
 * 코드가 없거나 General이면 General 표시.
 */
export function formatNumber(value: number, code: string | null): string {
  if (!code || /^general$/i.test(code.trim())) return generalNumber(value)

  const sections = splitSections(normalizeCode(code))
  let chosen: string
  let negativePrefix = ''
  let abs = Math.abs(value)

  if (value < 0) {
    if (sections.length >= 2) chosen = sections[1]
    else {
      chosen = sections[0]
      negativePrefix = '-'
    }
  } else if (value === 0 && sections.length >= 3) {
    chosen = sections[2]
  } else {
    chosen = sections[0]
  }
  if (value === 0) abs = 0

  const parsed = tokenize(chosen)
  if (isDateSection(parsed.tokens)) {
    return renderDate(value, parsed)
  }
  return renderNumber(abs, parsed, negativePrefix)
}

/**
 * 텍스트 값에 서식 코드를 적용한다. `@`가 든 구간이 있을 때만 의미가 있고(없으면 원문 그대로),
 * 4구간 코드면 네 번째 구간을 우선한다.
 */
export function formatText(text: string, code: string | null): string {
  if (!code) return text
  const sections = splitSections(code)
  const section = sections.length >= 4 ? sections[3] : sections.find((s) => tokenize(s).tokens.some((tk) => tk.t === 'text'))
  if (section === undefined) return text
  return tokenize(section)
    .tokens.map((tk) => (tk.t === 'text' ? text : tk.t === 'lit' ? tk.s : ''))
    .join('')
}

/**
 * 이 표시 형식 코드가 날짜/시간 형식인가 — 날짜 셀의 원래 값은 직렬번호(46299 같은 숫자)라
 * 사용자에게는 의미가 없어서, 검색 등에서 "원래 값"으로 취급하지 않기 위해 구분한다.
 */
export function isDateTimeFormat(code: string | null): boolean {
  if (!code || /^general$/i.test(code.trim())) return false
  return isDateSection(tokenize(splitSections(normalizeCode(code))[0]).tokens)
}
