/**
 * 엑셀 테마 색상(`xl/theme/theme1.xml`) 해석. ExcelJS는 이 파일을 공식적으로
 * 지원하지 않는다(소스 코드 자체 주석: "themes are not an exposed feature, meddle
 * at your peril!"). 원본 XML은 ZIP에서 직접 꺼내고(workbookParts.readThemeXml),
 * 여기서 `DOMParser`로 파싱한다.
 *
 * 왜 필요한가: 셀 서식의 색상은 직접 RGB(`argb`)가 아니라 테마 인덱스(`theme`)로
 * 지정되는 경우가 훨씬 흔하다 — Excel "채우기 색" 팔레트 맨 윗줄이 전부 테마
 * 색상이기 때문이다. `argb`만 읽으면 실사용 파일 대부분의 색이 안 보인다
 * (src/xlsx/cellStyle.ts가 이 모듈의 결과를 받아 그 간극을 메운다).
 */

/**
 * `<a:clrScheme>`의 12개 색 슬롯 — **XML 선언 순서가 아니라 셀 스타일의 theme
 * 인덱스(0~11) 순서**로 정렬한다. 둘이 다른 건 잘 알려진 OOXML 특성이다: XML은
 * dk1,lt1,dk2,lt2,accent1-6,hlink,folHlink 순으로 쓰여 있지만, 셀이 참조하는
 * 인덱스는 lt1,dk1,lt2,dk2,accent1-6,hlink,folHlink 순이다(배경/글자색이 앞에서
 * 한 번 뒤집힘).
 */
const THEME_INDEX_ORDER = [
  'lt1',
  'dk1',
  'lt2',
  'dk2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
] as const

export interface ThemeColors {
  /** theme 인덱스(0~11)로 바로 찾는 #RRGGBB 배열. 없는 슬롯은 빈 문자열. */
  slots: string[]
  /**
   * 테마의 선 스타일 목록(`a:lnStyleLst`)의 두께(EMU) — 도형이 `lnRef idx="2"`처럼 스타일
   * 참조만으로 선을 그릴 때 idx번째(1부터) 두께가 필요하다. Office 버전마다 값이 달라서
   * (2007: 9525/25400/38100, 2013+: 6350/12700/19050) 하드코딩하지 않고 파일에서 읽는다.
   * 테마에 없으면 undefined.
   */
  lineWidthsEmu?: number[]
}

/** `<a:srgbClr val="RRGGBB"/>` 또는 `<a:sysClr val="..." lastClr="RRGGBB"/>`에서 색을 뽑는다. */
function readColorElement(parent: Element | null): string {
  if (!parent) return ''
  const srgb = parent.getElementsByTagName('a:srgbClr')[0]
  if (srgb) return `#${srgb.getAttribute('val') ?? ''}`
  const sys = parent.getElementsByTagName('a:sysClr')[0]
  if (sys) return `#${sys.getAttribute('lastClr') ?? ''}`
  return ''
}

/** 테마 원본 XML 문자열을 파싱한다. 파싱 실패/구조가 다르면 null. */
export function parseTheme(xml: string): ThemeColors | null {
  const doc = new DOMParser().parseFromString(xml, 'application/xml')
  if (doc.getElementsByTagName('parsererror')[0]) return null

  const clrScheme = doc.getElementsByTagName('a:clrScheme')[0]
  if (!clrScheme) return null

  const slots = THEME_INDEX_ORDER.map((name) => {
    const el = Array.from(clrScheme.children).find((c) => c.tagName === `a:${name}`) ?? null
    return readColorElement(el)
  })

  const lineWidthsEmu = Array.from(doc.getElementsByTagName('a:lnStyleLst')[0]?.children ?? [])
    .map((ln) => Number(ln.getAttribute('w')))
    .filter((w) => Number.isFinite(w) && w > 0)

  return lineWidthsEmu.length > 0 ? { slots, lineWidthsEmu } : { slots }
}

export function hexToRgb(hex: string): [number, number, number] {
  const n = hex.replace('#', '')
  return [parseInt(n.slice(0, 2), 16), parseInt(n.slice(2, 4), 16), parseInt(n.slice(4, 6), 16)]
}

export function rgbToHex(r: number, g: number, b: number): string {
  const clamp = (v: number) => Math.max(0, Math.min(255, Math.round(v)))
  const hex = [r, g, b]
    .map(clamp)
    .map((v) => v.toString(16).padStart(2, '0'))
    .join('')
  return `#${hex}`.toUpperCase()
}

export function rgbToHsl(r: number, g: number, b: number): [number, number, number] {
  const rn = r / 255
  const gn = g / 255
  const bn = b / 255
  const max = Math.max(rn, gn, bn)
  const min = Math.min(rn, gn, bn)
  const l = (max + min) / 2
  if (max === min) return [0, 0, l]

  const d = max - min
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min)
  let h: number
  switch (max) {
    case rn:
      h = (gn - bn) / d + (gn < bn ? 6 : 0)
      break
    case gn:
      h = (bn - rn) / d + 2
      break
    default:
      h = (rn - gn) / d + 4
  }
  return [h / 6, s, l]
}

export function hslToRgb(h: number, s: number, l: number): [number, number, number] {
  if (s === 0) {
    const v = l * 255
    return [v, v, v]
  }
  const hue2rgb = (p: number, q: number, t: number) => {
    let tt = t
    if (tt < 0) tt += 1
    if (tt > 1) tt -= 1
    if (tt < 1 / 6) return p + (q - p) * 6 * tt
    if (tt < 1 / 2) return q
    if (tt < 2 / 3) return p + (q - p) * (2 / 3 - tt) * 6
    return p
  }
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s
  const p = 2 * l - q
  return [hue2rgb(p, q, h + 1 / 3) * 255, hue2rgb(p, q, h) * 255, hue2rgb(p, q, h - 1 / 3) * 255]
}

/**
 * Excel의 `tint` 보정 — HSL 명도(lightness)만 옮긴다(RGB 채널을 그냥 스케일하면
 * 흰색/검정 근처에서 색이 눈에 띄게 틀어진다). 공식은 OOXML 색상 변환 관례 그대로:
 * tint<0이면 더 어둡게(`l *= 1+tint`), tint>0이면 더 밝게(`l = l*(1-tint) + tint`).
 */
function applyTint(hex: string, tint: number): string {
  if (!tint) return hex
  const [r, g, b] = hexToRgb(hex)
  const [h, s, l] = rgbToHsl(r, g, b)
  const nl = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint
  const [nr, ng, nb] = hslToRgb(h, s, Math.max(0, Math.min(1, nl)))
  return rgbToHex(nr, ng, nb)
}

/** theme 인덱스(0~11) + 선택적 tint로 최종 CSS 색을 얻는다. 슬롯이 비어있으면 null. */
export function resolveThemeColor(theme: ThemeColors, index: number, tint?: number): string | null {
  const base = theme.slots[index]
  if (!base) return null
  return tint ? applyTint(base, tint) : base
}

/** DrawingML이 색을 이름으로 부를 때의 별칭 — bg1/tx1/bg2/tx2는 각각 lt1/dk1/lt2/dk2와 같은 슬롯이다. */
const SCHEME_NAME_ALIASES: Record<string, string> = { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' }

/**
 * 도형/그림 XML의 `<a:schemeClr val="accent1"/>`처럼 **이름**으로 참조된 테마 색을 찾는다.
 * (셀 서식은 인덱스, 도형은 이름으로 같은 팔레트를 가리킨다.) 모르는 이름이거나 슬롯이
 * 비어있으면 null.
 */
export function resolveThemeColorByName(theme: ThemeColors, name: string): string | null {
  const canonical = SCHEME_NAME_ALIASES[name] ?? name
  const index = (THEME_INDEX_ORDER as readonly string[]).indexOf(canonical)
  if (index < 0) return null
  return theme.slots[index] || null
}
