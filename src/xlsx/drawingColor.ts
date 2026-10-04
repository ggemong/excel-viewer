/**
 * DrawingML(도형·그림 XML)의 색 요소를 CSS 색으로 바꾼다.
 *
 * 셀 서식은 `theme`+`tint` 숫자로 색을 가리키지만, 도형은 `<a:schemeClr val="accent1">`
 * 처럼 **이름**으로 가리키고 그 뒤에 lumMod/lumOff/tint/shade/alpha 같은 변환이 자식으로
 * 붙는다. 팔레트(테마)는 같으므로 src/xlsx/themeColor.ts의 해석기를 재사용하고, 여기서는
 * 도형 쪽에만 있는 "이름 참조 + 변환 체인"만 더한다.
 */
import { hexToRgb, hslToRgb, resolveThemeColorByName, rgbToHex, rgbToHsl, type ThemeColors } from './themeColor'

const COLOR_ELEMENTS = ['srgbClr', 'sysClr', 'schemeClr', 'scrgbClr', 'prstClr']
/** DrawingML 값은 100000 = 100% */
const PERCENT_UNIT = 100_000

/** prstClr(이름 있는 색) 중 실제로 쓰이는 것들만. */
const PRESET_COLORS: Record<string, string> = {
  black: '#000000',
  white: '#FFFFFF',
  red: '#FF0000',
  green: '#008000',
  blue: '#0000FF',
  yellow: '#FFFF00',
  gray: '#808080',
  grey: '#808080',
}

/** sysClr은 lastClr(마지막으로 해석된 실제 색)이 있으면 그걸 쓴다. */
const SYSTEM_COLORS: Record<string, string> = { windowText: '#000000', window: '#FFFFFF' }

function childByLocalName(parent: Element, names: string[]): Element | null {
  for (const child of Array.from(parent.children)) {
    if (names.includes(child.localName)) return child
  }
  return null
}

function percentAttr(el: Element): number {
  return Number(el.getAttribute('val') ?? '0') / PERCENT_UNIT
}

/**
 * `parent`(solidFill, lnRef, fontRef 등) 안의 색 요소 하나를 찾아 CSS 색으로 만든다.
 * 색을 못 찾거나(테마 없음 포함) 알 수 없으면 null — 조용히 추정하지 않는다.
 *
 * @param placeholder 스타일 참조(lnRef 등)의 `phClr` 자리에 들어갈 색
 */
export function resolveDrawingColor(
  parent: Element | null | undefined,
  theme: ThemeColors | null,
  placeholder: string | null = null,
): string | null {
  if (!parent) return null
  const el = childByLocalName(parent, COLOR_ELEMENTS)
  if (!el) return null

  let hex: string | null = null
  switch (el.localName) {
    case 'srgbClr':
      hex = `#${el.getAttribute('val') ?? ''}`
      break
    case 'sysClr': {
      const last = el.getAttribute('lastClr')
      hex = last ? `#${last}` : (SYSTEM_COLORS[el.getAttribute('val') ?? ''] ?? null)
      break
    }
    case 'schemeClr': {
      const name = el.getAttribute('val') ?? ''
      hex = name === 'phClr' ? placeholder : theme ? resolveThemeColorByName(theme, name) : null
      break
    }
    case 'scrgbClr': {
      const ch = (a: string) => Math.round((Number(el.getAttribute(a) ?? '0') / PERCENT_UNIT) * 255)
      hex = rgbToHex(ch('r'), ch('g'), ch('b'))
      break
    }
    case 'prstClr':
      hex = PRESET_COLORS[el.getAttribute('val') ?? ''] ?? null
      break
  }
  if (!hex || !/^#[0-9a-fA-F]{6}$/.test(hex)) return null

  let [r, g, b] = hexToRgb(hex)
  let alpha = 1
  for (const t of Array.from(el.children)) {
    switch (t.localName) {
      case 'alpha':
        alpha = percentAttr(t)
        break
      case 'lumMod':
      case 'lumOff': {
        const [h, s, l] = rgbToHsl(r, g, b)
        const nl = t.localName === 'lumMod' ? l * percentAttr(t) : l + percentAttr(t)
        ;[r, g, b] = hslToRgb(h, s, Math.max(0, Math.min(1, nl)))
        break
      }
      case 'tint': {
        // 흰색 쪽으로: val이 작을수록 더 밝아진다.
        const k = percentAttr(t)
        ;[r, g, b] = [r, g, b].map((c) => 255 - (255 - c) * k) as [number, number, number]
        break
      }
      case 'shade': {
        const k = percentAttr(t)
        ;[r, g, b] = [r, g, b].map((c) => c * k) as [number, number, number]
        break
      }
      // satMod/hueMod 등 나머지 변환은 쓰이는 빈도가 낮아 의도적으로 무시한다.
    }
  }

  const out = rgbToHex(r, g, b)
  if (alpha >= 1) return out
  const [rr, gg, bb] = hexToRgb(out)
  return `rgba(${rr}, ${gg}, ${bb}, ${Number(alpha.toFixed(3))})`
}
