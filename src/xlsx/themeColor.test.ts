import { describe, expect, it } from 'vitest'
import { parseTheme, resolveThemeColor } from './themeColor'

// ExcelJS가 기본으로 묶어 쓰는 theme1.xml과 같은 구조(node_modules/exceljs/lib/xlsx/xml/theme1.xml) —
// 실제 파일 대부분이 이 Office 기본 팔레트를 그대로 쓰거나 색만 바꾼 변형이다.
const THEME_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F497D"/></a:dk2>
      <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
      <a:accent1><a:srgbClr val="4F81BD"/></a:accent1>
      <a:accent2><a:srgbClr val="C0504D"/></a:accent2>
      <a:accent3><a:srgbClr val="9BBB59"/></a:accent3>
      <a:accent4><a:srgbClr val="8064A2"/></a:accent4>
      <a:accent5><a:srgbClr val="4BACC6"/></a:accent5>
      <a:accent6><a:srgbClr val="F79646"/></a:accent6>
      <a:hlink><a:srgbClr val="0000FF"/></a:hlink>
      <a:folHlink><a:srgbClr val="800080"/></a:folHlink>
    </a:clrScheme>
  </a:themeElements>
</a:theme>`

describe('parseTheme', () => {
  it('구조가 다른 XML이면 null', () => {
    expect(parseTheme('<a:theme xmlns:a="x"><a:themeElements/></a:theme>')).toBeNull()
  })

  it('잘못된 XML이면 null', () => {
    expect(parseTheme('<not valid')).toBeNull()
  })

  it('12개 슬롯을 XML 선언 순서가 아니라 셀 theme 인덱스 순서로 뽑는다(lt1/dk1이 맨 앞으로 뒤집힘)', () => {
    const theme = parseTheme(THEME_XML)
    expect(theme?.slots).toEqual([
      '#FFFFFF', // 0 = lt1 (sysClr lastClr)
      '#000000', // 1 = dk1 (sysClr lastClr)
      '#EEECE1', // 2 = lt2
      '#1F497D', // 3 = dk2
      '#4F81BD', // 4 = accent1
      '#C0504D', // 5 = accent2
      '#9BBB59', // 6 = accent3
      '#8064A2', // 7 = accent4
      '#4BACC6', // 8 = accent5
      '#F79646', // 9 = accent6
      '#0000FF', // 10 = hlink
      '#800080', // 11 = folHlink
    ])
  })
})

describe('resolveThemeColor', () => {
  const theme = parseTheme(THEME_XML)!

  it('tint 없이 슬롯 색을 그대로 돌려준다', () => {
    expect(resolveThemeColor(theme, 4)).toBe('#4F81BD') // accent1
  })

  it('양수 tint는 밝게(흰쪽으로) 민다', () => {
    const lightened = resolveThemeColor(theme, 4, 0.6)!
    expect(lightened).not.toBe('#4F81BD')
    // 더 밝아졌는지는 RGB 평균으로 대충 확인(정확한 HSL 보정값을 하드코딩하지 않음)
    const avg = (hex: string) => {
      const n = hex.replace('#', '')
      return (parseInt(n.slice(0, 2), 16) + parseInt(n.slice(2, 4), 16) + parseInt(n.slice(4, 6), 16)) / 3
    }
    expect(avg(lightened)).toBeGreaterThan(avg('#4F81BD'))
  })

  it('음수 tint는 어둡게(검정쪽으로) 민다', () => {
    const darkened = resolveThemeColor(theme, 4, -0.5)!
    const avg = (hex: string) => {
      const n = hex.replace('#', '')
      return (parseInt(n.slice(0, 2), 16) + parseInt(n.slice(2, 4), 16) + parseInt(n.slice(4, 6), 16)) / 3
    }
    expect(avg(darkened)).toBeLessThan(avg('#4F81BD'))
  })

  it('범위 밖 인덱스는 null', () => {
    expect(resolveThemeColor(theme, 99)).toBeNull()
  })
})
