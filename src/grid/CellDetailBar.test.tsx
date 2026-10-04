import { act } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { CellDetail } from './cellDetail'
import { CellDetailBar } from './CellDetailBar'

;(globalThis as unknown as { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT = true

let root: Root
let host: HTMLDivElement

const base: CellDetail = { address: 'B5', text: '짧은 글', originalValue: null, formula: null, hyperlink: null, note: null }

function mount(detail: Partial<CellDetail> = {}, expanded = false) {
  const onToggleExpanded = vi.fn()
  act(() => root.render(<CellDetailBar detail={{ ...base, ...detail }} expanded={expanded} onToggleExpanded={onToggleExpanded} />))
  return onToggleExpanded
}
const q = (sel: string) => host.querySelector(sel) as HTMLElement | null
const button = (text: string) => Array.from(host.querySelectorAll('button')).find((b) => b.textContent === text) as HTMLButtonElement | undefined

beforeEach(() => {
  host = document.createElement('div')
  document.body.appendChild(host)
  root = createRoot(host)
})
afterEach(() => {
  act(() => root.unmount())
  host.remove()
})

describe('CellDetailBar', () => {
  it('주소와 전체 글자를 보여준다', () => {
    mount({ text: '긴 내용 전체' })
    expect(q('.cell-detail__addr')?.textContent).toBe('B5')
    expect(q('.cell-detail__text')?.textContent).toBe('긴 내용 전체')
  })

  it('빈 칸이면 (빈 셀), 복사는 비활성', () => {
    mount({ text: '' })
    expect(q('.cell-detail__text')?.textContent).toBe('(빈 셀)')
    expect(button('복사')?.disabled).toBe(true)
  })

  it('짧은 글에는 펼치기가 없고, 긴 글/줄바꿈 글에는 있다', () => {
    mount({ text: '짧음' })
    expect(button('펼치기')).toBeUndefined()
    mount({ text: '가'.repeat(61) })
    expect(button('펼치기')).toBeDefined()
    mount({ text: '첫 줄\n둘째 줄' })
    expect(button('펼치기')).toBeDefined()
  })

  it('펼침 상태에 따라 버튼 문구와 data 속성이 바뀌고, 누르면 토글 콜백', () => {
    const toggle = mount({ text: '가'.repeat(80) }, true)
    expect(q('.cell-detail')?.getAttribute('data-expanded')).toBe('true')
    act(() => button('접기')!.click())
    expect(toggle).toHaveBeenCalled()
  })

  it('복사는 서식 없는 글자 그대로 클립보드에 쓴다(탭/따옴표 이스케이프 없음)', async () => {
    const writeText = vi.fn().mockResolvedValue(undefined)
    Object.defineProperty(navigator, 'clipboard', { value: { writeText }, configurable: true })
    mount({ text: '탭\t과 "따옴표"' })
    await act(async () => button('복사')!.click())
    expect(writeText).toHaveBeenCalledWith('탭\t과 "따옴표"')
    expect(button('복사됨')).toBeDefined()
  })

  it('원래 값·수식·링크·메모를 보조 줄로 보여주고, 링크는 안전하게 새 탭', () => {
    mount({ originalValue: '1000', formula: '=A1+B1', hyperlink: 'https://a.kr', note: '확인 필요' })
    const extras = q('.cell-detail__extras')!.textContent ?? ''
    expect(extras).toContain('1000')
    expect(extras).toContain('=A1+B1')
    expect(extras).toContain('확인 필요')
    const a = q('.cell-detail__extras a') as HTMLAnchorElement
    expect(a.getAttribute('href')).toBe('https://a.kr')
    expect(a.getAttribute('rel')).toBe('noopener noreferrer')
    expect(a.getAttribute('target')).toBe('_blank')
  })

  it('보조 정보가 없으면 보조 줄 자체가 없다', () => {
    mount()
    expect(q('.cell-detail__extras')).toBeNull()
  })
})
