import { act } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { FilterMenu } from './FilterMenu'
import type { ValueEntry } from './viewFilter'

// React의 act() 환경 표시 — 없으면 테스트에서 경고가 난다.
;(globalThis as unknown as { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT = true

const entries: ValueEntry[] = [
  { value: '가', count: 3, included: true },
  { value: '나', count: 2, included: false },
  { value: '', count: 1, included: true },
]

let root: Root
let host: HTMLDivElement

function mount(props: Partial<Parameters<typeof FilterMenu>[0]> = {}) {
  const handlers = {
    onToggle: vi.fn(),
    onSetListed: vi.fn(),
    onClearColumn: vi.fn(),
    onResetToFile: vi.fn(),
    onClose: vi.fn(),
  }
  act(() => {
    root.render(
      <FilterMenu title="구분" anchor={{ left: 10, top: 10, bottom: 30 }} entries={entries} columnFiltered canResetToFile={false} {...handlers} {...props} />,
    )
  })
  return handlers
}

const menu = () => document.body.querySelector('.filter-menu') as HTMLElement
const checkboxes = () => Array.from(menu().querySelectorAll<HTMLInputElement>('.filter-menu__list input[type="checkbox"]'))

beforeEach(() => {
  host = document.createElement('div')
  document.body.appendChild(host)
  root = createRoot(host)
})
afterEach(() => {
  act(() => root.unmount())
  host.remove()
})

describe('FilterMenu', () => {
  it('제목과 값 목록을 보여주고, 빈 값은 (빈 셀)로 표시한다', () => {
    mount()
    expect(menu().querySelector('.filter-menu__title')?.textContent).toBe('구분')
    expect(Array.from(menu().querySelectorAll('.filter-menu__list .filter-menu__value')).map((e) => e.textContent)).toEqual(['가', '나', '(빈 셀)'])
    expect(checkboxes().map((c) => c.checked)).toEqual([true, false, true])
  })

  it('체크박스를 누르면 그 값으로 onToggle', () => {
    const h = mount()
    act(() => checkboxes()[1].click())
    expect(h.onToggle).toHaveBeenCalledWith('나')
  })

  it('검색어로 목록을 좁힌다', () => {
    mount()
    const input = menu().querySelector<HTMLInputElement>('.filter-menu__search')!
    act(() => {
      // React가 추적하는 value setter를 거쳐야 onChange가 발생한다
      Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!.call(input, '나')
      input.dispatchEvent(new Event('input', { bubbles: true }))
    })
    expect(checkboxes()).toHaveLength(1)
    expect(menu().querySelector('.filter-menu__list .filter-menu__value')?.textContent).toBe('나')
  })

  it('(모두 선택): 일부만 체크돼 있으면 전부 켜고, 전부 체크돼 있으면 전부 끈다', () => {
    const h = mount()
    const all = menu().querySelector<HTMLInputElement>('.filter-menu__row--all input')!
    expect(all.indeterminate).toBe(true)
    act(() => all.click())
    expect(h.onSetListed).toHaveBeenLastCalledWith(['가', '나', ''], true)

    act(() => root.unmount())
    root = createRoot(host)
    const h2 = mount({ entries: entries.map((e) => ({ ...e, included: true })) })
    act(() => menu().querySelector<HTMLInputElement>('.filter-menu__row--all input')!.click())
    expect(h2.onSetListed).toHaveBeenLastCalledWith(['가', '나', ''], false)
  })

  it('일치하는 값이 없으면 안내 문구', () => {
    mount({ entries: [] })
    expect(menu().textContent).toContain('일치하는 값이 없어요')
  })

  it('열 필터가 없으면 "이 열 필터 해제"는 비활성, 건드렸을 때만 "파일 원본 상태로"가 보인다', () => {
    mount({ columnFiltered: false, canResetToFile: false })
    const labels = () => Array.from(menu().querySelectorAll<HTMLButtonElement>('.filter-menu__footer button'))
    expect(labels().find((b) => b.textContent === '이 열 필터 해제')?.disabled).toBe(true)
    expect(labels().some((b) => b.textContent === '파일 원본 상태로')).toBe(false)

    act(() => root.unmount())
    root = createRoot(host)
    const h = mount({ columnFiltered: true, canResetToFile: true })
    const reset = labels().find((b) => b.textContent === '파일 원본 상태로')!
    act(() => reset.click())
    expect(h.onResetToFile).toHaveBeenCalled()
  })

  it('Esc와 바깥 클릭으로 닫히고, ▼ 버튼 클릭은 바깥으로 치지 않는다', () => {
    const h = mount()
    act(() => {
      document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }))
    })
    expect(h.onClose).toHaveBeenCalledTimes(1)

    const outside = document.createElement('div')
    document.body.appendChild(outside)
    act(() => {
      outside.dispatchEvent(new Event('pointerdown', { bubbles: true }))
    })
    expect(h.onClose).toHaveBeenCalledTimes(2)

    const btn = document.createElement('button')
    btn.className = 'grid-filter-btn'
    document.body.appendChild(btn)
    act(() => {
      btn.dispatchEvent(new Event('pointerdown', { bubbles: true }))
    })
    expect(h.onClose).toHaveBeenCalledTimes(2) // 토글은 버튼 쪽 몫이라 여기서 닫지 않는다
    outside.remove()
    btn.remove()
  })
})
