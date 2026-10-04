import type { RefObject } from 'react'
import type { SearchController } from './useSearch'

interface SearchBarProps {
  search: SearchController
  inputRef: RefObject<HTMLInputElement | null>
  onClose: () => void
}

/** 결과 상태 문구: "3 / 17", "검색 중…", "결과 없음". 검색어가 없으면 빈 문자열. */
function statusText(s: SearchController): string {
  if (s.query.trim() === '') return ''
  if (s.hits.length === 0) return s.searching ? '검색 중…' : '결과 없음'
  const position = s.currentIndex >= 0 ? s.currentIndex + 1 : '–'
  return `${position} / ${s.hits.length}${s.truncated ? '+' : ''}${s.searching ? ' …' : ''}`
}

/**
 * 검색 입력 줄. Enter/Shift+Enter로 다음/이전, Esc로 닫는다(Esc는 App의 전역 키 처리가 맡는다).
 * 범위 토글(현재 시트/전체 시트)과 옵션 토글은 버튼이지만 `aria-pressed`로 켜짐 상태를 알린다.
 */
export function SearchBar({ search: s, inputRef, onClose }: SearchBarProps) {
  const noHitsInSheet = s.scope === 'sheet' && s.query.trim() !== '' && s.hits.length === 0 && !s.searching

  return (
    <div className="search-bar" role="search">
      <div className="search-bar__row">
        <input
          ref={inputRef}
          type="search"
          className="search-bar__input"
          placeholder="검색어"
          aria-label="검색어"
          value={s.query}
          onChange={(e) => s.setQuery(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter') {
              e.preventDefault()
              s.go(e.shiftKey ? -1 : 1)
            }
          }}
        />
        <div className="search-bar__scope" role="group" aria-label="검색 범위">
          <button type="button" aria-pressed={s.scope === 'sheet'} onClick={() => s.setScope('sheet')}>
            현재 시트
          </button>
          <button type="button" aria-pressed={s.scope === 'all'} onClick={() => s.setScope('all')}>
            전체 시트
          </button>
        </div>
        <span className="search-bar__count" aria-live="polite">
          {statusText(s)}
        </span>
        <button type="button" className="btn btn--ghost search-bar__nav" onClick={() => s.go(-1)} disabled={s.hits.length === 0} aria-label="이전 결과" title="이전 (Shift+Enter)">
          ▲
        </button>
        <button type="button" className="btn btn--ghost search-bar__nav" onClick={() => s.go(1)} disabled={s.hits.length === 0} aria-label="다음 결과" title="다음 (Enter)">
          ▼
        </button>
        <button type="button" className="search-bar__toggle" aria-pressed={s.matchCase} onClick={() => s.setMatchCase(!s.matchCase)} title="대소문자 구분">
          Aa
        </button>
        <button type="button" className="search-bar__toggle" aria-pressed={s.wholeCell} onClick={() => s.setWholeCell(!s.wholeCell)} title="셀 전체가 같을 때만">
          전체 일치
        </button>
        <button type="button" className="btn btn--ghost search-bar__close" onClick={onClose} aria-label="검색 닫기">
          ×
        </button>
      </div>

      {(noHitsInSheet || s.hiddenExcluded > 0 || s.truncated) && (
        <div className="search-bar__note">
          {noHitsInSheet && (
            <button type="button" className="search-bar__link" onClick={() => s.setScope('all')}>
              이 시트에는 없어요 — 전체 시트에서 찾기
            </button>
          )}
          {s.hiddenExcluded > 0 && <span>숨겨진 행·열의 일치 {s.hiddenExcluded}건은 제외했어요</span>}
          {s.truncated && <span>결과가 많아 처음 {s.hits.length.toLocaleString('ko-KR')}건만 보여요</span>}
        </div>
      )}
    </div>
  )
}
