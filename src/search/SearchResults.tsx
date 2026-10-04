import { Fragment, useEffect, useRef } from 'react'
import { cellAddress } from '../xlsx/cellRef'
import { makeSnippet, type SearchHit } from './searchEngine'

/** 한 번에 그리는 결과 개수 상한 — 수천 건이어도 목록이 화면을 멈추지 않게. 나머지는 검색어를 좁히게 안내한다. */
const RENDER_LIMIT = 300

interface SearchResultsProps {
  hits: SearchHit[]
  sheetNames: string[]
  currentIndex: number
  onSelect: (index: number) => void
}

function locationLabel(hit: SearchHit): string {
  if (hit.kind === 'cell') return cellAddress(hit.row, hit.col)
  return hit.location ? `도형 · ${hit.location}` : '도형'
}

/**
 * 전체 시트 검색의 결과 목록: 시트별로 묶어 `주소 · 앞뒤 글자`를 보여주고, 누르면 그 시트·위치로
 * 이동한다. 현재 결과는 강조하고 이동할 때 목록 안에서도 보이게 스크롤한다.
 */
export function SearchResults({ hits, sheetNames, currentIndex, onSelect }: SearchResultsProps) {
  const listRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    listRef.current?.querySelector('[aria-current="true"]')?.scrollIntoView?.({ block: 'nearest' })
  }, [currentIndex])

  if (hits.length === 0) return null
  const shown = hits.slice(0, RENDER_LIMIT)

  return (
    <div className="search-results" ref={listRef}>
      {shown.map((hit, i) => {
        const startOfSheet = i === 0 || shown[i - 1].sheetIndex !== hit.sheetIndex
        const snippet = makeSnippet(hit)
        return (
          <Fragment key={i}>
            {startOfSheet && (
              <div className="search-results__sheet">
                {sheetNames[hit.sheetIndex] ?? `시트 ${hit.sheetIndex + 1}`}
                <span className="search-results__sheet-count">{hits.filter((h) => h.sheetIndex === hit.sheetIndex).length}</span>
              </div>
            )}
            <button type="button" className="search-results__item" aria-current={i === currentIndex} onClick={() => onSelect(i)}>
              <span className="search-results__loc">{locationLabel(hit)}</span>
              <span className="search-results__text">
                {snippet.before}
                {snippet.match && <mark>{snippet.match}</mark>}
                {snippet.after}
              </span>
            </button>
          </Fragment>
        )
      })}
      {hits.length > shown.length && <div className="search-results__more">{hits.length - shown.length}건 더 있어요 — 검색어를 더 구체적으로 입력해 보세요</div>}
    </div>
  )
}
