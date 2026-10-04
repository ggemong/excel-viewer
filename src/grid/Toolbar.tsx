import { useEffect, useRef, useState } from 'react'
import tigerFace from '../assets/tiger/tiger-face.png'
import { SUPPORTED_EXTENSIONS } from '../state/useWorkbookController'

interface ToolbarProps {
  fileName: string | null
  onOpenAnother: () => void
  compareFileName: string | null
  compareLoading: boolean
  onPickCompareFile: (file: File) => void
  onClearCompare: () => void
  /** 검색 줄 열기(Ctrl+F와 같은 동작). */
  onOpenSearch: () => void
  searchOpen: boolean
}

export function Toolbar({
  fileName,
  onOpenAnother,
  compareFileName,
  compareLoading,
  onPickCompareFile,
  onClearCompare,
  onOpenSearch,
  searchOpen,
}: ToolbarProps) {
  const compareInputRef = useRef<HTMLInputElement>(null)
  const [menuOpen, setMenuOpen] = useState(false)
  const menuRef = useRef<HTMLDivElement>(null)

  // 모바일 오버플로 메뉴 바깥을 누르거나 Escape를 누르면 닫는다 — 일반적인 드롭다운 동작.
  useEffect(() => {
    if (!menuOpen) return
    const onPointerDown = (e: PointerEvent) => {
      if (menuRef.current && !menuRef.current.contains(e.target as Node)) setMenuOpen(false)
    }
    const onKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setMenuOpen(false)
    }
    document.addEventListener('pointerdown', onPointerDown)
    document.addEventListener('keydown', onKeyDown)
    return () => {
      document.removeEventListener('pointerdown', onPointerDown)
      document.removeEventListener('keydown', onKeyDown)
    }
  }, [menuOpen])

  const compareInput = (
    <input
      ref={compareInputRef}
      type="file"
      accept={SUPPORTED_EXTENSIONS.join(',')}
      style={{ display: 'none' }}
      onChange={(e) => {
        const f = e.target.files?.[0]
        if (f) onPickCompareFile(f)
        e.target.value = ''
      }}
    />
  )

  return (
    <div className="toolbar">
      <div className="toolbar-group toolbar-brand">
        <span className="tiger-disk tiger-disk--40">
          <img src={tigerFace} alt="" />
        </span>
        <span className="wm">
          w<span className="wm-i">i</span>thv<span className="wm-i">i</span>be
        </span>
        {fileName && (
          <>
            <span aria-hidden="true">·</span>
            <span className="toolbar-filename">{fileName}</span>
          </>
        )}
      </div>
      <div className="toolbar-group">
        {fileName && (
          <>
            {/* 비교 중이면 상태 표시라 접지 않고 항상 보여준다 — 진입 버튼("버전 비교")만 좁은
                화면에서 오버플로 메뉴로 접는다(아래 .toolbar-overflow). */}
            {compareFileName && (
              <span className="pill pill--outline pill--accent">
                {compareFileName}와 비교 중
                <button type="button" className="pill-close" onClick={onClearCompare} aria-label="비교 종료">
                  ×
                </button>
              </span>
            )}

            {/* 검색은 자주 쓰는 기능이라 좁은 화면에서도 접지 않고 아이콘으로 항상 보여준다. */}
            <button
              type="button"
              className="btn btn--ghost toolbar-search"
              onClick={onOpenSearch}
              aria-label="검색"
              aria-pressed={searchOpen}
              title="검색 (Ctrl+F)"
            >
              <svg viewBox="0 0 20 20" width="16" height="16" aria-hidden="true">
                <circle cx="8.5" cy="8.5" r="5.5" fill="none" stroke="currentColor" strokeWidth="2" />
                <path d="M13 13l5 5" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
              </svg>
              <span className="toolbar-search__label">검색</span>
            </button>

            {/* 데스크톱 전용 — 넓은 화면에서는 접을 필요 없이 그냥 나열한다. */}
            <span className="toolbar-desktop-actions">
              {!compareFileName && (
                <button
                  type="button"
                  className="pill pill--outline"
                  onClick={() => compareInputRef.current?.click()}
                  disabled={compareLoading}
                >
                  {compareLoading ? '불러오는 중…' : '버전 비교'}
                </button>
              )}
              <button className="btn btn--ghost" onClick={onOpenAnother} type="button">
                다른 파일 열기
              </button>
            </span>

            {/* 모바일 전용 — 자주 안 쓰는 진입 액션만 "⋯"로 접어서 한 줄을 유지한다. */}
            <div className="toolbar-overflow" ref={menuRef}>
              <button
                type="button"
                className="btn btn--ghost toolbar-overflow-trigger"
                onClick={() => setMenuOpen((v) => !v)}
                aria-label="더보기"
                aria-expanded={menuOpen}
              >
                ⋯
              </button>
              {menuOpen && (
                <div className="toolbar-overflow-menu">
                  {!compareFileName && (
                    <button
                      type="button"
                      className="toolbar-overflow-item"
                      disabled={compareLoading}
                      onClick={() => {
                        setMenuOpen(false)
                        compareInputRef.current?.click()
                      }}
                    >
                      {compareLoading ? '불러오는 중…' : '버전 비교'}
                    </button>
                  )}
                  <button
                    type="button"
                    className="toolbar-overflow-item"
                    onClick={() => {
                      setMenuOpen(false)
                      onOpenAnother()
                    }}
                  >
                    다른 파일 열기
                  </button>
                </div>
              )}
            </div>
            {compareInput}
          </>
        )}
        <span className={`pill pill--live${fileName ? ' pill--live-loaded' : ''}`}>브라우저에서만 열림</span>
      </div>
    </div>
  )
}
