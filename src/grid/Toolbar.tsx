import { useRef } from 'react'
import tigerFace from '../assets/tiger/tiger-face.png'
import { SUPPORTED_EXTENSIONS } from '../state/useWorkbookController'

interface ToolbarProps {
  fileName: string | null
  onOpenAnother: () => void
  editMode: boolean
  onToggleEditMode: () => void
  isDirty: boolean
  editedCount: number
  saving: boolean
  onSave: () => void
  compareFileName: string | null
  compareLoading: boolean
  onPickCompareFile: (file: File) => void
  onClearCompare: () => void
}

export function Toolbar({
  fileName,
  onOpenAnother,
  editMode,
  onToggleEditMode,
  isDirty,
  editedCount,
  saving,
  onSave,
  compareFileName,
  compareLoading,
  onPickCompareFile,
  onClearCompare,
}: ToolbarProps) {
  const compareInputRef = useRef<HTMLInputElement>(null)

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
            <span className="toolbar-brand-sub" aria-hidden="true">
              ·
            </span>
            <span className="toolbar-brand-sub">{fileName}</span>
          </>
        )}
      </div>
      <div className="toolbar-group">
        {fileName && (
          <>
            <button
              type="button"
              className={`pill pill--outline${editMode ? ' pill--accent' : ''}`}
              aria-pressed={editMode}
              onClick={onToggleEditMode}
            >
              {editMode ? '편집 중' : '편집 모드'}
            </button>
            {isDirty && (
              <button type="button" className="btn btn--primary" onClick={onSave} disabled={saving}>
                {saving ? '저장 중…' : `변경 ${editedCount}개 저장`}
              </button>
            )}

            {compareFileName ? (
              <span className="pill pill--outline pill--accent">
                {compareFileName}와 비교 중
                <button type="button" className="pill-close" onClick={onClearCompare} aria-label="비교 종료">
                  ×
                </button>
              </span>
            ) : (
              <button
                type="button"
                className="pill pill--outline"
                onClick={() => compareInputRef.current?.click()}
                disabled={compareLoading}
              >
                {compareLoading ? '불러오는 중…' : '버전 비교'}
              </button>
            )}
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
          </>
        )}
        <span className="pill pill--live">브라우저에서만 열림</span>
        {fileName && (
          <button className="btn btn--ghost" onClick={onOpenAnother} type="button">
            다른 파일 열기
          </button>
        )}
      </div>
    </div>
  )
}
