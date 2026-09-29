import { useRef, useState, type DragEvent } from 'react'
import tigerBody from '../assets/tiger/tiger-body.png'

interface DropZoneProps {
  onFile: (file: File) => void
  error: string | null
  loading: boolean
}

export function DropZone({ onFile, error, loading }: DropZoneProps) {
  const [active, setActive] = useState(false)
  const inputRef = useRef<HTMLInputElement>(null)

  const handleDrop = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault()
    setActive(false)
    const file = e.dataTransfer.files[0]
    if (file) onFile(file)
  }

  return (
    <div className="app-empty">
      <div
        className="dropzone"
        data-active={active}
        onDragOver={(e) => {
          e.preventDefault()
          setActive(true)
        }}
        onDragLeave={() => setActive(false)}
        onDrop={handleDrop}
      >
        <span className="tiger-stage">
          <img src={tigerBody} alt="" />
        </span>
        <div className="dropzone-title">여기에 엑셀 파일을 끌어다 놓으세요</div>
        <div className="dropzone-sub">xlsx · csv 지원 — 업로드 없이, 이 브라우저 안에서 바로 열려요</div>
        <button
          className="btn btn--primary"
          type="button"
          onClick={() => inputRef.current?.click()}
          disabled={loading}
        >
          {loading ? '읽는 중…' : '파일 선택하기'}
        </button>
        <input
          ref={inputRef}
          type="file"
          accept=".xlsx,.csv"
          style={{ display: 'none' }}
          onChange={(e) => {
            const file = e.target.files?.[0]
            if (file) onFile(file)
            e.target.value = ''
          }}
        />
        {error && <div className="dropzone-hint" style={{ color: 'var(--accent)' }}>{error}</div>}
        <div className="dropzone-hint">회원가입도, 로그인도 필요 없어요</div>
      </div>
    </div>
  )
}
